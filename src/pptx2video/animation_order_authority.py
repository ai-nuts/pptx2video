"""Resolve user preference when reading order conflicts with Animation Pane."""

from __future__ import annotations

import json
import math
import re
import shutil
import sys
import textwrap
from pathlib import Path


_ASCII_BASE_CANVAS_HEIGHT = 28
_ASCII_MIN_OUTER_WIDTH = 26
_ASCII_MAX_OUTER_WIDTH = 122
_ASCII_MIN_OUTER_HEIGHT = 18
_ASCII_MIN_RESPONSIVE_OUTER_HEIGHT = 10
_ASCII_CELL_WIDTH_TO_HEIGHT = 0.5
_DEFAULT_SLIDE_ASPECT_RATIO = 16.0 / 9.0
_SEMANTIC_PROFILES = {"concise", "detailed"}
_ASCII_LEGEND_GAP = 3
# Keep enough room for useful descriptions even though the visible title is
# deliberately short. The threshold must not shrink with the heading text.
_ASCII_CONCISE_LEGEND_MIN_WIDTH = 26
_ASCII_DETAILED_LEGEND_MIN_WIDTH = 27


def _slide_index(slide: dict[str, object]) -> int:
    try:
        return int(slide.get("index") or 0)
    except (TypeError, ValueError):
        return 0


def _conflict_slides(report: dict[str, object]) -> list[dict[str, object]]:
    """Return conflict pages in deterministic PowerPoint slide order."""
    return sorted(
        (
            slide
            for slide in report.get("slides") or []
            if isinstance(slide, dict) and bool(slide.get("conflicts"))
        ),
        key=lambda slide: (_slide_index(slide), str(slide.get("section_id") or "")),
    )


def _display_label(item: dict[str, object]) -> str:
    visible = [
        " ".join(str(value).split())
        for value in item.get("visible_texts") or []
        if " ".join(str(value).split())
    ]
    return " | ".join(visible) or str(
        item.get("handle") or item.get("shape_id") or "<unnamed>"
    )


def _normalize_semantic_profile(value: object) -> str:
    profile = str(value or "concise").strip().lower()
    if profile not in _SEMANTIC_PROFILES:
        raise ValueError(
            "semantic_profile must be one of: "
            + ", ".join(sorted(_SEMANTIC_PROFILES))
        )
    return profile


def _semantic_text(item: dict[str, object], semantic_profile: str) -> str:
    """Read an authored semantic description without runtime inference.

    Newer reports may carry a two-profile ``semantic_meaning`` mapping. The
    scalar aliases keep the ASCII renderer forward-compatible while older
    reports fall back deterministically to the canonical Notes semantic,
    narration, visible text, handle, or shape id.
    """
    profile = _normalize_semantic_profile(semantic_profile)
    for container_key in ("semantic_meaning", "semantic_meanings"):
        container = item.get(container_key)
        if not isinstance(container, dict):
            continue
        value = container.get(profile)
        if isinstance(value, dict):
            value = value.get("text")
        normalized = " ".join(str(value or "").split())
        if normalized:
            return normalized

    aliases = (
        ("detailed_semantic", "semantic_detailed", "detailed")
        if profile == "detailed"
        else ("concise_semantic", "semantic_concise", "concise")
    )
    for key in aliases:
        normalized = " ".join(str(item.get(key) or "").split())
        if normalized:
            return normalized

    if profile == "detailed":
        transcript = " ".join(str(item.get("transcript") or "").split())
        if transcript:
            return transcript
    semantic = " ".join(str(item.get("semantic") or "").split())
    return semantic or _display_label(item)


def _identity_token(index: int) -> str:
    """Return spreadsheet-style stable labels: A..Z, AA..AZ, BA..."""
    if index < 1:
        raise ValueError("identity index must be positive")
    token = ""
    value = index
    while value:
        value, remainder = divmod(value - 1, 26)
        token = chr(ord("A") + remainder) + token
    return token


def _sequence_items(
    slide: dict[str, object],
    key: str,
) -> list[dict[str, object]]:
    items = [item for item in slide.get(key) or [] if isinstance(item, dict)]

    def rank(item: dict[str, object]) -> tuple[int, str]:
        try:
            value = int(item.get("rank") or 0)
        except (TypeError, ValueError):
            value = 0
        return value, str(item.get("shape_id") or "")

    return sorted(items, key=rank)


def _bbox(item: dict[str, object]) -> tuple[float, float, float, float] | None:
    value = item.get("bbox")
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    try:
        x, y, width, height = (float(part) for part in value)
    except (TypeError, ValueError):
        return None
    if not all(math.isfinite(part) for part in (x, y, width, height)):
        return None
    if width <= 0.0 or height <= 0.0:
        return None
    return x, y, width, height


def _slide_bbox_by_shape(
    slide: dict[str, object],
) -> dict[str, tuple[float, float, float, float]]:
    """Collect geometry from both current and backward-compatible reports."""
    boxes: dict[str, tuple[float, float, float, float]] = {}
    sources: list[dict[str, object]] = []
    sources.extend(_sequence_items(slide, "geometry_sequence"))
    sources.extend(_sequence_items(slide, "reading_sequence"))
    sources.extend(_sequence_items(slide, "animation_pane_sequence"))
    for conflict in slide.get("conflicts") or []:
        if not isinstance(conflict, dict):
            continue
        sources.append(conflict)
        sources.extend(
            item
            for item in conflict.get("inversions") or []
            if isinstance(item, dict)
        )
    for item in sources:
        shape_id = str(item.get("shape_id") or "")
        box = _bbox(item)
        if shape_id and box is not None:
            boxes.setdefault(shape_id, box)
    return boxes


def _slide_aspect_ratio(slide: dict[str, object]) -> float:
    """Return the real PPTX canvas aspect ratio or a legacy-report fallback."""
    size = slide.get("slide_size_emu")
    if isinstance(size, (list, tuple)) and len(size) == 2:
        try:
            width = float(size[0])
            height = float(size[1])
        except (TypeError, ValueError):
            pass
        else:
            if (
                math.isfinite(width)
                and math.isfinite(height)
                and width > 0
                and height > 0
            ):
                return width / height
    try:
        ratio = float(slide.get("canvas_aspect_ratio") or 0.0)
    except (TypeError, ValueError):
        ratio = 0.0
    if math.isfinite(ratio) and ratio > 0.0:
        return ratio
    return _DEFAULT_SLIDE_ASPECT_RATIO


def _ascii_canvas_dimensions(
    slide: dict[str, object],
    *,
    max_outer_width: int | None = None,
) -> tuple[int, int]:
    """Approximate the physical slide ratio in a monospace character grid.

    A typical code-block character is about half as wide as it is tall. The
    outer frame is included in the calculation because it is part of the
    displayed canvas. Width is capped for practical terminals; exceptionally
    wide slides reduce row count before accepting aspect-ratio error.
    """
    aspect_ratio = _slide_aspect_ratio(slide)
    outer_height = _ASCII_BASE_CANVAS_HEIGHT + 2
    outer_width = int(
        round(outer_height * aspect_ratio / _ASCII_CELL_WIDTH_TO_HEIGHT)
    )
    width_limit = _ASCII_MAX_OUTER_WIDTH
    responsive = max_outer_width is not None
    if responsive:
        width_limit = min(width_limit, max(_ASCII_MIN_OUTER_WIDTH, max_outer_width))
    if outer_width > width_limit:
        minimum_height = (
            _ASCII_MIN_RESPONSIVE_OUTER_HEIGHT
            if responsive
            else _ASCII_MIN_OUTER_HEIGHT
        )
        candidate_widths = range(
            max(_ASCII_MIN_OUTER_WIDTH, width_limit - 8),
            width_limit + 1,
        )
        candidates: list[tuple[float, int, int]] = []
        for candidate_width in candidate_widths:
            candidate_height = max(
                minimum_height,
                int(
                    round(
                        candidate_width
                        * _ASCII_CELL_WIDTH_TO_HEIGHT
                        / aspect_ratio
                    )
                ),
            )
            displayed_ratio = (
                candidate_width
                * _ASCII_CELL_WIDTH_TO_HEIGHT
                / candidate_height
            )
            relative_error = abs(displayed_ratio - aspect_ratio) / aspect_ratio
            candidates.append(
                (relative_error, candidate_width, candidate_height)
            )
        acceptable = [item for item in candidates if item[0] <= 0.01]
        if acceptable:
            _error, outer_width, outer_height = max(
                acceptable,
                key=lambda item: (item[1], item[2]),
            )
        else:
            _error, outer_width, outer_height = min(
                candidates,
                key=lambda item: (item[0], -item[1]),
            )
    elif outer_width < _ASCII_MIN_OUTER_WIDTH:
        outer_width = _ASCII_MIN_OUTER_WIDTH
    return outer_width - 2, outer_height - 2


def _expand_frame_interval(
    start: int,
    end: int,
    *,
    minimum_span: int,
    maximum: int,
) -> tuple[int, int]:
    """Expand a raster interval around its center without leaving the canvas."""
    if start > end:
        start, end = end, start
    missing = max(0, minimum_span - (end - start))
    start -= missing // 2
    end += missing - missing // 2
    if start < 0:
        end = min(maximum, end - start)
        start = 0
    if end > maximum:
        start = max(0, start - (end - maximum))
        end = maximum
    return start, end


def _ascii_frame_bounds(
    box: tuple[float, float, float, float],
    *,
    canvas_width: int,
    canvas_height: int,
) -> tuple[int, int, int, int]:
    """Scale one normalized OOXML bbox to an inclusive ASCII frame."""
    x, y, width, height = box
    left = int(round(min(1.0, max(0.0, x)) * (canvas_width - 1)))
    right = int(round(min(1.0, max(0.0, x + width)) * (canvas_width - 1)))
    top = int(round(min(1.0, max(0.0, y)) * (canvas_height - 1)))
    bottom = int(round(min(1.0, max(0.0, y + height)) * (canvas_height - 1)))
    left, right = _expand_frame_interval(
        left,
        right,
        minimum_span=3,
        maximum=canvas_width - 1,
    )
    # Three total rows are the minimum needed for a framed one-line element.
    top, bottom = _expand_frame_interval(
        top,
        bottom,
        minimum_span=2,
        maximum=canvas_height - 1,
    )
    return left, right, top, bottom


def _merge_frame_character(current: str, incoming: str) -> str:
    if current == " ":
        return incoming
    if current == incoming:
        return current
    return "+"


def _draw_ascii_frame(
    canvas: list[list[str]],
    bounds: tuple[int, int, int, int],
) -> None:
    left, right, top, bottom = bounds
    for column in range(left, right + 1):
        incoming = "+" if column in {left, right} else "-"
        canvas[top][column] = _merge_frame_character(canvas[top][column], incoming)
        canvas[bottom][column] = _merge_frame_character(
            canvas[bottom][column], incoming
        )
    for row in range(top + 1, bottom):
        canvas[row][left] = _merge_frame_character(canvas[row][left], "|")
        canvas[row][right] = _merge_frame_character(canvas[row][right], "|")


def _place_ascii_frame_label(
    canvas: list[list[str]],
    bounds: tuple[int, int, int, int],
    token: str,
) -> bool:
    left, right, top, bottom = bounds
    interior_width = right - left - 1
    if interior_width <= 0 or bottom - top <= 1:
        return False
    if len(token) > interior_width:
        return False
    center_row = (top + bottom) // 2
    rows = sorted(range(top + 1, bottom), key=lambda row: (abs(row - center_row), row))
    center_column = (left + right) // 2
    preferred_column = center_column - len(token) // 2
    columns = sorted(
        range(left + 1, right - len(token) + 1),
        key=lambda column: (abs(column - preferred_column), column),
    )
    for row in rows:
        for column in columns:
            if any(
                canvas[row][index] != " "
                for index in range(column, column + len(token))
            ):
                continue
            canvas[row][column : column + len(token)] = list(token)
            return True
    return False


def _element_key(item: dict[str, object]) -> str:
    shape_id = str(item.get("shape_id") or "").strip()
    if shape_id:
        return shape_id
    handle = str(item.get("handle") or "").strip()
    return f"handle:{handle}" if handle else ""


def _ascii_token_canvas(
    items: list[dict[str, object]],
    token_by_shape: dict[str, str],
    boxes: dict[str, tuple[float, float, float, float]],
    *,
    canvas_width: int,
    canvas_height: int,
) -> tuple[list[str], list[str]]:
    """Draw one bbox canvas using caller-supplied rank or identity tokens."""
    canvas = [[" " for _ in range(canvas_width)] for _ in range(canvas_height)]
    unplaced: list[str] = []
    frames: list[tuple[int, str, str, int, tuple[int, int, int, int]]] = []

    for item_index, item in enumerate(items):
        shape_id = str(item.get("shape_id") or "")
        token = str(token_by_shape.get(_element_key(item)) or "?")
        box = boxes.get(shape_id)
        if box is None:
            unplaced.append(f"{token} [shape {shape_id}]")
            continue
        bounds = _ascii_frame_bounds(
            box,
            canvas_width=canvas_width,
            canvas_height=canvas_height,
        )
        left, right, top, bottom = bounds
        frames.append(
            (
                (right - left) * (bottom - top),
                shape_id,
                token,
                item_index,
                bounds,
            )
        )

    # Draw large containers first so smaller overlapping elements stay legible.
    for _area, _shape_id, _token, _item_index, bounds in sorted(
        frames, key=lambda item: (-item[0], item[4], item[1])
    ):
        _draw_ascii_frame(canvas, bounds)
    for _area, shape_id, token, item_index, bounds in sorted(
        frames, key=lambda item: (item[3], item[1])
    ):
        if not _place_ascii_frame_label(canvas, bounds, token):
            unplaced.append(f"{token} [shape {shape_id}]")

    return [
        "+" + "-" * canvas_width + "+",
        *("|" + "".join(row) + "|" for row in canvas),
        "+" + "-" * canvas_width + "+",
    ], unplaced


def _ascii_rank_canvas(
    items: list[dict[str, object]],
    rank_by_shape: dict[str, int],
    boxes: dict[str, tuple[float, float, float, float]],
    *,
    canvas_width: int,
    canvas_height: int,
) -> tuple[list[str], list[str]]:
    """Draw one bbox canvas whose frame labels contain only rank numbers."""
    return _ascii_token_canvas(
        items,
        {
            _element_key(item): (
                str(rank_by_shape.get(_element_key(item)) or "?")
            )
            for item in items
        },
        boxes,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
    )


def _merge_identity_item(
    existing: dict[str, object],
    candidate: dict[str, object],
) -> None:
    """Backfill optional semantic fields while preserving geometry identity."""
    for key, value in candidate.items():
        if key not in existing or existing.get(key) in (None, "", [], {}):
            existing[key] = value


def _identity_items(
    slide: dict[str, object],
    reading: list[dict[str, object]],
    pane: list[dict[str, object]],
) -> list[dict[str, object]]:
    """Collect elements and sort them independently of either candidate order."""
    ordered: list[dict[str, object]] = []
    by_identity: dict[str, dict[str, object]] = {}
    supplemental: list[dict[str, object]] = []
    for conflict in slide.get("conflicts") or []:
        if not isinstance(conflict, dict):
            continue
        supplemental.append(conflict)
        supplemental.extend(
            item
            for item in conflict.get("inversions") or []
            if isinstance(item, dict)
        )
    for item in [*reading, *pane, *supplemental]:
        element_key = _element_key(item)
        if not element_key:
            continue
        if element_key in by_identity:
            _merge_identity_item(by_identity[element_key], item)
            continue
        copied = dict(item)
        by_identity[element_key] = copied
        ordered.append(copied)

    def natural_parts(value: str) -> tuple[tuple[int, object], ...]:
        parts: list[tuple[int, object]] = []
        cursor = 0
        for match in re.finditer(r"\d+", value):
            if match.start() > cursor:
                parts.append((1, value[cursor : match.start()].casefold()))
            parts.append((0, int(match.group())))
            cursor = match.end()
        if cursor < len(value):
            parts.append((1, value[cursor:].casefold()))
        return tuple(parts)

    def identity_key(item: dict[str, object]) -> tuple[object, ...]:
        shape_id = str(item.get("shape_id") or "").strip()
        handle = str(item.get("handle") or "").strip()
        if shape_id:
            return (0, natural_parts(shape_id), handle.casefold())
        return (1, natural_parts(handle), _display_label(item).casefold())

    return sorted(ordered, key=identity_key)


def _identity_labels(items: list[dict[str, object]]) -> dict[str, str]:
    return {
        _element_key(item): _identity_token(index)
        for index, item in enumerate(items, start=1)
    }


def _semantic_legend_lines(
    items: list[dict[str, object]],
    token_by_shape: dict[str, object],
    *,
    title: str,
    semantic_profile: str,
    width: int,
) -> list[str]:
    profile = _normalize_semantic_profile(semantic_profile)
    effective_width = max(8, int(width))
    lines = [title]
    for item in items:
        token = token_by_shape.get(_element_key(item)) or "?"
        prefix = f"{token} = "
        wrapped = textwrap.wrap(
            _semantic_text(item, profile),
            width=effective_width,
            initial_indent=prefix,
            subsequent_indent=" " * len(prefix),
            break_long_words=True,
            break_on_hyphens=False,
        )
        lines.extend(wrapped or [prefix.rstrip()])
    return lines


def _sequence_legend_items(
    sequence: list[dict[str, object]],
    identity_items: list[dict[str, object]],
) -> list[dict[str, object]]:
    """Return sequence-ordered items with the richest available semantics."""
    canonical = {_element_key(item): item for item in identity_items}
    result: list[dict[str, object]] = []
    for item in sequence:
        copied = dict(item)
        richer = canonical.get(_element_key(item))
        if richer is not None:
            _merge_identity_item(copied, richer)
        result.append(copied)
    return result


def _side_by_side_lines(
    left: list[str],
    right: list[str],
    *,
    left_width: int,
) -> list[str]:
    line_count = max(len(left), len(right))
    lines: list[str] = []
    for index in range(line_count):
        left_line = left[index] if index < len(left) else ""
        right_line = right[index] if index < len(right) else ""
        if right_line:
            lines.append(
                left_line.ljust(left_width)
                + " " * _ASCII_LEGEND_GAP
                + right_line
            )
        else:
            lines.append(left_line)
    return lines


def _responsive_canvas_layout(
    *,
    max_output_width: int | None,
    semantic_profile: str,
) -> tuple[int | None, int, bool]:
    """Return canvas width budget, legend width, and side-by-side mode."""
    profile = _normalize_semantic_profile(semantic_profile)
    total_width = (
        _ASCII_MAX_OUTER_WIDTH
        if max_output_width is None
        else max(1, int(max_output_width))
    )
    legend_minimum = (
        _ASCII_DETAILED_LEGEND_MIN_WIDTH
        if profile == "detailed"
        else _ASCII_CONCISE_LEGEND_MIN_WIDTH
    )
    side_by_side = (
        total_width
        >= _ASCII_MIN_OUTER_WIDTH + _ASCII_LEGEND_GAP + legend_minimum
    )
    if not side_by_side:
        return max_output_width, total_width, False
    legend_width = max(legend_minimum, min(48, total_width // 3))
    canvas_budget = total_width - _ASCII_LEGEND_GAP - legend_width
    return canvas_budget, legend_width, True


def _ascii_relative_layout(
    slide: dict[str, object],
    *,
    max_output_width: int | None = None,
    semantic_profile: str = "concise",
) -> list[str]:
    """Draw three responsive bbox views without image or LLM input."""
    profile = _normalize_semantic_profile(semantic_profile)
    reading = _sequence_items(slide, "geometry_sequence") or _sequence_items(
        slide, "reading_sequence"
    )
    pane = _sequence_items(slide, "animation_pane_sequence")
    reading_rank = {
        _element_key(item): index
        for index, item in enumerate(reading, start=1)
    }
    pane_rank = {
        _element_key(item): index
        for index, item in enumerate(pane, start=1)
    }
    items = _identity_items(slide, reading, pane)
    identity_by_shape = _identity_labels(items)
    reading_legend_items = _sequence_legend_items(reading, items)
    pane_legend_items = _sequence_legend_items(pane, items)

    boxes = _slide_bbox_by_shape(slide)
    if max_output_width is not None and max_output_width < _ASCII_MIN_OUTER_WIDTH:
        return [
            "Relative layout omitted: terminal is narrower than "
            f"{_ASCII_MIN_OUTER_WIDTH} columns.",
            _sequence_line("Geometry reading order (text)", reading),
            *_semantic_legend_lines(
                reading_legend_items,
                reading_rank,
                title="LEGEND",
                semantic_profile=profile,
                width=max_output_width,
            ),
            _sequence_line("Animation Pane order (text)", pane),
            *_semantic_legend_lines(
                pane_legend_items,
                pane_rank,
                title="LEGEND",
                semantic_profile=profile,
                width=max_output_width,
            ),
            "IDENTITY MAP (letters identify elements, not order)",
            *_semantic_legend_lines(
                items,
                identity_by_shape,
                title="LEGEND",
                semantic_profile=profile,
                width=max_output_width,
            ),
        ]
    # Monospace characters are taller than they are wide. Width and height are
    # fitted together so compact terminals retain the slide's physical ratio.
    canvas_width_budget, legend_width, side_by_side = _responsive_canvas_layout(
        max_output_width=max_output_width,
        semantic_profile=profile,
    )
    canvas_width, canvas_height = _ascii_canvas_dimensions(
        slide,
        max_outer_width=canvas_width_budget,
    )
    reading_canvas, reading_unplaced = _ascii_rank_canvas(
        items,
        reading_rank,
        boxes,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
    )
    pane_canvas, pane_unplaced = _ascii_rank_canvas(
        items,
        pane_rank,
        boxes,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
    )
    identity_canvas, identity_unplaced = _ascii_token_canvas(
        items,
        identity_by_shape,
        boxes,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
    )
    reading_legend = _semantic_legend_lines(
        reading_legend_items,
        reading_rank,
        title="LEGEND",
        semantic_profile=profile,
        width=legend_width,
    )
    pane_legend = _semantic_legend_lines(
        pane_legend_items,
        pane_rank,
        title="LEGEND",
        semantic_profile=profile,
        width=legend_width,
    )
    identity_legend = _semantic_legend_lines(
        items,
        identity_by_shape,
        title="LEGEND",
        semantic_profile=profile,
        width=legend_width,
    )
    if side_by_side:
        reading_view = _side_by_side_lines(
            reading_canvas,
            reading_legend,
            left_width=canvas_width + 2,
        )
        pane_view = _side_by_side_lines(
            pane_canvas,
            pane_legend,
            left_width=canvas_width + 2,
        )
        identity_view = _side_by_side_lines(
            identity_canvas,
            identity_legend,
            left_width=canvas_width + 2,
        )
    else:
        reading_view = [*reading_canvas, *reading_legend]
        pane_view = [*pane_canvas, *pane_legend]
        identity_view = [*identity_canvas, *identity_legend]
    lines = [
        "Relative layout (same OOXML bbox frames in all three views):",
        "GEOMETRY READING ORDER",
        *reading_view,
        "",
        "ANIMATION PANE ORDER",
        *pane_view,
        "",
        "IDENTITY MAP (letters identify elements, not order)",
        *identity_view,
    ]
    if reading_unplaced:
        lines.append(
            "  Geometry bbox/label unavailable or overlapping: "
            + ", ".join(reading_unplaced)
        )
    if pane_unplaced:
        lines.append(
            "  Pane bbox/label unavailable or overlapping: "
            + ", ".join(pane_unplaced)
        )
    if identity_unplaced:
        lines.append(
            "  Identity bbox/label unavailable or overlapping: "
            + ", ".join(identity_unplaced)
        )
    return lines


def _sequence_line(title: str, items: list[dict[str, object]]) -> str:
    labels = [
        f"{index}. {_display_label(item)}"
        for index, item in enumerate(items, start=1)
    ]
    return f"{title}: " + (" -> ".join(labels) if labels else "<empty>")


def slide_order_conflict_summary(
    slide: dict[str, object],
    *,
    max_output_width: int | None = None,
    semantic_profile: str = "concise",
) -> str:
    """Build one self-contained conflict-page explanation."""
    reading = _sequence_items(slide, "geometry_sequence") or _sequence_items(
        slide, "reading_sequence"
    )
    pane = _sequence_items(slide, "animation_pane_sequence")
    lines = [
        f"Slide {_slide_index(slide)} order conflict",
        *_ascii_relative_layout(
            slide,
            max_output_width=max_output_width,
            semantic_profile=semantic_profile,
        ),
    ]
    lines.append(_sequence_line("Geometry reading order", reading))
    lines.append(_sequence_line("Animation Pane order", pane))
    for conflict in sorted(
        (
            item
            for item in slide.get("conflicts") or []
            if isinstance(item, dict)
        ),
        key=lambda item: (
            int(item.get("reading_rank") or 0),
            str(item.get("shape_id") or ""),
        ),
    ):
        lines.append(
            f"  Conflict: {_display_label(conflict)!r} "
            f"(shape {conflict.get('shape_id')}, geometry "
            f"{conflict.get('geometry_rank') or conflict.get('reading_rank')}, "
            f"Pane {conflict.get('pane_rank')}, "
            f"native row "
            f"{conflict.get('native_order')})"
        )
        crossed = ", ".join(
            str(item.get("handle") or item.get("shape_id") or "")
            for item in conflict.get("inversions") or []
            if isinstance(item, dict)
        )
        if crossed:
            lines.append(f"    relative order is reversed against: {crossed}")
        move = conflict.get("recommended_pane_move") or {}
        if not isinstance(move, dict) or not move:
            continue
        after = move.get("place_after") or {}
        before = move.get("place_before") or {}
        if not isinstance(after, dict):
            after = {}
        if not isinstance(before, dict):
            before = {}
        lines.append(
            f"    To match geometry: move Pane row "
            f"{move.get('current_native_order')} after row "
            f"{after.get('native_order') or '<start>'} "
            f"({after.get('handle') or '<start>'}), before row "
            f"{before.get('native_order') or '<end>'} "
            f"({before.get('handle') or '<end>'})"
        )
    return "\n".join(lines)


def order_conflict_summary(
    report: dict[str, object],
    *,
    max_output_width: int | None = None,
    semantic_profile: str = "concise",
) -> str:
    """Build deterministic ASCII explanations for conflict pages only."""
    slides = _conflict_slides(report)
    if not slides:
        return ""
    return "\n\n".join(
        slide_order_conflict_summary(
            slide,
            max_output_width=max_output_width,
            semantic_profile=semantic_profile,
        )
        for slide in slides
    )


def _slide_decisions(
    slides: list[dict[str, object]],
    policy: str,
    source: str,
) -> list[dict[str, object]]:
    return [
        {
            "slide_index": _slide_index(slide),
            "section_id": str(slide.get("section_id") or ""),
            "selected_policy": policy,
            "selection_source": source,
        }
        for slide in slides
    ]


def _identity_order_contract(
    slide: dict[str, object],
) -> dict[str, object]:
    """Return the stable per-slide Identity Map and both reference orders."""
    geometry = _sequence_items(slide, "geometry_sequence") or _sequence_items(
        slide, "reading_sequence"
    )
    pane = _sequence_items(slide, "animation_pane_sequence")
    items = _identity_items(slide, geometry, pane)
    token_by_element = _identity_labels(items)
    rows: list[dict[str, object]] = []
    for item in items:
        element_key = _element_key(item)
        rows.append(
            {
                "identity": token_by_element[element_key],
                "element_key": element_key,
                "shape_id": str(item.get("shape_id") or ""),
                "handle": str(item.get("handle") or ""),
            }
        )
    token_by_key = {
        str(row["element_key"]): str(row["identity"]) for row in rows
    }

    def sequence_tokens(sequence: list[dict[str, object]]) -> list[str]:
        return [
            token_by_key[_element_key(item)]
            for item in sequence
            if _element_key(item) in token_by_key
        ]

    return {
        "rows": rows,
        "tokens": [str(row["identity"]) for row in rows],
        "geometry_tokens": sequence_tokens(geometry),
        "pane_tokens": sequence_tokens(pane),
    }


def _canonical_identity_order(tokens: list[str]) -> str:
    return (
        "".join(tokens)
        if all(len(token) == 1 for token in tokens)
        else ",".join(tokens)
    )


def _parse_identity_order(
    slide: dict[str, object],
    raw_order: object,
) -> dict[str, object]:
    """Parse a complete order-neutral Identity permutation without an LLM."""
    contract = _identity_order_contract(slide)
    expected = [str(token) for token in contract["tokens"]]
    if not expected:
        raise ValueError("this slide has no Identity Map elements")
    normalized = str(raw_order or "").strip().upper()
    if not normalized:
        raise ValueError("identity order is empty")
    if re.search(r"[^A-Z,;:\s>\-→]", normalized):
        raise ValueError(
            "identity order may contain only A-Z letters or separators"
        )

    compact = re.fullmatch(r"[A-Z]+", normalized) is not None
    single_letter_contract = all(len(token) == 1 for token in expected)
    if compact:
        if not single_letter_contract:
            raise ValueError(
                "compact identity order is ambiguous after Z; separate tokens "
                "such as A,B,...,AA"
            )
        tokens = list(normalized)
    else:
        groups = re.findall(r"[A-Z]+", normalized)
        tokens = (
            [letter for group in groups for letter in group]
            if single_letter_contract
            else groups
        )
    if not tokens:
        raise ValueError("identity order contains no element labels")

    expected_set = set(expected)
    unknown = sorted({token for token in tokens if token not in expected_set})
    duplicates = sorted(
        token for token in expected_set if tokens.count(token) > 1
    )
    missing = [token for token in expected if token not in tokens]
    problems: list[str] = []
    if unknown:
        problems.append("unknown " + ", ".join(unknown))
    if duplicates:
        problems.append("duplicate " + ", ".join(duplicates))
    if missing:
        problems.append("missing " + ", ".join(missing))
    if len(tokens) != len(expected) and not problems:
        problems.append(
            f"expected {len(expected)} identities but received {len(tokens)}"
        )
    if problems:
        raise ValueError("invalid identity order: " + "; ".join(problems))

    rows_by_token = {
        str(row["identity"]): row for row in contract["rows"]
    }
    requested_rows = [rows_by_token[token] for token in tokens]
    pane_tokens = [str(token) for token in contract["pane_tokens"]]
    geometry_tokens = [str(token) for token in contract["geometry_tokens"]]
    if tokens == pane_tokens:
        matched_policy = "animation-pane"
    elif tokens == geometry_tokens:
        matched_policy = "reading-order"
    else:
        matched_policy = "custom-order"
    return {
        "identity_order": _canonical_identity_order(tokens),
        "identity_tokens": tokens,
        "shape_id_order": [str(row["shape_id"]) for row in requested_rows],
        "handle_order": [str(row["handle"]) for row in requested_rows],
        "pane_identity_order": _canonical_identity_order(pane_tokens),
        "geometry_identity_order": _canonical_identity_order(geometry_tokens),
        "matched_policy": matched_policy,
    }


def _identity_order_decision(
    slide: dict[str, object],
    raw_order: object,
    *,
    source: str,
) -> dict[str, object]:
    parsed = _parse_identity_order(slide, raw_order)
    return {
        "slide_index": _slide_index(slide),
        "section_id": str(slide.get("section_id") or ""),
        "selected_policy": parsed["matched_policy"],
        "selection_source": source,
        "input_protocol": "identity_order.v1",
        **parsed,
    }


def _identity_order_specs(
    raw_specs: list[str],
    slides: list[dict[str, object]],
) -> dict[int, str]:
    """Resolve repeatable CLI values such as CABGDEF or 2=CABGDEF."""
    if not raw_specs:
        return {}
    slide_indices = {_slide_index(slide) for slide in slides}
    result: dict[int, str] = {}
    unscoped: list[str] = []
    for raw_spec in raw_specs:
        spec = str(raw_spec or "").strip()
        match = re.fullmatch(r"(?:slide\s*)?(\d+)\s*[:=]\s*(.+)", spec, re.I)
        if match is None:
            unscoped.append(spec)
            continue
        slide_index = int(match.group(1))
        if slide_index not in slide_indices:
            raise ValueError(
                f"identity order references non-conflict slide {slide_index}"
            )
        if slide_index in result:
            raise ValueError(
                f"identity order for slide {slide_index} was supplied more than once"
            )
        result[slide_index] = match.group(2).strip()
    if unscoped:
        if len(unscoped) > 1:
            raise ValueError("more than one unscoped identity order was supplied")
        if len(slides) != 1:
            raise ValueError(
                "unscoped identity order requires exactly one conflict slide; "
                "use SLIDE=ORDER for multiple pages"
            )
        slide_index = _slide_index(slides[0])
        if slide_index in result:
            raise ValueError(
                f"identity order for slide {slide_index} was supplied more than once"
            )
        result[slide_index] = unscoped[0]
    return result


def _finish_slide_decisions(
    resolved: dict[str, object],
    decisions: list[dict[str, object]],
    *,
    report_path: Path,
    selection_source: str,
) -> tuple[str | None, dict[str, object]]:
    revision_policies = {
        str(decision.get("selected_policy") or "")
        for decision in decisions
        if str(decision.get("selected_policy") or "") != "animation-pane"
    }
    if revision_policies:
        selected_policy = (
            "custom-order"
            if "custom-order" in revision_policies
            else "reading-order"
        )
        resolved.update(
            {
                "status": "pptx_revision_required",
                "selected_policy": selected_policy,
                "selection_source": selection_source,
                "slide_decisions": decisions,
                "recommended_action": (
                    "Reorder the named native Animation Pane rows to each "
                    "decision's identity_order, then rerun with "
                    "--animation-order-policy auto."
                ),
            }
        )
        _write_json(report_path, resolved)
        print(
            "[render_edited_pptx] at least one selected Identity order differs "
            "from the source Animation Pane. Update the PPTX first; no output "
            "bundle was created.\n"
            f"[render_edited_pptx] decision report: {report_path}",
            file=sys.stderr,
        )
        return None, resolved

    resolved.update(
        {
            "status": "resolved",
            "selected_policy": "animation-pane",
            "selection_source": selection_source,
            "slide_decisions": decisions,
        }
    )
    return "animation-pane", resolved


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def select_animation_order_policy(
    requested: str,
    report: dict[str, object],
    *,
    report_path: Path,
    semantic_profile: str = "concise",
    identity_order_inputs: list[str] | None = None,
) -> tuple[str | None, dict[str, object]]:
    """Resolve conflict pages in slide order, or stop after recording choices."""
    profile = _normalize_semantic_profile(semantic_profile)
    resolved = dict(report)
    resolved["semantic_profile"] = profile
    raw_identity_orders = list(identity_order_inputs or [])
    if int(report.get("conflict_count") or 0) == 0:
        if raw_identity_orders:
            resolved.update(
                {
                    "status": "invalid_identity_order",
                    "selected_policy": None,
                    "selection_source": "command_line_identity_sequence",
                    "slide_decisions": [],
                    "identity_order_error": (
                        "identity order was supplied but the deck has no conflict pages"
                    ),
                }
            )
            _write_json(report_path, resolved)
            print(
                "[render_edited_pptx] identity order was supplied but the deck has "
                "no conflict pages.",
                file=sys.stderr,
            )
            return None, resolved
        resolved.update(
            {
                "status": "no_conflict",
                "selected_policy": "animation-pane",
                "selection_source": "no_conflict_default",
                "slide_decisions": [],
            }
        )
        return "animation-pane", resolved

    slides = _conflict_slides(report)
    terminal_width = shutil.get_terminal_size(fallback=(80, 24)).columns
    summary = order_conflict_summary(
        report,
        max_output_width=terminal_width,
        semantic_profile=profile,
    )
    if raw_identity_orders and requested != "auto":
        resolved.update(
            {
                "status": "invalid_identity_order",
                "selected_policy": None,
                "selection_source": "command_line_identity_sequence",
                "slide_decisions": [],
                "identity_order_error": (
                    "--animation-order-sequence requires "
                    "--animation-order-policy auto"
                ),
            }
        )
        _write_json(report_path, resolved)
        print(summary, file=sys.stderr)
        print(
            "[render_edited_pptx] --animation-order-sequence requires "
            "--animation-order-policy auto.",
            file=sys.stderr,
        )
        return None, resolved
    if raw_identity_orders:
        try:
            supplied = _identity_order_specs(raw_identity_orders, slides)
            missing_slides = [
                _slide_index(slide)
                for slide in slides
                if _slide_index(slide) not in supplied
            ]
            if missing_slides:
                raise ValueError(
                    "missing identity order for conflict slide(s) "
                    + ", ".join(str(index) for index in missing_slides)
                )
            decisions = [
                _identity_order_decision(
                    slide,
                    supplied[_slide_index(slide)],
                    source="command_line_identity_sequence",
                )
                for slide in slides
            ]
        except ValueError as exc:
            resolved.update(
                {
                    "status": "invalid_identity_order",
                    "selected_policy": None,
                    "selection_source": "command_line_identity_sequence",
                    "slide_decisions": [],
                    "identity_order_error": str(exc),
                }
            )
            _write_json(report_path, resolved)
            print(summary, file=sys.stderr)
            print(
                f"[render_edited_pptx] invalid identity order: {exc}\n"
                f"[render_edited_pptx] decision report: {report_path}",
                file=sys.stderr,
            )
            return None, resolved
        print(summary, file=sys.stderr)
        return _finish_slide_decisions(
            resolved,
            decisions,
            report_path=report_path,
            selection_source="command_line_identity_sequence",
        )
    if requested == "animation-pane":
        resolved.update(
            {
                "status": "resolved",
                "selected_policy": "animation-pane",
                "selection_source": "command_line",
                "slide_decisions": _slide_decisions(
                    slides,
                    "animation-pane",
                    "command_line",
                ),
            }
        )
        print(summary, file=sys.stderr)
        print(
            "[render_edited_pptx] Animation Pane order selected for all conflict "
            "slides.",
            file=sys.stderr,
        )
        return "animation-pane", resolved
    if requested == "reading-order":
        resolved.update(
            {
                "status": "pptx_revision_required",
                "selected_policy": "reading-order",
                "selection_source": "command_line",
                "slide_decisions": _slide_decisions(
                    slides,
                    "reading-order",
                    "command_line",
                ),
                "recommended_action": (
                    "Move the named native animation rows to the reported reading-order "
                    "placement in PowerPoint, then rerun with "
                    "--animation-order-policy auto."
                ),
            }
        )
        _write_json(report_path, resolved)
        print(summary, file=sys.stderr)
        print(
            "[render_edited_pptx] reading order was selected. The source PPTX must be "
            "updated so its Animation Pane matches; no output bundle was created.\n"
            f"[render_edited_pptx] decision report: {report_path}",
            file=sys.stderr,
        )
        return None, resolved

    if not sys.stdin.isatty():
        resolved["slide_decisions"] = []
        _write_json(report_path, resolved)
        print(summary, file=sys.stderr)
        print(
            "[render_edited_pptx] animation order needs a user decision. Rerun with "
            "--animation-order-policy animation-pane to confirm the Pane order, or "
            "--animation-order-policy reading-order to request the reported PPTX "
            "revision.\n"
            f"[render_edited_pptx] decision report: {report_path}",
            file=sys.stderr,
        )
        return None, resolved

    decisions: list[dict[str, object]] = []
    for position, slide in enumerate(slides, start=1):
        print(
            f"\nConflict page {position}/{len(slides)}",
            file=sys.stderr,
        )
        print(
            slide_order_conflict_summary(
                slide,
                max_output_width=terminal_width,
                semantic_profile=profile,
            ),
            file=sys.stderr,
        )
        print("Choose this slide's order:", file=sys.stderr)
        print("  1. Keep Animation Pane order and render", file=sys.stderr)
        print(
            "  2. Use geometry reading order; revise the reported Pane rows",
            file=sys.stderr,
        )
        example = _canonical_identity_order(
            [
                str(token)
                for token in _identity_order_contract(slide)["pane_tokens"]
            ]
        )
        print(
            f"  Type a complete Identity order, for example {example}",
            file=sys.stderr,
        )
        print("  3. Cancel", file=sys.stderr)
        while True:
            try:
                answer = input(
                    f"Slide {_slide_index(slide)} selection [1/2/ORDER/3]: "
                ).strip().lower()
            except EOFError:
                resolved.update(
                    {
                        "status": "cancelled",
                        "selected_policy": None,
                        "selection_source": "interactive_eof",
                        "slide_decisions": decisions,
                    }
                )
                _write_json(report_path, resolved)
                print(
                    "\n[render_edited_pptx] input closed; rendering cancelled and "
                    "no output bundle was created.\n"
                    f"[render_edited_pptx] decision report: {report_path}",
                    file=sys.stderr,
                )
                return None, resolved
            if answer in {"1", "pane", "animation-pane"}:
                decisions.extend(
                    _slide_decisions([slide], "animation-pane", "interactive")
                )
                break
            if answer in {"2", "geometry", "reading", "reading-order"}:
                decisions.extend(
                    _slide_decisions([slide], "reading-order", "interactive")
                )
                break
            if answer in {"3", "cancel", "q", "quit"}:
                resolved.update(
                    {
                        "status": "cancelled",
                        "selected_policy": None,
                        "selection_source": "interactive",
                        "slide_decisions": decisions,
                    }
                )
                _write_json(report_path, resolved)
                print(
                    "[render_edited_pptx] rendering cancelled; no output bundle "
                    "was created.\n"
                    f"[render_edited_pptx] decision report: {report_path}",
                    file=sys.stderr,
                )
                return None, resolved
            try:
                decisions.append(
                    _identity_order_decision(
                        slide,
                        answer,
                        source="interactive_identity_sequence",
                    )
                )
            except ValueError as exc:
                print(
                    f"Enter 1, 2, 3, or a complete Identity order: {exc}",
                    file=sys.stderr,
                )
                continue
            break

    return _finish_slide_decisions(
        resolved,
        decisions,
        report_path=report_path,
        selection_source="interactive",
    )
