#!/usr/bin/env python3
"""Deterministic PowerPoint Author Notes and Alt Text protocol utilities.

This module treats the PPTX as the editable source. It reads native PowerPoint
entrance and emphasis rows, reconciles their targets with authoritative Author
Notes, generates compact Alt Text, and produces narration and animation metadata
without an LLM or a ppt-master project.
"""

from __future__ import annotations

import hashlib
import json
import math
import posixpath
import re
import shutil
import sys
import tempfile
import unicodedata
from collections import OrderedDict
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P_NS, "a": A_NS, "r": R_NS}

PROTOCOL_SCHEMA_VERSION = "paper2video_editable_pptx.v2"
MANIFEST_SCHEMA_VERSION = "paper2video_animation_manifest.v1"
ALT_HANDLE_RE = re.compile(r"^\s*\[([^\]\n]+)\]\s*$")
LEGACY_ALT_ID_RE = re.compile(r"^\s*\[ID\]\s+(.+?)\s*$", re.IGNORECASE)
BLOCK_RE = re.compile(r"^\s{0,3}##\s+\[([^\]\n]+)\](?:\s+(.*\S))?\s*$")
# Match both the legacy point marker ``[[Name]]`` and the spoken-span marker
# ``[[Spotlight] words that remain in narration]``. The second capture is
# intentionally restricted to one line and one bracket-delimited span so a
# malformed protocol fails closed instead of swallowing adjacent Notes text.
MARKER_RE = re.compile(r"\[\[\s*([^\]\n]+?)\s*\](?:\s*([^\]\n]+?)\s*)?\]")
ALT_FIELD_RE = re.compile(
    r"^\s*(Animations|Script|Shape|Script-Baseline-SHA256)\s*:\s*(.*?)\s*$",
    re.IGNORECASE,
)
ALT_METADATA_START = "[Paper2Video]"
ALT_METADATA_END = "[/Paper2Video]"
SCRIPT_PROVENANCE_NS = (
    "https://github.com/microsoft/ResearchStudio/paper2video/"
    "script-provenance/2026"
)
SCRIPT_PROVENANCE_EXT_URI = SCRIPT_PROVENANCE_NS
SEMANTIC_PROVENANCE_NS = (
    "https://github.com/microsoft/ResearchStudio/paper2video/"
    "semantic-provenance/2026"
)
SEMANTIC_PROVENANCE_EXT_URI = SEMANTIC_PROVENANCE_NS
SCRIPT_HASH_RE = re.compile(r"^[0-9a-f]{64}$")
ORDER_SOURCES = {"author_notes", "animation_pane", "geometry"}
NARRATION_ORDER_POLICIES = {"geometry", "author_notes"}
CLICK_GROUP_POLICIES = {"normalize", "preserve"}

etree.register_namespace("p2v", SCRIPT_PROVENANCE_NS)
etree.register_namespace("p2vs", SEMANTIC_PROVENANCE_NS)

# PPT Master's established presetID/presetSubtype compatibility contract.
# presetID identifies the effect; presetSubtype carries its direction or shape
# variant, and one effect legitimately ships under several subtypes. PowerPoint
# itself emits one set, and generators driving the MsoAnimEffect enumeration
# emit another for the same visible effect -- ppt-master writes Wipe In as
# ("22", "4") where PowerPoint writes ("22", "1").
#
# Listing only one variant per effect rejected decks that were valid
# PowerPoint. Ten common entrances were affected, including Wipe and Zoom, so
# an authoring tool picking any of them aborted the render at bootstrap with
# "unsupported native PowerPoint entrance tuple". Both spellings map to the
# same strategy, so accepting both changes nothing downstream.
EFFECT_NAMES = {
    ("1", "0"): "Appear",
    ("10", "0"): "Fade In",
    ("2", "4"): "Fly In",
    ("42", "8"): "Cut In",
    ("23", "0"): "Zoom In",
    ("23", "16"): "Zoom In",
    ("22", "1"): "Wipe In",
    ("22", "4"): "Wipe In",
    ("16", "21"): "Split In",
    ("3", "10"): "Blinds In",
    ("5", "6"): "Checkerboard In",
    ("5", "10"): "Checkerboard In",
    ("9", "0"): "Dissolve In",
    ("14", "10"): "Random Bars In",
    ("12", "4"): "Peek In",
    ("21", "0"): "Wheel In",
    ("21", "1"): "Wheel In",
    ("4", "0"): "Box In",
    ("4", "16"): "Box In",
    ("6", "0"): "Circle In",
    ("6", "16"): "Circle In",
    ("8", "0"): "Diamond In",
    ("8", "16"): "Diamond In",
    ("13", "0"): "Plus In",
    ("13", "16"): "Plus In",
    ("18", "12"): "Strips In",
    ("20", "0"): "Wedge In",
    ("17", "0"): "Stretch In",
    ("17", "10"): "Stretch In",
    ("50", "0"): "Expand In",
    ("19", "0"): "Swivel In",
    ("19", "10"): "Swivel In",
    # presetID 42 carries two distinct effects rather than two directions of
    # one: subtype 8 is Cut In, subtype 0 is Ascend. Ascend rises into place,
    # so Fly In represents it far better than the generic fallback would.
    ("42", "0"): "Fly In",
}

# Entrances with no mapping at all. PowerPoint defines far more of them than
# this runtime implements strategies for -- Bounce, Spiral, Light Speed and
# roughly thirty others -- and any one of them used to abort the whole render.
#
# A deck is not broken for using one: the shape does enter, and Fade In is an
# honest reduction of "something appears here". Refusing to render is the worse
# answer, because it discards a deck that is otherwise entirely valid over one
# decorative choice. The substitution is reported per effect so it is visible
# rather than silent.
UNSUPPORTED_ENTRANCE_FALLBACK = "Fade In"

VIDEO_EFFECT_SECONDS = {
    "Appear": 0.12,
    "Fade In": 0.48,
    "Fly In": 0.56,
    "Cut In": 0.12,
    "Zoom In": 0.48,
    "Wipe In": 0.52,
    "Split In": 0.52,
    "Blinds In": 0.52,
    "Checkerboard In": 0.56,
    "Dissolve In": 0.48,
    "Random Bars In": 0.56,
    "Peek In": 0.52,
    "Wheel In": 0.56,
    "Box In": 0.52,
    "Circle In": 0.52,
    "Diamond In": 0.52,
    "Plus In": 0.52,
    "Strips In": 0.56,
    "Wedge In": 0.56,
    "Stretch In": 0.52,
    "Expand In": 0.52,
    "Swivel In": 0.56,
}


class ProtocolError(ValueError):
    """The editable PPTX protocol is incomplete, ambiguous, or inconsistent."""


def normalize_narration_order_policy(value: object) -> str:
    """Return the canonical narration-order policy name.

    Geometry is the ordinary pptx2video behavior. ``author_notes`` remains an
    explicit opt-in for authors who intentionally arrange the visible Notes
    blocks into a custom spoken sequence.
    """
    policy = _oneline(value).lower().replace("-", "_")
    if policy not in NARRATION_ORDER_POLICIES:
        choices = ", ".join(sorted(NARRATION_ORDER_POLICIES))
        raise ProtocolError(
            f"invalid narration order policy {value!r}; expected one of {choices}"
        )
    return policy


def normalize_click_group_policy(value: object) -> str:
    """Return the PowerPoint click-group normalization policy."""
    policy = _oneline(value).lower().replace("-", "_")
    if policy not in CLICK_GROUP_POLICIES:
        choices = ", ".join(sorted(CLICK_GROUP_POLICIES))
        raise ProtocolError(
            f"invalid click-group policy {value!r}; expected one of {choices}"
        )
    return policy


def _utc_now() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _oneline(value: object) -> str:
    return " ".join(str(value or "").split())


def _semantic_values(
    *,
    semantic_concise: object = "",
    semantic: object = "",
    semantic_detailed: object = "",
    transcript: object = "",
    visible_texts: Iterable[object] = (),
    shape_name: object = "",
    handle: object = "",
) -> tuple[str, str]:
    """Resolve both semantic tiers with deterministic, non-LLM fallbacks."""
    visible = [_oneline(value) for value in visible_texts if _oneline(value)]
    concise = (
        _oneline(semantic_concise)
        or _oneline(semantic)
        or (visible[0] if visible else "")
        or _oneline(shape_name)
        or _oneline(handle)
    )
    detailed = (
        _oneline(semantic_detailed)
        or _oneline(transcript)
        or _oneline(" | ".join(visible))
        or concise
    )
    return concise, detailed


def _hydrate_block_semantics(
    block: dict[str, object],
    shape: dict[str, object],
) -> tuple[str, str]:
    """Attach canonical semantic fields while retaining ``semantic`` as an alias."""
    concise, detailed = _semantic_values(
        semantic_concise=(
            block.get("semantic_concise") or shape.get("semantic_concise")
        ),
        semantic=block.get("semantic"),
        semantic_detailed=(
            block.get("semantic_detailed") or shape.get("semantic_detailed")
        ),
        transcript=block.get("transcript"),
        visible_texts=shape.get("visible_texts") or block.get("visible_texts") or [],
        shape_name=shape.get("shape_name") or block.get("shape_name"),
        handle=block.get("handle") or shape.get("handle"),
    )
    block["semantic"] = concise
    block["semantic_concise"] = concise
    block["semantic_detailed"] = detailed
    return concise, detailed


def _normalized_chars(value: object) -> str:
    return "".join(re.findall(r"[a-z0-9]+", str(value or "").lower()))


def normalize_script_for_hash(raw_script: object) -> str:
    """Return the stable marker-bearing script representation used for provenance."""
    return _oneline(unicodedata.normalize("NFKC", str(raw_script or "")))


def script_sha256(raw_script: object) -> str:
    """Hash narration without losing named animation markers or their order."""
    return hashlib.sha256(
        normalize_script_for_hash(raw_script).encode("utf-8")
    ).hexdigest()


def _script_provenance_from_cnvpr(
    node: etree._Element,
) -> dict[str, object] | None:
    ext_lst = node.find(f"{{{A_NS}}}extLst")
    if ext_lst is None:
        return None
    baselines: list[etree._Element] = []
    for extension in ext_lst.findall(f"{{{A_NS}}}ext"):
        if extension.get("uri") != SCRIPT_PROVENANCE_EXT_URI:
            continue
        baselines.extend(
            extension.findall(f"{{{SCRIPT_PROVENANCE_NS}}}scriptBaseline")
        )
    if not baselines:
        return None
    if len(baselines) != 1:
        raise ProtocolError("shape has duplicate Paper2Video script baseline metadata")
    value = str(baselines[0].get("sha256") or "").lower()
    if not SCRIPT_HASH_RE.fullmatch(value):
        raise ProtocolError("shape has an invalid Paper2Video script baseline SHA-256")
    provenance: dict[str, object] = {"sha256": value}
    order_source = _oneline(baselines[0].get("orderSource")).lower().replace(
        "-", "_"
    )
    if order_source:
        if order_source not in ORDER_SOURCES:
            raise ProtocolError(
                "shape has an invalid Paper2Video orderSource provenance value"
            )
        provenance["order_source"] = order_source
    raw_order_index = _oneline(baselines[0].get("orderIndex"))
    if raw_order_index:
        try:
            order_index = int(raw_order_index)
        except ValueError as exc:
            raise ProtocolError(
                "shape has an invalid Paper2Video orderIndex provenance value"
            ) from exc
        if order_index < 0:
            raise ProtocolError(
                "shape has a negative Paper2Video orderIndex provenance value"
            )
        provenance["order_index"] = order_index
    return provenance


def _script_baseline_from_cnvpr(node: etree._Element) -> str | None:
    provenance = _script_provenance_from_cnvpr(node)
    return str(provenance["sha256"]) if provenance is not None else None


def _semantic_provenance_from_cnvpr(
    node: etree._Element,
) -> dict[str, str] | None:
    """Read optional shape semantics from an independent OOXML extension."""
    ext_lst = node.find(f"{{{A_NS}}}extLst")
    if ext_lst is None:
        return None
    semantics: list[etree._Element] = []
    for extension in ext_lst.findall(f"{{{A_NS}}}ext"):
        if extension.get("uri") != SEMANTIC_PROVENANCE_EXT_URI:
            continue
        semantics.extend(
            extension.findall(f"{{{SEMANTIC_PROVENANCE_NS}}}blockSemantics")
        )
    if not semantics:
        return None
    if len(semantics) != 1:
        raise ProtocolError("shape has duplicate Paper2Video semantic metadata")
    concise_nodes = semantics[0].findall(
        f"{{{SEMANTIC_PROVENANCE_NS}}}concise"
    )
    detailed_nodes = semantics[0].findall(
        f"{{{SEMANTIC_PROVENANCE_NS}}}detailed"
    )
    if len(concise_nodes) > 1 or len(detailed_nodes) > 1:
        raise ProtocolError("shape has duplicate Paper2Video semantic tiers")
    concise = _oneline(concise_nodes[0].text if concise_nodes else "")
    detailed = _oneline(detailed_nodes[0].text if detailed_nodes else "")
    if not concise and not detailed:
        return None
    return {
        "semantic_concise": concise,
        "semantic_detailed": detailed,
    }


def _set_semantic_provenance_on_cnvpr(
    node: etree._Element,
    *,
    semantic_concise: object,
    semantic_detailed: object,
) -> dict[str, str]:
    """Write both semantic tiers without disturbing script/order provenance."""
    concise, detailed = _semantic_values(
        semantic_concise=semantic_concise,
        semantic_detailed=semantic_detailed,
    )
    ext_lst = node.find(f"{{{A_NS}}}extLst")
    if ext_lst is None:
        ext_lst = etree.SubElement(node, f"{{{A_NS}}}extLst")
    for extension in list(ext_lst.findall(f"{{{A_NS}}}ext")):
        if extension.get("uri") == SEMANTIC_PROVENANCE_EXT_URI:
            ext_lst.remove(extension)
    if not concise and not detailed:
        return {"semantic_concise": "", "semantic_detailed": ""}
    extension = etree.SubElement(
        ext_lst,
        f"{{{A_NS}}}ext",
        uri=SEMANTIC_PROVENANCE_EXT_URI,
    )
    semantics = etree.SubElement(
        extension,
        f"{{{SEMANTIC_PROVENANCE_NS}}}blockSemantics",
        schemaVersion="1",
    )
    concise_node = etree.SubElement(
        semantics,
        f"{{{SEMANTIC_PROVENANCE_NS}}}concise",
    )
    concise_node.text = concise
    detailed_node = etree.SubElement(
        semantics,
        f"{{{SEMANTIC_PROVENANCE_NS}}}detailed",
    )
    detailed_node.text = detailed
    return {
        "semantic_concise": concise,
        "semantic_detailed": detailed,
    }


def _set_script_baseline_on_cnvpr(
    node: etree._Element,
    raw_script: object,
    *,
    order_source: str | None = None,
    order_index: int | None = None,
) -> str:
    baseline_hash = script_sha256(raw_script)
    previous = _script_provenance_from_cnvpr(node) or {}
    effective_order_source = (
        _oneline(order_source).lower().replace("-", "_")
        if order_source is not None
        else str(previous.get("order_source") or "")
    )
    if effective_order_source and effective_order_source not in ORDER_SOURCES:
        raise ProtocolError(
            f"invalid Paper2Video order source {effective_order_source!r}"
        )
    effective_order_index = (
        int(order_index)
        if order_index is not None
        else previous.get("order_index")
    )
    if effective_order_index is not None and int(effective_order_index) < 0:
        raise ProtocolError("Paper2Video order index must not be negative")
    ext_lst = node.find(f"{{{A_NS}}}extLst")
    if ext_lst is None:
        ext_lst = etree.SubElement(node, f"{{{A_NS}}}extLst")
    for extension in list(ext_lst.findall(f"{{{A_NS}}}ext")):
        if extension.get("uri") == SCRIPT_PROVENANCE_EXT_URI:
            ext_lst.remove(extension)
    extension = etree.SubElement(
        ext_lst,
        f"{{{A_NS}}}ext",
        uri=SCRIPT_PROVENANCE_EXT_URI,
    )
    baseline = etree.SubElement(
        extension,
        f"{{{SCRIPT_PROVENANCE_NS}}}scriptBaseline",
    )
    baseline.set("algorithm", "sha256")
    baseline.set("normalization", "oneline-v1")
    baseline.set("sha256", baseline_hash)
    if effective_order_source:
        baseline.set("orderSource", effective_order_source)
    if effective_order_index is not None:
        baseline.set("orderIndex", str(int(effective_order_index)))
    return baseline_hash


def parse_alt_id(description: str | None) -> str | None:
    """Return the handle from canonical ``## [handle] semantic`` Alt Text.

    The previous ``[handle]`` and ``[ID] handle`` forms remain readable so
    delivered decks continue to rerender, but every writer emits the same block
    header grammar used by Author Notes.
    """
    if not description:
        return None
    first = description.splitlines()[0]
    match = BLOCK_RE.fullmatch(first)
    if match is None:
        match = ALT_HANDLE_RE.fullmatch(first)
    if match is None:
        match = LEGACY_ALT_ID_RE.fullmatch(first)
    if not match:
        return None
    handle = _oneline(match.group(1))
    if not handle:
        raise ProtocolError("Alt Text handle must not be empty")
    return handle


def _managed_alt_order_source(description: str | None) -> str | None:
    inside_managed = False
    for line in (description or "").splitlines()[1:]:
        marker = line.strip()
        if marker == ALT_METADATA_START:
            inside_managed = True
            continue
        if marker == ALT_METADATA_END:
            inside_managed = False
            continue
        if not inside_managed:
            continue
        match = re.fullmatch(r"\s*Order-Source\s*:\s*(.*?)\s*", line, re.IGNORECASE)
        if match is None:
            continue
        source = _oneline(match.group(1)).lower().replace("-", "_")
        if source in {"author_notes", "animation_pane", "geometry"}:
            return source
    return None


def build_system_alt_text(
    *,
    handle: str,
    animation_names: Iterable[str],
    raw_script: str,
    shape_name: str,
    shape_id: str,
    slide_index: int,
    existing_description: str | None = None,
    order_source: str = "author_notes",
    baseline_hash: str | None = None,
) -> str:
    """Build the compact Alt Text block using the Author Notes grammar.

    Older decks may contain a verbose ``[Paper2Video]`` block. Readers remain
    backward compatible with that format, but every writeback intentionally
    migrates it to one canonical block. Script provenance and ordering live in
    the shape's ``p2v:scriptBaseline`` OOXML extension, while native animation
    details remain in PowerPoint's ``p:timing`` tree.
    """
    clean_handle = _oneline(handle)
    if not clean_handle or "]" in clean_handle or "\n" in clean_handle:
        raise ProtocolError(f"invalid Alt Text handle: {handle!r}")
    semantic = _oneline(shape_name)
    header = f"## [{clean_handle}]"
    if semantic:
        header += f" {semantic}"
    script = canonicalize_user_script(raw_script)
    return f"{header}\n{script}" if script else header


def canonicalize_user_script(raw_script: object) -> str:
    """Keep narration and Spotlight while migrating legacy entrance markers.

    Native entrance effects are authored in PowerPoint's Animation Pane. Older
    decks that placed ``[[Fade In]]`` or similar markers in Notes remain
    readable, but normalized user-visible text drops that redundant control
    syntax. ``Spotlight`` remains because it is a video-only supplemental cue.
    """
    raw = str(raw_script or "")
    pieces: list[str] = []
    cursor = 0
    for match in MARKER_RE.finditer(raw):
        pieces.append(raw[cursor:match.start()])
        name = _oneline(match.group(1))
        if name == "Spotlight":
            scope = _oneline(match.group(2)) if match.group(2) is not None else ""
            pieces.append(f"[[Spotlight] {scope}]" if scope else "[[Spotlight]]")
        cursor = match.end()
    pieces.append(raw[cursor:])
    return _oneline("".join(pieces))


def parse_marked_transcript(raw: str) -> tuple[str, list[dict[str, object]]]:
    """Strip point markers while retaining optional Spotlight spoken spans."""
    pieces: list[str] = []
    markers: list[dict[str, object]] = []
    cursor = 0
    for match in MARKER_RE.finditer(raw or ""):
        pieces.append(raw[cursor:match.start()])
        spoken_prefix = "".join(pieces)
        name = _oneline(match.group(1))
        if not name:
            raise ProtocolError("animation marker name must not be empty")
        marker = {
            "name": name,
            "word": len(spoken_prefix.split()),
            "normalized_char": len(_normalized_chars(spoken_prefix)),
        }
        if match.group(2) is not None:
            scope_text = _oneline(match.group(2))
            if not scope_text or not _normalized_chars(scope_text):
                raise ProtocolError("Spotlight spoken scope must not be empty")
            if name != "Spotlight":
                raise ProtocolError(
                    f"spoken-span syntax is supported only for Spotlight, not {name!r}"
                )
            marker["scope_text"] = scope_text
            marker["normalized_end_char"] = (
                int(marker["normalized_char"]) + len(_normalized_chars(scope_text))
            )
            pieces.append(scope_text)
        markers.append(marker)
        cursor = match.end()
    pieces.append((raw or "")[cursor:])
    clean = _oneline("".join(pieces))
    return clean, markers


def parse_alt_protocol(description: str | None) -> dict[str, object] | None:
    """Read narration and Spotlight cues from Shape Alt Text.

    Canonical Alt Text is one ``## [handle]`` block, identical to the Author
    Notes grammar. Legacy ``[handle]\nScript: ...`` and ``[Paper2Video]``
    metadata remain readable for migration.
    """
    handle = parse_alt_id(description)
    if not handle:
        return None
    first_line = (description or "").splitlines()[0] if description else ""
    if BLOCK_RE.fullmatch(first_line):
        blocks = parse_notes_blocks(description or "")
        if len(blocks) != 1:
            raise ProtocolError("Shape Alt Text must contain exactly one ## [handle] block")
        return blocks[0]

    fields: dict[str, str] = {}
    managed_fields: set[str] = set()
    inside_managed = False
    for line in (description or "").splitlines()[1:]:
        marker = line.strip()
        if marker == ALT_METADATA_START:
            inside_managed = True
            continue
        if marker == ALT_METADATA_END:
            inside_managed = False
            continue
        match = ALT_FIELD_RE.fullmatch(line)
        if match:
            key = match.group(1).lower()
            fields[key] = match.group(2).strip()
            if inside_managed:
                managed_fields.add(key)
    if "script" not in fields:
        return None
    legacy_raw_script = fields.get("script", "")
    legacy_marker_names = [
        _oneline(match.group(1))
        for match in MARKER_RE.finditer(legacy_raw_script)
        if _oneline(match.group(1)) != "Spotlight"
    ]
    raw_script = canonicalize_user_script(legacy_raw_script)
    if not raw_script:
        return None
    transcript, markers = parse_marked_transcript(raw_script)
    animation_names = list(legacy_marker_names)
    if "animations" not in managed_fields:
        animation_names.extend(
            [
            _oneline(name)
            for name in re.split(r"\s*;\s*", fields.get("animations", ""))
            if _oneline(name)
            ]
        )
    semantic = _oneline(fields.get("shape"))
    return {
        "handle": handle,
        "semantic": semantic,
        "semantic_concise": semantic,
        "semantic_detailed": "",
        "raw_transcript": raw_script,
        "transcript": transcript,
        "markers": markers,
        "animation_names": animation_names,
        "baseline_hash_mirror": fields.get("script-baseline-sha256") or None,
    }


def _legacy_plain_alt_script(description: str | None) -> str:
    """Return narration from a pre-protocol plain Alt Text description.

    This compatibility path is used only while normalizing an animated target
    that has neither canonical Notes nor an explicit managed ``Script:``. Once
    consumed, the text is written back as a real Script field and canonical
    Notes, so later renders no longer depend on this heuristic.
    """
    text = str(description or "").strip()
    if not text or parse_alt_id(text) is not None or ALT_METADATA_START in text:
        return ""
    if any(ALT_FIELD_RE.fullmatch(line) for line in text.splitlines()):
        return ""
    return _oneline(text)


def parse_notes_blocks(notes: str) -> list[dict[str, object]]:
    """Parse the canonical ``## [handle]`` Author Notes grammar."""
    blocks: list[dict[str, object]] = []
    current: dict[str, object] | None = None
    preamble: list[str] = []
    for line_number, line in enumerate((notes or "").splitlines(), start=1):
        match = BLOCK_RE.fullmatch(line)
        if match:
            handle = _oneline(match.group(1))
            if not handle:
                raise ProtocolError(f"line {line_number}: block handle must not be empty")
            current = {
                "handle": handle,
                "semantic": _oneline(match.group(2)),
                "semantic_concise": _oneline(match.group(2)),
                "semantic_detailed": "",
                "lines": [],
            }
            blocks.append(current)
        elif current is None:
            if line.strip():
                preamble.append(line.strip())
        else:
            current["lines"].append(line)  # type: ignore[index,union-attr]
    if preamble:
        raise ProtocolError("Author Notes contain text before the first ## [handle] block")
    if not blocks:
        raise ProtocolError("Author Notes contain no ## [handle] blocks")

    seen: set[str] = set()
    parsed: list[dict[str, object]] = []
    for block in blocks:
        handle = str(block["handle"])
        if handle in seen:
            raise ProtocolError(f"duplicate Author Notes handle: {handle!r}")
        seen.add(handle)
        legacy_raw = "\n".join(block["lines"]).strip()  # type: ignore[arg-type]
        animation_names = [
            _oneline(match.group(1))
            for match in MARKER_RE.finditer(legacy_raw)
            if _oneline(match.group(1)) != "Spotlight"
        ]
        raw = canonicalize_user_script(legacy_raw)
        transcript, markers = parse_marked_transcript(raw)
        parsed.append(
            {
                "handle": handle,
                "semantic": block["semantic"],
                "semantic_concise": block["semantic_concise"],
                "semantic_detailed": block["semantic_detailed"],
                "raw_transcript": raw,
                "transcript": transcript,
                "markers": markers,
                "animation_names": animation_names,
            }
        )
    return parsed


def _resolve_part(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _rels_part(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def _relationships(archive: ZipFile, source_part: str) -> dict[str, tuple[str, str]]:
    rels_name = _rels_part(source_part)
    if rels_name not in archive.namelist():
        return {}
    root = etree.fromstring(archive.read(rels_name))
    result: dict[str, tuple[str, str]] = {}
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rid = str(rel.get("Id") or "")
        target = str(rel.get("Target") or "")
        rel_type = str(rel.get("Type") or "")
        if rid and target:
            result[rid] = (_resolve_part(source_part, target), rel_type)
    return result


def presentation_slides(archive: ZipFile) -> list[dict[str, object]]:
    """Return slide parts in presentation order with stable ``p:sldId`` IDs."""
    presentation_part = "ppt/presentation.xml"
    root = etree.fromstring(archive.read(presentation_part))
    rels = _relationships(archive, presentation_part)
    slides: list[dict[str, object]] = []
    for index, node in enumerate(root.xpath("./p:sldIdLst/p:sldId", namespaces=NS), start=1):
        rid = str(node.get(f"{{{R_NS}}}id") or "")
        relation = rels.get(rid)
        if relation is None:
            raise ProtocolError(f"presentation slide relationship {rid!r} is missing")
        slides.append(
            {
                "index": index,
                "stable_id": str(node.get("id") or index),
                "part": relation[0],
            }
        )
    if not slides:
        raise ProtocolError("PPTX presentation contains no slides")
    return slides


def presentation_size(archive: ZipFile) -> tuple[int, int]:
    root = etree.fromstring(archive.read("ppt/presentation.xml"))
    nodes = root.xpath("./p:sldSz", namespaces=NS)
    if len(nodes) != 1:
        raise ProtocolError("PPTX presentation must contain one p:sldSz")
    try:
        width = int(nodes[0].get("cx"))
        height = int(nodes[0].get("cy"))
    except (TypeError, ValueError) as exc:
        raise ProtocolError("PPTX presentation has invalid slide dimensions") from exc
    if width <= 0 or height <= 0:
        raise ProtocolError("PPTX presentation has non-positive slide dimensions")
    return width, height


def _notes_part_optional(archive: ZipFile, slide_part: str) -> str | None:
    notes_suffix = "/notesSlide"
    for target, rel_type in _relationships(archive, slide_part).values():
        if rel_type.endswith(notes_suffix):
            return target
    return None


def _notes_part(archive: ZipFile, slide_part: str) -> str:
    notes_part = _notes_part_optional(archive, slide_part)
    if notes_part is not None:
        return notes_part
    raise ProtocolError(f"{slide_part} has no Author Notes part")


def _notes_text(archive: ZipFile, slide_part: str) -> str:
    notes_part = _notes_part(archive, slide_part)
    root = etree.fromstring(archive.read(notes_part))
    bodies = root.xpath(
        './/p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]', namespaces=NS
    )
    if len(bodies) != 1:
        raise ProtocolError(
            f"{notes_part} must contain exactly one notes body placeholder, found {len(bodies)}"
        )
    paragraphs = bodies[0].xpath("./p:txBody/a:p", namespaces=NS)
    return "\n".join("".join(p.xpath(".//a:t/text()", namespaces=NS)) for p in paragraphs)


def _notes_text_optional(archive: ZipFile, slide_part: str) -> str:
    notes_part = _notes_part_optional(archive, slide_part)
    if notes_part is None:
        return ""
    root = etree.fromstring(archive.read(notes_part))
    bodies = root.xpath(
        './/p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]', namespaces=NS
    )
    if not bodies:
        return ""
    if len(bodies) != 1:
        raise ProtocolError(
            f"{notes_part} must contain at most one notes body placeholder, found {len(bodies)}"
        )
    paragraphs = bodies[0].xpath("./p:txBody/a:p", namespaces=NS)
    return "\n".join("".join(p.xpath(".//a:t/text()", namespaces=NS)) for p in paragraphs)


def _canonical_notes_blocks(notes: str) -> list[dict[str, object]]:
    """Return canonical blocks, treating ordinary presenter notes as absent."""
    if not any(BLOCK_RE.fullmatch(line) for line in (notes or "").splitlines()):
        return []
    return parse_notes_blocks(notes)


def _native_effects(slide_root: etree._Element) -> list[dict[str, object]]:
    effects: list[dict[str, object]] = []
    unsupported_classes = sorted(
        {
            str(value)
            for value in slide_root.xpath('.//p:cTn/@presetClass', namespaces=NS)
            if str(value) not in {"entr", "emph"}
        }
    )
    if unsupported_classes:
        raise ProtocolError(
            "unsupported native PowerPoint animation classes: "
            + ", ".join(unsupported_classes)
        )
    for ctn in slide_root.xpath(
        './/p:cTn[@presetClass="entr" or @presetClass="emph"]', namespaces=NS
    ):
        spids = list(OrderedDict.fromkeys(ctn.xpath(".//p:spTgt/@spid", namespaces=NS)))
        if len(spids) != 1:
            raise ProtocolError(
                "each native entrance effect must resolve to exactly one shape target"
            )
        effect_kind = str(ctn.get("presetClass") or "")
        preset_id = str(ctn.get("presetID") or "")
        subtype = str(ctn.get("presetSubtype") or "0")
        if effect_kind == "entr":
            name = EFFECT_NAMES.get((preset_id, subtype))
            if name is None:
                name = UNSUPPORTED_ENTRANCE_FALLBACK
                print(
                    "[video_pptx_protocol] entrance tuple "
                    f"{(preset_id, subtype)!r} has no mapped strategy; "
                    f"rendering it as {name}",
                    file=sys.stderr,
                    flush=True,
                )
            kind = "entrance"
        else:
            name = "Spotlight"
            kind = "emphasis"
        durations = [
            int(value)
            for value in ctn.xpath(".//p:cTn/@dur", namespaces=NS)
            if str(value).isdigit()
        ]
        raw_delays = ctn.xpath("./p:stCondLst/p:cond/@delay", namespaces=NS)
        delay_ms = next(
            (int(value) for value in raw_delays if str(value).isdigit()),
            0,
        )
        effects.append(
            {
                "native_order": len(effects) + 1,
                "shape_id": spids[0],
                "name": name,
                "kind": kind,
                "preset_id": preset_id,
                "preset_subtype": subtype,
                "trigger": str(ctn.get("nodeType") or ""),
                "delay_seconds": round(delay_ms / 1000.0, 3),
                "duration_seconds": (
                    round(max(durations) / 1000.0, 3)
                    if durations
                    else (2.4 if kind == "emphasis" else None)
                ),
            }
        )

    timeline_end = 0.0
    previous_start = 0.0
    previous_end = 0.0
    click_group = 0
    simultaneous_group = 0
    for index, effect in enumerate(effects):
        trigger = str(effect.get("trigger") or "").lower()
        delay = float(effect.get("delay_seconds") or 0.0)
        duration = float(
            effect.get("duration_seconds")
            or (
                2.4
                if str(effect.get("kind") or "") == "emphasis"
                else VIDEO_EFFECT_SECONDS.get(str(effect.get("name") or ""), 0.48)
            )
        )
        if index == 0:
            click_group = 1
            simultaneous_group = 1
            start = delay
        elif trigger == "witheffect":
            start = previous_start + delay
        elif trigger == "aftereffect":
            simultaneous_group += 1
            start = previous_end + delay
        else:
            click_group += 1
            simultaneous_group += 1
            start = timeline_end + delay
        end = start + max(0.0, duration)
        effect["pane_start_seconds"] = round(start, 3)
        effect["pane_end_seconds"] = round(end, 3)
        effect["click_group"] = click_group
        effect["simultaneous_group"] = simultaneous_group
        previous_start = start
        previous_end = end
        timeline_end = max(timeline_end, end)
    return effects


def _single_child(parent: etree._Element, tag: str, *, label: str) -> etree._Element:
    children = parent.findall(tag)
    if len(children) != 1:
        raise ProtocolError(
            f"{label} must contain exactly one {etree.QName(tag).localname}"
        )
    return children[0]


def _numeric_condition_delay(
    ctn: etree._Element,
    *,
    label: str,
) -> tuple[etree._Element, int]:
    conditions = ctn.xpath("./p:stCondLst/p:cond", namespaces=NS)
    numeric: list[tuple[etree._Element, int]] = []
    for condition in conditions:
        raw = str(condition.get("delay") or "")
        if raw.isdigit():
            numeric.append((condition, int(raw)))
    if len(numeric) != 1:
        raise ProtocolError(f"{label} must contain exactly one numeric wrapper delay")
    return numeric[0]


def _main_sequence_row_groups(
    main_sequence: etree._Element,
) -> tuple[etree._Element, list[list[dict[str, object]]]]:
    """Return exact Animation Pane rows grouped by current top-level click group.

    PowerPoint stores each click badge as a direct ``mainSeq`` child. Each such
    child contains one inner row list. A row is either a preset ``p:cTn``
    directly or a structural wrapper whose numeric delay places the row on the
    local Pane clock. This parser intentionally fails closed for other timing
    topologies instead of dropping behaviors while rewriting them.
    """
    child_list = _single_child(
        main_sequence,
        f"{{{P_NS}}}childTnLst",
        label="PowerPoint mainSeq",
    )
    direct_groups = list(child_list)
    if not direct_groups:
        return child_list, []
    if any(group.tag != f"{{{P_NS}}}par" for group in direct_groups):
        raise ProtocolError("PowerPoint mainSeq contains a non-parallel top-level row group")

    grouped_rows: list[list[dict[str, object]]] = []
    for group_index, group in enumerate(direct_groups, start=1):
        outer = _single_child(
            group,
            f"{{{P_NS}}}cTn",
            label=f"PowerPoint click group {group_index}",
        )
        row_list = _single_child(
            outer,
            f"{{{P_NS}}}childTnLst",
            label=f"PowerPoint click group {group_index} outer timing node",
        )
        row_pars = list(row_list)
        if not row_pars or any(row.tag != f"{{{P_NS}}}par" for row in row_pars):
            raise ProtocolError(
                f"PowerPoint click group {group_index} has an unsupported row list"
            )
        rows: list[dict[str, object]] = []
        for row_index, row in enumerate(row_pars, start=1):
            row_root = _single_child(
                row,
                f"{{{P_NS}}}cTn",
                label=f"PowerPoint click group {group_index} row {row_index}",
            )
            effects = row.xpath(
                './/p:cTn[@presetClass="entr" or @presetClass="emph"]',
                namespaces=NS,
            )
            if len(effects) != 1:
                raise ProtocolError(
                    f"PowerPoint click group {group_index} row {row_index} must "
                    "contain exactly one supported preset effect"
                )
            effect = effects[0]
            if row_root is effect:
                wrapper_condition = None
                wrapper_delay_ms = 0
            else:
                if effect not in row_root.iterdescendants():
                    raise ProtocolError(
                        f"PowerPoint click group {group_index} row {row_index} "
                        "has an unsupported effect wrapper"
                    )
                wrapper_condition, wrapper_delay_ms = _numeric_condition_delay(
                    row_root,
                    label=f"PowerPoint click group {group_index} row {row_index}",
                )
            trigger = str(effect.get("nodeType") or "")
            if trigger.lower() not in {"clickeffect", "aftereffect", "witheffect"}:
                raise ProtocolError(
                    f"PowerPoint click group {group_index} row {row_index} has "
                    f"unsupported trigger {trigger!r}"
                )
            shape_ids = list(
                OrderedDict.fromkeys(
                    effect.xpath(".//p:spTgt/@spid", namespaces=NS)
                )
            )
            if len(shape_ids) != 1:
                raise ProtocolError(
                    f"PowerPoint click group {group_index} row {row_index} must "
                    "resolve to exactly one shape"
                )
            rows.append(
                {
                    "row": row,
                    "effect": effect,
                    "trigger": trigger,
                    "wrapper_condition": wrapper_condition,
                    "wrapper_delay_ms": wrapper_delay_ms,
                    "shape_id": str(shape_ids[0]),
                    "source_group": group_index,
                }
            )
        grouped_rows.append(rows)
    return child_list, grouped_rows


def _click_phase_partition(
    grouped_rows: list[list[dict[str, object]]],
) -> list[list[dict[str, object]]]:
    phases: list[list[dict[str, object]]] = []
    for rows in grouped_rows:
        current: list[dict[str, object]] | None = None
        for row in rows:
            trigger = str(row["trigger"]).lower()
            if trigger == "witheffect":
                if current is None:
                    raise ProtocolError(
                        "With Previous cannot be the first row of a PowerPoint click group"
                    )
                current.append(row)
            else:
                current = [row]
                phases.append(current)
    return phases


def _main_sequence_is_numbered(
    main_sequence: etree._Element,
    grouped_rows: list[list[dict[str, object]]],
) -> bool:
    if not grouped_rows:
        return True
    for rows in grouped_rows:
        if str(rows[0]["trigger"]).lower() != "clickeffect":
            return False
        if any(
            str(row["trigger"]).lower() != "witheffect" for row in rows[1:]
        ):
            return False
    for condition in main_sequence.xpath(
        "./p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond",
        namespaces=NS,
    ):
        if str(condition.get("evt") or "").lower() == "onbegin":
            return False
    return True


def _fresh_timing_id(slide_root: etree._Element, next_id: list[int]) -> str:
    value = next_id[0]
    next_id[0] += 1
    return str(value)


def _new_numbered_click_group(
    slide_root: etree._Element,
    rows: list[dict[str, object]],
    *,
    next_id: list[int],
) -> etree._Element:
    phase = etree.Element(f"{{{P_NS}}}par")
    outer = etree.SubElement(
        phase,
        f"{{{P_NS}}}cTn",
        id=_fresh_timing_id(slide_root, next_id),
        fill="hold",
    )
    outer_conditions = etree.SubElement(outer, f"{{{P_NS}}}stCondLst")
    etree.SubElement(outer_conditions, f"{{{P_NS}}}cond", delay="indefinite")
    row_list = etree.SubElement(outer, f"{{{P_NS}}}childTnLst")

    leader_delay = int(rows[0]["wrapper_delay_ms"])
    for row_index, source_row in enumerate(rows):
        row = deepcopy(source_row["row"])
        effects = row.xpath(
            './/p:cTn[@presetClass="entr" or @presetClass="emph"]',
            namespaces=NS,
        )
        if len(effects) != 1:
            raise ProtocolError("PowerPoint animation row changed during normalization")
        effect = effects[0]
        if row_index == 0:
            effect.set("nodeType", "clickEffect")
            effect.set("grpId", str(effect.get("grpId") or "0"))
        else:
            effect.set("nodeType", "withEffect")

        row_root = _single_child(
            row,
            f"{{{P_NS}}}cTn",
            label="copied PowerPoint animation row",
        )
        if row_root is not effect:
            condition, old_delay = _numeric_condition_delay(
                row_root,
                label="copied PowerPoint animation row",
            )
            rebased = old_delay - leader_delay
            if rebased < 0:
                raise ProtocolError(
                    "PowerPoint animation wrapper delay becomes negative after "
                    "click-group normalization"
                )
            condition.set("delay", str(rebased))
        elif leader_delay:
            raise ProtocolError(
                "PowerPoint direct animation row has a non-zero structural delay"
            )
        row_list.append(row)
    return phase


def _normalized_click_group_summary(
    main_sequence: etree._Element,
) -> list[dict[str, object]]:
    _, grouped_rows = _main_sequence_row_groups(main_sequence)
    groups: list[dict[str, object]] = []
    for badge, rows in enumerate(grouped_rows, start=1):
        groups.append(
            {
                "badge": badge,
                "shape_ids": [str(row["shape_id"]) for row in rows],
                "triggers": [str(row["trigger"]) for row in rows],
                "wrapper_delays_ms": [
                    int(row["wrapper_delay_ms"]) for row in rows
                ],
            }
        )
    return groups


def _normalize_slide_animation_click_groups(
    slide_root: etree._Element,
) -> dict[str, object]:
    """Make every sequential Pane phase a numbered top-level click group."""
    main_sequences = slide_root.xpath(
        './/p:cTn[@nodeType="mainSeq"]', namespaces=NS
    )
    if not main_sequences:
        return {
            "changed": False,
            "effect_count": 0,
            "old_group_count": 0,
            "new_group_count": 0,
            "changed_trigger_count": 0,
            "preserved_with_previous_count": 0,
            "groups": [],
        }
    if len(main_sequences) != 1:
        raise ProtocolError("PPTX slide must contain at most one PowerPoint mainSeq")
    main_sequence = main_sequences[0]
    child_list, grouped_rows = _main_sequence_row_groups(main_sequence)
    rows = [row for group in grouped_rows for row in group]
    if not rows:
        return {
            "changed": False,
            "effect_count": 0,
            "old_group_count": len(grouped_rows),
            "new_group_count": len(grouped_rows),
            "changed_trigger_count": 0,
            "preserved_with_previous_count": 0,
            "groups": [],
        }
    phases = _click_phase_partition(grouped_rows)
    if _main_sequence_is_numbered(main_sequence, grouped_rows):
        groups = _normalized_click_group_summary(main_sequence)
        return {
            "changed": False,
            "effect_count": len(rows),
            "old_group_count": len(grouped_rows),
            "new_group_count": len(groups),
            "changed_trigger_count": 0,
            "preserved_with_previous_count": sum(
                1
                for row in rows
                if str(row["trigger"]).lower() == "witheffect"
            ),
            "groups": groups,
        }

    before = _native_effects(slide_root)
    before_rows = [
        {
            "native_order": int(effect["native_order"]),
            "shape_id": str(effect["shape_id"]),
            "trigger": str(effect["trigger"]),
        }
        for effect in before
    ]
    timing_ids = [
        int(value)
        for value in slide_root.xpath(".//p:timing//p:cTn/@id", namespaces=NS)
        if str(value).isdigit()
    ]
    next_id = [max(timing_ids, default=0) + 1]
    for child in list(child_list):
        child_list.remove(child)
    for phase_rows in phases:
        child_list.append(
            _new_numbered_click_group(
                slide_root,
                phase_rows,
                next_id=next_id,
            )
        )

    all_ids = slide_root.xpath(".//p:timing//p:cTn/@id", namespaces=NS)
    if len(all_ids) != len(set(str(value) for value in all_ids)):
        raise ProtocolError("click-group normalization produced duplicate p:cTn IDs")
    known_ids = {str(value) for value in all_ids}
    dangling = sorted(
        {
            str(value)
            for value in slide_root.xpath(".//p:timing//p:tn/@val", namespaces=NS)
            if str(value) not in known_ids
        }
    )
    if dangling:
        raise ProtocolError(
            "click-group normalization produced dangling timing references: "
            + ", ".join(dangling)
        )

    after = _native_effects(slide_root)
    invariant_fields = (
        "native_order",
        "shape_id",
        "name",
        "kind",
        "preset_id",
        "preset_subtype",
        "delay_seconds",
        "duration_seconds",
        "pane_start_seconds",
        "pane_end_seconds",
        "simultaneous_group",
    )
    before_invariants = [
        {field: effect.get(field) for field in invariant_fields} for effect in before
    ]
    after_invariants = [
        {field: effect.get(field) for field in invariant_fields} for effect in after
    ]
    if after_invariants != before_invariants:
        raise ProtocolError(
            "click-group normalization would change the renderer's native Pane schedule"
        )
    groups = _normalized_click_group_summary(main_sequence)
    return {
        "changed": True,
        "effect_count": len(rows),
        "old_group_count": len(grouped_rows),
        "new_group_count": len(groups),
        "changed_trigger_count": sum(
            1
            for old, new in zip(before, after)
            if str(old.get("trigger") or "") != str(new.get("trigger") or "")
        ),
        "preserved_with_previous_count": sum(
            1
            for effect in after
            if str(effect.get("trigger") or "").lower() == "witheffect"
        ),
        "rows": [
            {
                **old,
                "after_trigger": str(new.get("trigger") or ""),
                "badge": int(new.get("click_group") or 0),
            }
            for old, new in zip(before_rows, after)
        ],
        "groups": groups,
    }


def normalize_animation_click_groups(
    source_pptx: Path,
    output_pptx: Path,
    *,
    policy: str = "normalize",
    slide_indices: Iterable[int] | None = None,
) -> dict[str, object]:
    """Write a PPTX whose sequential animation phases have numbered badges.

    The default processes every slide. ``With Previous`` rows remain in their
    leader's click group. The transform is accepted only when the renderer's
    Pane start/end clock and simultaneous groups remain unchanged.
    """
    source_pptx = source_pptx.resolve()
    output_pptx = output_pptx.resolve()
    policy = normalize_click_group_policy(policy)
    if not source_pptx.is_file():
        raise ProtocolError(f"PPTX not found: {source_pptx}")
    source_hash = file_sha256(source_pptx)
    requested = (
        {int(value) for value in slide_indices}
        if slide_indices is not None
        else None
    )
    if requested is not None and any(value <= 0 for value in requested):
        raise ProtocolError("slide indices must be positive")
    replacements: dict[str, bytes] = {}
    slides_out: list[dict[str, object]] = []
    temporary_path: Path | None = None
    try:
        with ZipFile(source_pptx) as source:
            if any(name.startswith("_xmlsignatures/") for name in source.namelist()):
                raise ProtocolError(
                    "digitally signed PPTX files cannot be safely normalized"
                )
            refs = presentation_slides(source)
            available = {int(ref["index"]) for ref in refs}
            if requested is not None:
                unknown = sorted(requested - available)
                if unknown:
                    raise ProtocolError(
                        "unknown slide indices: " + ", ".join(map(str, unknown))
                    )
            for ref in refs:
                index = int(ref["index"])
                part = str(ref["part"])
                selected = requested is None or index in requested
                if policy == "preserve" or not selected:
                    slides_out.append(
                        {
                            "index": index,
                            "stable_slide_id": str(ref["stable_id"]),
                            "selected": selected,
                            "changed": False,
                        }
                    )
                    continue
                slide_root = etree.fromstring(source.read(part))
                result = _normalize_slide_animation_click_groups(slide_root)
                result.update(
                    {
                        "index": index,
                        "stable_slide_id": str(ref["stable_id"]),
                        "selected": True,
                    }
                )
                slides_out.append(result)
                if result["changed"]:
                    replacements[part] = etree.tostring(
                        slide_root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )

            output_pptx.parent.mkdir(parents=True, exist_ok=True)
            if replacements:
                with tempfile.NamedTemporaryFile(
                    prefix=output_pptx.stem + ".click-groups.",
                    suffix=".pptx",
                    dir=output_pptx.parent,
                    delete=False,
                ) as temporary:
                    temporary_path = Path(temporary.name)
                with ZipFile(temporary_path, "w") as destination:
                    for info in source.infolist():
                        destination.writestr(
                            _copy_zipinfo(info),
                            replacements.get(info.filename, source.read(info.filename)),
                        )
        if temporary_path is not None:
            shutil.move(temporary_path, output_pptx)
        elif source_pptx != output_pptx:
            shutil.copy2(source_pptx, output_pptx)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

    return {
        "schema_version": "paper2video_click_group_normalization.v1",
        "created_at": _utc_now(),
        "policy": policy,
        "source_pptx": str(source_pptx),
        "source_sha256": source_hash,
        "output_pptx": str(output_pptx),
        "output_sha256": file_sha256(output_pptx),
        "slide_count": len(slides_out),
        "selected_slide_count": sum(1 for slide in slides_out if slide["selected"]),
        "changed_slide_count": sum(1 for slide in slides_out if slide["changed"]),
        "changed_trigger_count": sum(
            int(slide.get("changed_trigger_count") or 0) for slide in slides_out
        ),
        "preserved_with_previous_count": sum(
            int(slide.get("preserved_with_previous_count") or 0)
            for slide in slides_out
        ),
        "slides": slides_out,
    }


def _shape_maps(
    slide_root: etree._Element,
    *,
    slide_width: int,
    slide_height: int,
) -> tuple[dict[str, dict[str, object]], dict[str, dict[str, object]]]:
    by_id: dict[str, dict[str, object]] = {}
    by_handle: dict[str, dict[str, object]] = {}
    for node in slide_root.xpath(".//p:spTree//p:cNvPr", namespaces=NS):
        shape_id = str(node.get("id") or "")
        if not shape_id:
            continue
        top = node
        while top.getparent() is not None and top.getparent().tag != f"{{{P_NS}}}spTree":
            top = top.getparent()
        top_id_nodes = top.xpath(
            "./p:nvSpPr/p:cNvPr | ./p:nvPicPr/p:cNvPr | "
            "./p:nvGraphicFramePr/p:cNvPr | ./p:nvGrpSpPr/p:cNvPr",
            namespaces=NS,
        )
        top_id = str(top_id_nodes[0].get("id") or "") if top_id_nodes else ""
        xfrms = top.xpath(
            "./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./p:xfrm",
            namespaces=NS,
        )
        bbox: list[float] | None = None
        if xfrms:
            offsets = xfrms[0].xpath("./a:off", namespaces=NS)
            extents = xfrms[0].xpath("./a:ext", namespaces=NS)
            if offsets and extents:
                try:
                    x = int(offsets[0].get("x"))
                    y = int(offsets[0].get("y"))
                    width = int(extents[0].get("cx"))
                    height = int(extents[0].get("cy"))
                    bbox = [
                        round(x / slide_width, 6),
                        round(y / slide_height, 6),
                        round(width / slide_width, 6),
                        round(height / slide_height, 6),
                    ]
                except (TypeError, ValueError):
                    bbox = None
        provenance = _script_provenance_from_cnvpr(node)
        semantic_provenance = _semantic_provenance_from_cnvpr(node) or {}
        info: dict[str, object] = {
            "shape_id": shape_id,
            "shape_name": str(node.get("name") or ""),
            "description": str(node.get("descr") or ""),
            "visible_texts": [
                text
                for text in OrderedDict(
                    (_oneline(value), None)
                    for value in top.xpath(".//a:t/text()", namespaces=NS)
                    if _oneline(value)
                )
            ],
            "script_baseline_sha256": (
                str(provenance["sha256"]) if provenance is not None else None
            ),
            "order_source": (
                str(provenance.get("order_source") or "")
                if provenance is not None
                else None
            ),
            "order_index": (
                int(provenance["order_index"])
                if provenance is not None and "order_index" in provenance
                else None
            ),
            "semantic_concise": str(
                semantic_provenance.get("semantic_concise") or ""
            ),
            "semantic_detailed": str(
                semantic_provenance.get("semantic_detailed") or ""
            ),
            "top_level": top_id == shape_id,
            "bbox": bbox,
        }
        if shape_id in by_id:
            raise ProtocolError(f"duplicate shape id {shape_id!r} on one slide")
        by_id[shape_id] = info
        handle = parse_alt_id(info["description"])
        if handle:
            if handle in by_handle:
                raise ProtocolError(f"duplicate Alt Text handle {handle!r} on one slide")
            info["handle"] = handle
            by_handle[handle] = info
    return by_id, by_handle


def _section_ids_from_script(path: Path) -> list[str]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"could not read section IDs from {path}: {exc}") from exc
    sections = payload.get("sections") or []
    ids = [str(section.get("id") or "").strip() for section in sections]
    if not ids or any(not item for item in ids):
        raise ProtocolError(f"{path} has missing or empty section IDs")
    return ids


def _effect_kind(name: str, *, context: str) -> str:
    if name == "Spotlight":
        return "emphasis"
    if name in VIDEO_EFFECT_SECONDS:
        return "entrance"
    raise ProtocolError(f"{context} requests unsupported video effect {name!r}")


def _default_marker(name: str, source: str) -> dict[str, object]:
    return {
        "name": name,
        "word": 0,
        "normalized_char": 0,
        "source": source,
    }


def _merge_effect_intents(
    *,
    shape_id: str,
    handle: str,
    native_effects: list[dict[str, object]],
    markers: list[dict[str, object]],
    animation_names: list[str],
    authority: str,
    synthetic_order_base: int,
) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    """Keep native rows authoritative and add only supplemental Spotlight cues.

    Legacy script entrance markers are reported but never replace, create, or
    reorder a native animation. This preserves old decks without exposing a
    second user-visible animation system beside PowerPoint's Animation Pane.
    """
    explicit_markers = [dict(marker) for marker in markers]
    explicit_markers.extend(
        _default_marker(name, f"{authority}_legacy") for name in animation_names
    )
    markers_out: list[dict[str, object]] = []
    effects_out: list[dict[str, object]] = []
    conflicts: list[dict[str, object]] = []
    synthetic_index = 0

    for native in native_effects:
        native_name = str(native["name"])
        markers_out.append(_default_marker(native_name, "animation_pane_default"))
        effects_out.append(
            {
                **native,
                "native_name": native_name,
                "authority": "animation_pane",
                "native_present": True,
            }
        )

    for marker in explicit_markers:
        requested_name = str(marker["name"])
        if requested_name != "Spotlight":
            matching_native = next(
                (
                    effect
                    for effect in native_effects
                    if str(effect.get("kind") or "entrance") == "entrance"
                ),
                None,
            )
            conflicts.append(
                {
                    "handle": handle,
                    "shape_id": shape_id,
                    "kind": "entrance",
                    "native_name": (
                        str(matching_native.get("name"))
                        if matching_native is not None
                        else None
                    ),
                    "requested_name": requested_name,
                    "resolution": (
                        "animation_pane"
                        if matching_native is not None
                        else "ignored_without_native_row"
                    ),
                    f"{authority}_name": requested_name,
                    "legacy_script_marker": True,
                }
            )
            continue

        synthetic_index += 1
        marker_out = dict(marker)
        marker_out["source"] = authority
        markers_out.append(marker_out)
        effects_out.append(
            {
                "native_order": synthetic_order_base + synthetic_index,
                "shape_id": shape_id,
                "name": "Spotlight",
                "native_name": None,
                "kind": "emphasis",
                "preset_id": "",
                "preset_subtype": "",
                "trigger": authority,
                "duration_seconds": 2.4,
                "authority": authority,
                "native_present": False,
            }
        )
    return markers_out, effects_out, conflicts


def _spatial_box(item: dict[str, object]) -> tuple[float, float, float, float] | None:
    bbox = item.get("bbox")
    if not isinstance(bbox, list) or len(bbox) != 4:
        return None
    try:
        x, y, width, height = (float(value) for value in bbox)
    except (TypeError, ValueError):
        return None
    if not all(math.isfinite(value) for value in (x, y, width, height)):
        return None
    if width <= 0 or height <= 0:
        return None
    return x, y, width, height


def _spatially_ordered(
    items: Iterable[dict[str, object]],
) -> list[dict[str, object]]:
    """Order targets by horizontal bands and left-to-right vertical lanes.

    Substantial vertical overlap forms a presentation-level horizontal band.
    Inside each band, substantial horizontal overlap forms a vertical lane.
    Bands are read top-to-bottom, lanes left-to-right, and every lane is
    completed top-to-bottom before advancing right. Targets without usable
    geometry remain deterministic at the end.
    """
    valid: list[dict[str, object]] = []
    invalid: list[dict[str, object]] = []
    for item in items:
        spatial_box = _spatial_box(item)
        if spatial_box is None:
            invalid.append(item)
            continue
        copy = dict(item)
        copy["_spatial_box"] = spatial_box
        valid.append(copy)

    def stable_id(item: dict[str, object]) -> str:
        return str(item.get("shape_id") or item.get("handle") or "")

    def overlap_ratio(
        first_start: float,
        first_size: float,
        second_start: float,
        second_size: float,
    ) -> float:
        overlap = max(
            0.0,
            min(first_start + first_size, second_start + second_size)
            - max(first_start, second_start),
        )
        return overlap / min(first_size, second_size)

    def connected_components(
        members: list[dict[str, object]],
        *,
        axis: str,
        threshold: float,
    ) -> list[list[dict[str, object]]]:
        parents = list(range(len(members)))

        def find(index: int) -> int:
            while parents[index] != index:
                parents[index] = parents[parents[index]]
                index = parents[index]
            return index

        def union(left: int, right: int) -> None:
            left_root = find(left)
            right_root = find(right)
            if left_root != right_root:
                parents[right_root] = left_root

        coordinate = 1 if axis == "y" else 0
        extent = 3 if axis == "y" else 2
        for left in range(len(members)):
            left_box = members[left]["_spatial_box"]
            for right in range(left + 1, len(members)):
                right_box = members[right]["_spatial_box"]
                if overlap_ratio(
                    float(left_box[coordinate]),
                    float(left_box[extent]),
                    float(right_box[coordinate]),
                    float(right_box[extent]),
                ) >= threshold:
                    union(left, right)

        components: dict[int, list[dict[str, object]]] = {}
        for index, member in enumerate(members):
            components.setdefault(find(index), []).append(member)
        return list(components.values())

    # Transitive vertical overlap keeps staggered cards in one large band.
    bands = connected_components(valid, axis="y", threshold=0.35)
    bands.sort(
        key=lambda band: (
            min(float(item["_spatial_box"][1]) for item in band),
            min(float(item["_spatial_box"][0]) for item in band),
            min(stable_id(item) for item in band),
        )
    )

    ordered: list[dict[str, object]] = []
    for band_index, band in enumerate(bands):
        # The stricter horizontal threshold prevents narrow inter-column
        # connectors from being absorbed into a neighboring content lane.
        lanes = connected_components(band, axis="x", threshold=0.5)
        lanes.sort(
            key=lambda lane: (
                min(float(item["_spatial_box"][0]) for item in lane),
                min(float(item["_spatial_box"][1]) for item in lane),
                min(stable_id(item) for item in lane),
            )
        )
        for lane_index, lane in enumerate(lanes):
            lane.sort(
                key=lambda item: (
                    float(item["_spatial_box"][1]),
                    float(item["_spatial_box"][0]),
                    stable_id(item),
                )
            )
            for item in lane:
                item["_reading_order"] = {
                    "strategy": "horizontal_bands_vertical_lanes",
                    "band_index": band_index,
                    "lane_index": lane_index,
                    "band_direction": "top_to_bottom",
                    "lane_direction": "left_to_right",
                    "within_lane_direction": "top_to_bottom",
                }
                item.pop("_spatial_box", None)
                ordered.append(item)
    ordered.extend(
        sorted(
            invalid,
            key=stable_id,
        )
    )
    return ordered


def _notes_block_order_source(
    block: dict[str, object],
    shape: dict[str, object],
    *,
    handle_resolution: str,
    user_reordered: bool = False,
) -> str:
    """Resolve whether Notes order is explicit or was generated by the system."""
    if user_reordered:
        return "author_notes"
    stored_source = _oneline(shape.get("order_source")).lower().replace("-", "_")
    if stored_source in ORDER_SOURCES:
        return stored_source
    alt_source = _managed_alt_order_source(str(shape.get("description") or ""))
    if alt_source in ORDER_SOURCES:
        return alt_source

    # One-time migration for pre-provenance decks. The old normalizer generated
    # a Notes handle from the PowerPoint shape name, then a user could replace
    # the whole Alt Text description with plain narration. That exact legacy
    # signature is automatic ordering even though the visible Notes block now
    # contains speech.
    description = str(shape.get("description") or "")
    if (
        shape.get("script_baseline_sha256") is None
        and handle_resolution == "shape_name"
        and _oneline(block.get("handle")) == _oneline(shape.get("shape_name"))
        and bool(_legacy_plain_alt_script(description))
    ):
        return "geometry"
    return "author_notes"


def _notes_sequence_was_reordered(
    blocks: list[dict[str, object]],
    resolved: dict[int, tuple[dict[str, object], str]],
) -> bool:
    """Detect an explicit Notes reorder against stored canonical order indices."""
    indexed: list[int] = []
    for block_index in range(len(blocks)):
        shape_and_resolution = resolved.get(block_index)
        if shape_and_resolution is None:
            continue
        shape, _ = shape_and_resolution
        order_index = shape.get("order_index")
        if order_index is not None:
            indexed.append(int(order_index))
    return len(indexed) >= 2 and indexed != sorted(indexed)


def _records_in_notes_order(
    records: list[dict[str, object]],
) -> list[dict[str, object]]:
    proxies: list[dict[str, object]] = []
    for record in records:
        block = record["block"]
        shape = record["shape"]
        assert isinstance(block, dict) and isinstance(shape, dict)
        proxies.append(
            {
                **record,
                "handle": block.get("handle"),
                "shape_id": shape.get("shape_id"),
                "bbox": shape.get("bbox"),
                "order_source": record.get("order_source"),
                "notes_anchor": record.get("notes_anchor"),
            }
        )
    return _merge_partial_notes_order(proxies)


def _is_notes_anchor(block: dict[str, object]) -> bool:
    """Return whether a block was resolved from Notes in this input pass."""
    if "notes_anchor" in block:
        return bool(block.get("notes_anchor"))
    return str(block.get("order_source") or "author_notes") == "author_notes"


def _order_blocks_by_narration_policy(
    blocks: list[dict[str, object]],
    *,
    narration_order_policy: str = "geometry",
) -> list[dict[str, object]]:
    """Preserve existing Notes order and geometrically insert missing blocks.

    ``geometry`` is incremental: canonical blocks already present in Author
    Notes remain sequence anchors, while blocks absent from Notes are inserted
    around those anchors by their current canvas geometry. ``author_notes`` is
    retained as a compatibility alias for the same incremental ordering.
    """
    normalize_narration_order_policy(narration_order_policy)
    missing_geometry: list[str] = []
    for block in blocks:
        if _is_notes_anchor(block):
            continue
        script_block = block.get("block")
        transcript = (
            script_block.get("transcript")
            if isinstance(script_block, dict)
            else block.get("transcript")
        )
        if _oneline(transcript) and _spatial_box(block) is None:
            missing_geometry.append(
                str(block.get("handle") or block.get("shape_id") or "<unknown>")
            )
    if missing_geometry:
        raise ProtocolError(
            "geometry narration order requires usable top-level PowerPoint "
            "geometry for every new narrated block; missing geometry for "
            + ", ".join(repr(handle) for handle in missing_geometry)
        )

    ordered = _merge_partial_notes_order(blocks)
    for block in ordered:
        if _is_notes_anchor(block):
            continue
        block["order_source"] = "geometry"
        if "reading_model" not in block:
            block["reading_model"] = dict(
                block.pop(
                    "_reading_order",
                    {
                        "strategy": "horizontal_bands_vertical_lanes",
                        "geometry_available": False,
                    },
                )
            )
    return ordered


def _order_records_by_narration_policy(
    records: list[dict[str, object]],
    *,
    narration_order_policy: str = "geometry",
) -> list[dict[str, object]]:
    proxies: list[dict[str, object]] = []
    for record in records:
        block = record["block"]
        shape = record["shape"]
        assert isinstance(block, dict) and isinstance(shape, dict)
        proxies.append(
            {
                **record,
                "handle": block.get("handle"),
                "shape_id": shape.get("shape_id"),
                "bbox": shape.get("bbox"),
                "order_source": record.get("order_source"),
                "notes_anchor": record.get("notes_anchor"),
            }
        )
    return _order_blocks_by_narration_policy(
        proxies,
        narration_order_policy=narration_order_policy,
    )


def _fallback_handle(shape: dict[str, object], used: set[str]) -> str:
    base = (
        parse_alt_id(str(shape.get("description") or ""))
        or _oneline(shape.get("shape_name"))
        or f"Shape {shape['shape_id']}"
    )
    handle = base
    if handle in used:
        handle = f"{base} #{shape['shape_id']}"
    if handle in used:
        raise ProtocolError(f"could not create a unique handle for shape {shape['shape_id']}")
    return handle


def _first_native_order(block: dict[str, object]) -> int | None:
    orders = [
        int(effect["native_order"])
        for effect in block.get("effects") or []
        if effect.get("native_present") and effect.get("native_order") is not None
    ]
    return min(orders) if orders else None


def _merge_partial_notes_order(
    blocks: list[dict[str, object]],
) -> list[dict[str, object]]:
    """Insert new targets by band-and-lane geometry around explicit Notes."""
    notes_blocks = [block for block in blocks if _is_notes_anchor(block)]
    if not notes_blocks:
        return _spatially_ordered(blocks)

    automatic_blocks = [
        block
        for block in blocks
        if not _is_notes_anchor(block)
    ]
    if not automatic_blocks:
        return notes_blocks

    spatial = _spatially_ordered(blocks)
    ranks = {
        str(block.get("shape_id") or block.get("handle")): index
        for index, block in enumerate(spatial)
    }
    buckets: list[list[dict[str, object]]] = [
        [] for _ in range(len(notes_blocks) + 1)
    ]
    automatic_ids = {
        str(block.get("shape_id") or block.get("handle"))
        for block in automatic_blocks
    }
    for block in spatial:
        block_id = str(block.get("shape_id") or block.get("handle"))
        if block_id not in automatic_ids:
            continue
        block["reading_model"] = dict(
            block.pop(
                "_reading_order",
                {
                    "strategy": "horizontal_bands_vertical_lanes",
                    "geometry_available": False,
                },
            )
        )
        block_rank = ranks[block_id]
        insertion = len(notes_blocks)
        for note_index, note_block in enumerate(notes_blocks):
            note_rank = ranks[
                str(note_block.get("shape_id") or note_block.get("handle"))
            ]
            if block_rank < note_rank:
                insertion = note_index
                break
        buckets[insertion].append(block)

    merged: list[dict[str, object]] = []
    for note_index, note_block in enumerate(notes_blocks):
        merged.extend(buckets[note_index])
        merged.append(note_block)
    merged.extend(buckets[-1])
    return merged


def extract_protocol(
    pptx_path: Path,
    *,
    section_ids: Iterable[str] | None = None,
    ids_from_script: Path | None = None,
    narration_order_policy: str = "geometry",
) -> dict[str, object]:
    """Extract narration and animation intent from an editable PPTX.

    Canonical Author Notes are authoritative when present. Otherwise explicit
    Shape Alt Text ``Script:`` fields provide narration. Existing canonical
    Notes blocks keep their relative spoken order; current canvas geometry
    places only blocks absent from Notes around those anchors. Callers may keep
    selecting ``author_notes`` for the explicit Notes authority path.
    """
    pptx_path = pptx_path.resolve()
    narration_order_policy = normalize_narration_order_policy(
        narration_order_policy
    )
    explicit_ids = list(section_ids) if section_ids is not None else None
    if ids_from_script is not None:
        if explicit_ids is not None:
            raise ProtocolError("section_ids and ids_from_script are mutually exclusive")
        explicit_ids = _section_ids_from_script(ids_from_script.resolve())

    with ZipFile(pptx_path) as archive:
        slide_refs = presentation_slides(archive)
        slide_width, slide_height = presentation_size(archive)
        if explicit_ids is not None and len(explicit_ids) != len(slide_refs):
            raise ProtocolError(
                f"section ID count {len(explicit_ids)} != PPTX slide count {len(slide_refs)}"
            )
        slides: list[dict[str, object]] = []
        represented_effect_count = 0
        represented_native_effect_count = 0
        protocol_sources: set[str] = set()
        for ref in slide_refs:
            index = int(ref["index"])
            stable_id = str(ref["stable_id"])
            section_id = (
                str(explicit_ids[index - 1])
                if explicit_ids is not None
                else f"slide-{stable_id}"
            )
            slide_root = etree.fromstring(archive.read(str(ref["part"])))
            effects = _native_effects(slide_root)
            by_shape, by_handle = _shape_maps(
                slide_root,
                slide_width=slide_width,
                slide_height=slide_height,
            )
            notes = _notes_text_optional(archive, str(ref["part"]))
            notes_blocks = _canonical_notes_blocks(notes)

            effects_by_shape: dict[str, list[dict[str, object]]] = OrderedDict()
            for effect in effects:
                shape_id = str(effect["shape_id"])
                if shape_id not in by_shape:
                    raise ProtocolError(
                        f"slide {index} native animation targets missing shape id {shape_id!r}"
                    )
                effects_by_shape.setdefault(shape_id, []).append(effect)
            ordered_shape_ids = list(effects_by_shape)

            resolved_blocks: list[dict[str, object]] = []
            represented_shape_ids: set[str] = set()
            effect_conflicts: list[dict[str, object]] = []
            ignored_stale_notes: list[dict[str, object]] = []
            used_handles: set[str] = set()
            alt_by_shape: dict[str, dict[str, object]] = {}
            for shape_id, shape in by_shape.items():
                if not shape.get("top_level"):
                    continue
                alt = parse_alt_protocol(str(shape.get("description") or ""))
                if alt is not None:
                    alt_by_shape[shape_id] = alt

            def append_block(
                block: dict[str, object],
                shape: dict[str, object],
                *,
                source: str,
                handle_resolution: str,
                order_source: str,
                notes_anchor: bool = False,
            ) -> None:
                _hydrate_block_semantics(block, shape)
                shape_id = str(shape["shape_id"])
                if not shape.get("top_level"):
                    raise ProtocolError(
                        f"slide {index} handle {block['handle']!r} targets a nested shape; "
                        "editable video targets must be top-level PowerPoint elements"
                    )
                if shape_id in represented_shape_ids:
                    raise ProtocolError(
                        f"slide {index} shape {shape_id} is represented by more than one script block"
                    )
                handle = str(block["handle"])
                if handle in used_handles:
                    raise ProtocolError(f"slide {index} has duplicate resolved handle {handle!r}")
                represented_shape_ids.add(shape_id)
                used_handles.add(handle)
                native = effects_by_shape.get(shape_id) or []
                animation_names = [str(name) for name in block.get("animation_names") or []]
                markers, effective_effects, conflicts = _merge_effect_intents(
                    shape_id=shape_id,
                    handle=handle,
                    native_effects=native,
                    markers=list(block.get("markers") or []),
                    animation_names=animation_names,
                    authority=source,
                    synthetic_order_base=1_000_000 + index * 10_000 + len(resolved_blocks) * 100,
                )
                effect_conflicts.extend(conflicts)
                resolved_blocks.append(
                    {
                        **block,
                        "markers": markers,
                        "shape_id": shape_id,
                        "shape_name": shape["shape_name"],
                        "visible_texts": list(shape.get("visible_texts") or []),
                        "bbox": shape["bbox"],
                        "script_source": source,
                        "order_source": order_source,
                        "notes_anchor": notes_anchor,
                        "handle_resolution": handle_resolution,
                        "alt_text_handle": parse_alt_id(str(shape.get("description") or "")),
                        "effects": effective_effects,
                    }
                )

            if notes_blocks:
                protocol_sources.add("author_notes")
                shapes_by_name: dict[str, list[dict[str, object]]] = {}
                for candidate in by_shape.values():
                    if candidate.get("top_level"):
                        shapes_by_name.setdefault(
                            _oneline(candidate.get("shape_name")), []
                        ).append(candidate)
                pre_resolved: dict[int, tuple[dict[str, object], str]] = {}
                claimed_shape_ids: set[str] = set()
                unresolved_block_indexes: list[int] = []
                for block_index, block in enumerate(notes_blocks):
                    handle = str(block["handle"])
                    shape = by_handle.get(handle)
                    handle_resolution = "alt_text_handle"
                    if shape is None:
                        name_matches = shapes_by_name.get(_oneline(handle)) or []
                        if len(name_matches) == 1:
                            shape = name_matches[0]
                            handle_resolution = "shape_name"
                    if shape is None:
                        unresolved_block_indexes.append(block_index)
                        continue
                    shape_id = str(shape["shape_id"])
                    if shape_id in claimed_shape_ids:
                        raise ProtocolError(
                            f"slide {index} Notes handles resolve more than once to shape {shape_id!r}"
                        )
                    claimed_shape_ids.add(shape_id)
                    pre_resolved[block_index] = (shape, handle_resolution)

                ignored_stale_notes.extend(
                    {
                        "block_index": block_index,
                        "handle": str(notes_blocks[block_index]["handle"]),
                        "resolution": "ignored_missing_shape",
                    }
                    for block_index in unresolved_block_indexes
                )
                notes_reordered = _notes_sequence_was_reordered(
                    notes_blocks,
                    pre_resolved,
                )

                for block_index, block in enumerate(notes_blocks):
                    if block_index not in pre_resolved:
                        continue
                    shape, handle_resolution = pre_resolved[block_index]
                    append_block(
                        block,
                        shape,
                        source="author_notes",
                        handle_resolution=handle_resolution,
                        order_source=(
                            "author_notes"
                            if narration_order_policy == "author_notes"
                            else _notes_block_order_source(
                                block,
                                shape,
                                handle_resolution=handle_resolution,
                                user_reordered=notes_reordered,
                            )
                        ),
                        notes_anchor=True,
                    )

            for shape_id in ordered_shape_ids:
                if shape_id in represented_shape_ids:
                    continue
                shape = by_shape[shape_id]
                alt = alt_by_shape.get(shape_id)
                if alt is not None:
                    protocol_sources.add("alt_text")
                    append_block(
                        alt,
                        shape,
                        source="alt_text",
                        handle_resolution="alt_text_handle",
                        order_source="geometry",
                    )
                    continue
                handle = _fallback_handle(shape, used_handles)
                append_block(
                    {
                        "handle": handle,
                        "semantic": "",
                        "semantic_concise": "",
                        "semantic_detailed": "",
                        "raw_transcript": "",
                        "transcript": "",
                        "markers": [],
                        "animation_names": [],
                    },
                    shape,
                    source="animation_pane",
                    handle_resolution="animation_pane",
                    order_source="geometry",
                )

            remaining_alt_shapes = _spatially_ordered(
                [
                    by_shape[shape_id]
                    for shape_id in alt_by_shape
                    if shape_id not in represented_shape_ids
                ]
            )
            for shape in remaining_alt_shapes:
                shape_id = str(shape["shape_id"])
                protocol_sources.add("alt_text")
                append_block(
                    alt_by_shape[shape_id],
                    shape,
                    source="alt_text",
                    handle_resolution="geometry_order",
                    order_source="geometry",
                )

            resolved_blocks = _order_blocks_by_narration_policy(
                resolved_blocks,
                narration_order_policy=narration_order_policy,
            )

            transcript = _oneline(" ".join(str(block["transcript"]) for block in resolved_blocks))
            slide_effect_count = sum(len(block["effects"]) for block in resolved_blocks)
            if not transcript and not slide_effect_count:
                raise ProtocolError(
                    f"slide {index} has neither narration nor a supported native animation"
                )
            represented_effect_count += slide_effect_count
            represented_native_effect_count += len(effects)
            slide_sources = list(
                OrderedDict(
                    (str(block["script_source"]), None)
                    for block in resolved_blocks
                    if str(block.get("transcript") or "")
                )
            )
            slides.append(
                {
                    "index": index,
                    "stable_slide_id": stable_id,
                    "section_id": section_id,
                    "slide_part": ref["part"],
                    "notes_part": _notes_part_optional(archive, str(ref["part"])),
                    "transcript": transcript,
                    "script_sources": slide_sources,
                    "block_count": len(resolved_blocks),
                    "effect_count": slide_effect_count,
                    "native_effect_count": len(effects),
                    "animation_duration_seconds": round(
                        max(
                            (
                                float(effect.get("pane_end_seconds") or 0.0)
                                for effect in effects
                            ),
                            default=0.0,
                        ),
                        3,
                    ),
                    "effect_conflicts": effect_conflicts,
                    "ignored_stale_notes": ignored_stale_notes,
                    "narration_order_policy": narration_order_policy,
                    "blocks": resolved_blocks,
                }
            )

    return {
        "schema_version": PROTOCOL_SCHEMA_VERSION,
        "created_at": _utc_now(),
        "source_pptx": str(pptx_path),
        "source_sha256": file_sha256(pptx_path),
        "slide_count": len(slides),
        "slide_size_emu": [slide_width, slide_height],
        "effect_count": represented_effect_count,
        "native_effect_count": represented_native_effect_count,
        "narration_order_policy": narration_order_policy,
        "script_sources": sorted(protocol_sources),
        "slides": slides,
    }


def analyze_animation_order_conflicts(
    protocol: dict[str, object],
    authority_report: dict[str, object],
) -> dict[str, object]:
    """Find ambiguous narration-order versus Animation Pane inversions.

    Only blocks absent from canonical Notes and inserted during this pass are
    decision candidates. Existing Notes/Pane differences do not repeatedly
    interrupt later renders. Empty decorative targets, emphasis-only targets,
    and effects that start simultaneously are excluded from conflict pairs.
    """
    narration_order_policy = normalize_narration_order_policy(
        protocol.get("narration_order_policy") or "author_notes"
    )
    slide_size_raw = protocol.get("slide_size_emu")
    slide_size_emu: list[int] | None = None
    if isinstance(slide_size_raw, (list, tuple)) and len(slide_size_raw) == 2:
        try:
            slide_width = int(slide_size_raw[0])
            slide_height = int(slide_size_raw[1])
        except (TypeError, ValueError):
            pass
        else:
            if slide_width > 0 and slide_height > 0:
                slide_size_emu = [slide_width, slide_height]
    inserted_by_slide: dict[int, set[str]] = {}
    inserted_metadata: dict[tuple[int, str], dict[str, object]] = {}
    for slide in authority_report.get("slides") or []:
        slide_index = int(slide.get("index") or 0)
        for block in slide.get("inserted_blocks") or []:
            shape_id = str(block.get("shape_id") or "")
            if not shape_id:
                continue
            inserted_by_slide.setdefault(slide_index, set()).add(shape_id)
            inserted_metadata[(slide_index, shape_id)] = dict(block)

    slides_out: list[dict[str, object]] = []
    conflict_count = 0
    inversion_pairs: set[tuple[int, str, str]] = set()
    fingerprint_rows: list[dict[str, object]] = []
    for slide in protocol.get("slides") or []:
        slide_index = int(slide.get("index") or 0)
        candidates = inserted_by_slide.get(slide_index) or set()
        rows: list[dict[str, object]] = []
        for block_index, block in enumerate(slide.get("blocks") or []):
            transcript = _oneline(block.get("transcript"))
            if not transcript:
                continue
            entrance_effects = [
                effect
                for effect in block.get("effects") or []
                if effect.get("native_present")
                and str(effect.get("kind") or "entrance") == "entrance"
            ]
            if not entrance_effects:
                continue
            first = min(
                entrance_effects,
                key=lambda effect: int(effect.get("native_order") or 0),
            )
            block_shape_id = str(block.get("shape_id") or "")
            block_metadata = inserted_metadata.get((slide_index, block_shape_id)) or {}
            is_inserted_candidate = block_shape_id in candidates
            semantic_concise, semantic_detailed = _semantic_values(
                semantic_concise=(
                    block.get("semantic_concise")
                    or block_metadata.get("semantic_concise")
                ),
                semantic=block.get("semantic"),
                semantic_detailed=(
                    block.get("semantic_detailed")
                    or block_metadata.get("semantic_detailed")
                ),
                transcript=transcript,
                visible_texts=(
                    block.get("visible_texts")
                    or block_metadata.get("visible_texts")
                    or []
                ),
                shape_name=block.get("shape_name"),
                handle=block.get("handle"),
            )
            rows.append(
                {
                    "shape_id": block_shape_id,
                    "handle": str(block.get("handle") or ""),
                    "shape_name": str(block.get("shape_name") or ""),
                    "visible_texts": list(block.get("visible_texts") or []),
                    "transcript": transcript,
                    "semantic_concise": semantic_concise,
                    "semantic_detailed": semantic_detailed,
                    "bbox": block.get("bbox"),
                    "order_source": str(block.get("order_source") or ""),
                    "reading_model": block.get("reading_model"),
                    "effective_narration_rank": block_index + 1,
                    "native_order": int(first.get("native_order") or 0),
                    "pane_start_seconds": round(
                        float(first.get("pane_start_seconds") or 0.0), 3
                    ),
                    "simultaneous_group": int(
                        first.get("simultaneous_group")
                        or first.get("native_order")
                        or 0
                    ),
                    "trigger": str(first.get("trigger") or ""),
                    "effect_name": str(first.get("name") or ""),
                    "is_inserted_candidate": is_inserted_candidate,
                    "is_conflict_candidate": is_inserted_candidate,
                }
            )

        pane_rows = sorted(
            rows,
            key=lambda row: (
                float(row["pane_start_seconds"]),
                int(row["native_order"]),
                str(row["shape_id"]),
            ),
        )
        pane_rank = {
            str(row["shape_id"]): rank
            for rank, row in enumerate(pane_rows, start=1)
        }
        geometry_rows = _spatially_ordered(rows)
        geometry_rank = {
            str(row["shape_id"]): rank
            for rank, row in enumerate(geometry_rows, start=1)
        }
        geometry_model = {
            str(row["shape_id"]): dict(row.get("_reading_order") or {})
            for row in geometry_rows
        }
        slide_conflicts: list[dict[str, object]] = []
        for candidate in rows:
            shape_id = str(candidate["shape_id"])
            if not candidate["is_conflict_candidate"]:
                continue
            inversions: list[dict[str, object]] = []
            for anchor in rows:
                anchor_shape_id = str(anchor["shape_id"])
                if anchor_shape_id == shape_id:
                    continue
                if int(anchor["simultaneous_group"]) == int(
                    candidate["simultaneous_group"]
                ):
                    continue
                if abs(
                    float(anchor["pane_start_seconds"])
                    - float(candidate["pane_start_seconds"])
                ) <= 0.002:
                    continue
                geometry_relation = geometry_rank[shape_id] < geometry_rank[
                    anchor_shape_id
                ]
                pane_relation = (
                    float(candidate["pane_start_seconds"]),
                    int(candidate["native_order"]),
                ) < (
                    float(anchor["pane_start_seconds"]),
                    int(anchor["native_order"]),
                )
                if geometry_relation == pane_relation:
                    continue
                inversions.append(
                    {
                        "shape_id": anchor_shape_id,
                        "handle": str(anchor["handle"]),
                        "visible_texts": list(anchor["visible_texts"]),
                        "semantic_concise": str(anchor["semantic_concise"]),
                        "semantic_detailed": str(anchor["semantic_detailed"]),
                        "geometry_rank": geometry_rank[anchor_shape_id],
                        "reading_rank": geometry_rank[anchor_shape_id],
                        "pane_rank": pane_rank[anchor_shape_id],
                        "pane_start_seconds": anchor["pane_start_seconds"],
                    }
                )
            if not inversions:
                continue
            metadata = inserted_metadata.get((slide_index, shape_id)) or {}
            geometry_position = geometry_rank[shape_id]
            geometry_predecessor_row = (
                geometry_rows[geometry_position - 2]
                if geometry_position > 1
                else None
            )
            geometry_successor_row = (
                geometry_rows[geometry_position]
                if geometry_position < len(geometry_rows)
                else None
            )
            conflict = {
                **candidate,
                "visible_texts": list(
                    candidate.get("visible_texts")
                    or metadata.get("visible_texts")
                    or []
                ),
                "geometry_rank": geometry_position,
                "reading_rank": geometry_position,
                "geometry_model": geometry_model.get(shape_id) or None,
                "pane_rank": pane_rank[shape_id],
                "effective_narration_rank": int(
                    candidate.get("effective_narration_rank") or 0
                ),
                "geometry_predecessor": (
                    geometry_predecessor_row["handle"]
                    if geometry_predecessor_row is not None
                    else None
                ),
                "geometry_successor": (
                    geometry_successor_row["handle"]
                    if geometry_successor_row is not None
                    else None
                ),
                "reading_predecessor": (
                    geometry_predecessor_row["handle"]
                    if geometry_predecessor_row is not None
                    else None
                ),
                "reading_successor": (
                    geometry_successor_row["handle"]
                    if geometry_successor_row is not None
                    else None
                ),
                "pane_predecessor": (
                    pane_rows[pane_rank[shape_id] - 2]["handle"]
                    if pane_rank[shape_id] > 1
                    else None
                ),
                "pane_successor": (
                    pane_rows[pane_rank[shape_id]]["handle"]
                    if pane_rank[shape_id] < len(pane_rows)
                    else None
                ),
                "inversion_count": len(inversions),
                "inversions": inversions,
                "recommended_pane_move": {
                    "shape_id": shape_id,
                    "handle": str(candidate["handle"]),
                    "current_native_order": int(candidate["native_order"]),
                    "place_after": (
                        {
                            "shape_id": str(geometry_predecessor_row["shape_id"]),
                            "handle": str(geometry_predecessor_row["handle"]),
                            "native_order": int(
                                geometry_predecessor_row["native_order"]
                            ),
                            "visible_texts": list(
                                geometry_predecessor_row["visible_texts"]
                            ),
                            "semantic_concise": str(
                                geometry_predecessor_row["semantic_concise"]
                            ),
                            "semantic_detailed": str(
                                geometry_predecessor_row["semantic_detailed"]
                            ),
                        }
                        if geometry_predecessor_row is not None
                        else None
                    ),
                    "place_before": (
                        {
                            "shape_id": str(geometry_successor_row["shape_id"]),
                            "handle": str(geometry_successor_row["handle"]),
                            "native_order": int(geometry_successor_row["native_order"]),
                            "visible_texts": list(
                                geometry_successor_row["visible_texts"]
                            ),
                            "semantic_concise": str(
                                geometry_successor_row["semantic_concise"]
                            ),
                            "semantic_detailed": str(
                                geometry_successor_row["semantic_detailed"]
                            ),
                        }
                        if geometry_successor_row is not None
                        else None
                    ),
                },
            }
            slide_conflicts.append(conflict)
            conflict_count += 1
            for inversion in inversions:
                inversion_pairs.add(
                    (
                        slide_index,
                        *sorted((shape_id, str(inversion["shape_id"]))),
                    )
                )
            fingerprint_rows.append(
                {
                    "slide_index": slide_index,
                    "shape_id": shape_id,
                    # Keep the v1 fingerprint field name for strict-QA
                    # compatibility. Its value is now the pure geometry rank.
                    "reading_rank": conflict["geometry_rank"],
                    "pane_rank": conflict["pane_rank"],
                    "inversions": [item["shape_id"] for item in inversions],
                }
            )
        if slide_conflicts:
            slide_out: dict[str, object] = {
                "index": slide_index,
                "section_id": str(slide.get("section_id") or ""),
                "geometry_sequence": [
                        {
                            "rank": rank,
                            "shape_id": row["shape_id"],
                            "handle": row["handle"],
                            "visible_texts": row["visible_texts"],
                            "semantic_concise": row["semantic_concise"],
                            "semantic_detailed": row["semantic_detailed"],
                            "bbox": (
                                list(row["bbox"])
                                if isinstance(row.get("bbox"), (list, tuple))
                                else None
                            ),
                        }
                        for rank, row in enumerate(geometry_rows, start=1)
                ],
                # Backward-compatible alias. This is the same pure geometry
                # sequence, not the Notes-anchor narration sequence.
                "reading_sequence": [
                        {
                            "rank": rank,
                            "shape_id": row["shape_id"],
                            "handle": row["handle"],
                            "visible_texts": row["visible_texts"],
                            "semantic_concise": row["semantic_concise"],
                            "semantic_detailed": row["semantic_detailed"],
                            "bbox": (
                                list(row["bbox"])
                                if isinstance(row.get("bbox"), (list, tuple))
                                else None
                            ),
                        }
                        for rank, row in enumerate(geometry_rows, start=1)
                ],
                "animation_pane_sequence": [
                        {
                            "rank": rank,
                            "shape_id": row["shape_id"],
                            "handle": row["handle"],
                            "visible_texts": row["visible_texts"],
                            "semantic_concise": row["semantic_concise"],
                            "semantic_detailed": row["semantic_detailed"],
                            "bbox": (
                                list(row["bbox"])
                                if isinstance(row.get("bbox"), (list, tuple))
                                else None
                            ),
                            "native_order": row["native_order"],
                            "pane_start_seconds": row["pane_start_seconds"],
                        }
                        for rank, row in enumerate(pane_rows, start=1)
                ],
                "conflicts": slide_conflicts,
            }
            if slide_size_emu is not None:
                slide_out["slide_size_emu"] = list(slide_size_emu)
            slides_out.append(slide_out)

    conflict_set_payload = {
        "conflicts": fingerprint_rows,
    }
    conflict_set_fingerprint = hashlib.sha256(
        json.dumps(
            conflict_set_payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    fingerprint_payload = {
        "source_sha256": str(protocol.get("source_sha256") or ""),
        "conflict_set_fingerprint": conflict_set_fingerprint,
    }
    fingerprint = hashlib.sha256(
        json.dumps(
            fingerprint_payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    return {
        "schema_version": "paper2video_animation_order_authority.v1",
        "created_at": _utc_now(),
        "source_pptx": protocol.get("source_pptx"),
        "source_sha256": protocol.get("source_sha256"),
        "slide_size_emu": (
            list(slide_size_emu) if slide_size_emu is not None else None
        ),
        "narration_order_policy": narration_order_policy,
        "status": "decision_required" if conflict_count else "no_conflict",
        "selected_policy": None,
        "selection_source": None,
        "conflict_set_fingerprint": conflict_set_fingerprint,
        "conflict_fingerprint": fingerprint,
        "conflict_count": conflict_count,
        "inversion_pair_count": len(inversion_pairs),
        "slides": slides_out,
    }


def apply_confirmed_animation_pane_order(
    protocol: dict[str, object],
    *,
    slide_indices: Iterable[int] | None = None,
) -> tuple[dict[str, object], dict[str, object]]:
    """Serialize narrated native-entrance blocks in Animation Pane order.

    Geometry remains the default preflight reading policy. Once the user has
    confirmed that a conflicting Animation Pane is intentional, however, the
    renderer must make the same choice on its narration clock. Only narrated
    blocks with a native entrance participate in this reorder. Their existing
    block slots are filled in Pane order, so narrated blocks without an
    entrance and silent/decorative blocks keep their positions.

    When ``slide_indices`` is provided, only those slides are changed. This is
    used by per-slide conflict resolution so a Pane choice on one conflict page
    cannot rewrite block order or provenance on any other page.

    The returned protocol is a deep copy. It is suitable for Notes/OOXML
    writeback and for downstream TTS scheduling; the input decision protocol
    remains unchanged for conflict-fingerprint verification.
    """
    updated = deepcopy(protocol)
    selected_slide_indices = (
        None if slide_indices is None else {int(index) for index in slide_indices}
    )
    available_slide_indices = {
        int(slide.get("index") or 0) for slide in updated.get("slides") or []
    }
    if selected_slide_indices is not None:
        unknown_slide_indices = selected_slide_indices - available_slide_indices
        if unknown_slide_indices:
            raise ProtocolError(
                "confirmed Animation Pane order references unknown slide indices: "
                + ", ".join(str(index) for index in sorted(unknown_slide_indices))
            )
    slide_reports: list[dict[str, object]] = []
    changed_slide_count = 0
    moved_block_count = 0

    for slide in updated.get("slides") or []:
        slide_index = int(slide.get("index") or 0)
        if (
            selected_slide_indices is not None
            and slide_index not in selected_slide_indices
        ):
            continue
        blocks = list(slide.get("blocks") or [])
        candidates: list[dict[str, object]] = []
        for block_index, block in enumerate(blocks):
            if not _oneline(block.get("transcript")):
                continue
            entrances = [
                effect
                for effect in block.get("effects") or []
                if effect.get("native_present")
                and str(effect.get("kind") or "entrance") == "entrance"
            ]
            if not entrances:
                continue
            first = min(
                entrances,
                key=lambda effect: (
                    float(effect.get("pane_start_seconds") or 0.0),
                    int(effect.get("native_order") or 0),
                ),
            )
            candidates.append(
                {
                    "slot": block_index,
                    "original_rank": len(candidates) + 1,
                    "block": block,
                    "shape_id": str(block.get("shape_id") or ""),
                    "handle": str(block.get("handle") or ""),
                    "pane_start_seconds": round(
                        float(first.get("pane_start_seconds") or 0.0), 3
                    ),
                    "native_order": int(first.get("native_order") or 0),
                    "simultaneous_group": int(
                        first.get("simultaneous_group")
                        or first.get("native_order")
                        or 0
                    ),
                }
            )

        pane_order = sorted(
            candidates,
            key=lambda candidate: (
                float(candidate["pane_start_seconds"]),
                int(candidate["native_order"]),
                int(candidate["original_rank"]),
                str(candidate["shape_id"]),
            ),
        )
        candidate_slots = [int(candidate["slot"]) for candidate in candidates]
        before_shape_ids = [str(candidate["shape_id"]) for candidate in candidates]
        after_shape_ids = [str(candidate["shape_id"]) for candidate in pane_order]

        moved_blocks: list[dict[str, object]] = []
        for pane_rank, (slot, candidate) in enumerate(
            zip(candidate_slots, pane_order, strict=True),
            start=1,
        ):
            block = candidate["block"]
            assert isinstance(block, dict)
            blocks[slot] = block
            original_slot = int(candidate["slot"])
            if original_slot != slot:
                moved_blocks.append(
                    {
                        "shape_id": candidate["shape_id"],
                        "handle": candidate["handle"],
                        "from_block_index": original_slot + 1,
                        "to_block_index": slot + 1,
                    }
                )
            block["order_source"] = "animation_pane"
            block["reading_model"] = {
                "strategy": "confirmed_animation_pane",
                "pane_rank": pane_rank,
                "pane_start_seconds": candidate["pane_start_seconds"],
                "native_order": candidate["native_order"],
                "simultaneous_group": candidate["simultaneous_group"],
            }

        changed = before_shape_ids != after_shape_ids
        if changed:
            changed_slide_count += 1
        moved_block_count += len(moved_blocks)
        slide["blocks"] = blocks
        slide["block_count"] = len(blocks)
        slide["transcript"] = _oneline(
            " ".join(str(block.get("transcript") or "") for block in blocks)
        )
        slide["render_narration_order"] = "animation_pane"
        slide_reports.append(
            {
                "index": slide_index,
                "section_id": str(slide.get("section_id") or ""),
                "changed": changed,
                "eligible_block_count": len(candidates),
                "before_shape_ids": before_shape_ids,
                "after_shape_ids": after_shape_ids,
                "moved_blocks": moved_blocks,
            }
        )

    updated["render_narration_order"] = "animation_pane"
    application = {
        "schema_version": "paper2video_animation_pane_narration_order.v1",
        "policy": "animation-pane",
        "scope": "narrated_native_entrance_blocks",
        "selected_slide_indices": (
            sorted(selected_slide_indices)
            if selected_slide_indices is not None
            else None
        ),
        "changed_slide_count": changed_slide_count,
        "moved_block_count": moved_block_count,
        "slides": slide_reports,
    }
    return updated, application


def script_from_protocol(
    protocol: dict[str, object],
    *,
    voice: str | None = None,
) -> dict[str, object]:
    sections = []
    for slide in protocol.get("slides") or []:
        sections.append(
            {
                "id": slide["section_id"],
                "heading": slide["section_id"],
                "text": slide["transcript"],
                "duration_seconds": max(
                    1.0,
                    round(float(slide.get("animation_duration_seconds") or 0.0) + 0.35, 3),
                ),
            }
        )
    payload: dict[str, object] = {"provider": "edge", "sections": sections}
    if voice:
        payload["edge_voice"] = voice
    return payload


def _native_rows_from_block(block: dict[str, object]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for effect in block.get("effects") or []:
        if not effect.get("native_present"):
            continue
        native_name = str(effect.get("native_name") or effect.get("name") or "")
        rows.append(
            {
                **effect,
                "name": native_name,
                "native_name": native_name,
                "authority": "animation_pane",
                "native_present": True,
            }
        )
    rows.sort(key=lambda effect: int(effect.get("native_order") or 0))
    return rows


def apply_user_script(
    protocol: dict[str, object],
    script_json: Path,
) -> tuple[dict[str, object], dict[str, object]]:
    """Apply an explicit user-edited script to an extracted PPTX protocol.

    Standard section-level ``text`` replaces slide narration and preserves all
    resolved effects at deterministic block-start timing. For precise per-item
    timing, a section may provide ``elements`` entries with ``handle`` and
    marker-bearing ``script`` fields.
    """
    script_json = script_json.resolve()
    try:
        payload = json.loads(script_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"could not read user script {script_json}: {exc}") from exc
    sections = payload.get("sections") or []
    slides = protocol.get("slides") or []
    if len(sections) != len(slides):
        raise ProtocolError(
            f"user script section count {len(sections)} != PPTX slide count {len(slides)}"
        )

    updated = deepcopy(protocol)
    report_slides: list[dict[str, object]] = []
    for slide, section in zip(updated.get("slides") or [], sections):
        section_id = str(section.get("id") or slide["section_id"])
        if section_id != str(slide["section_id"]):
            raise ProtocolError(
                f"user script section {section_id!r} != PPTX section {slide['section_id']!r}"
            )
        blocks = list(slide.get("blocks") or [])
        has_element_override = "elements" in section or "blocks" in section
        elements = section.get("elements") or section.get("blocks") or []
        overridden_handles: list[str] = []
        mode = "elements" if has_element_override else "section_text"
        if has_element_override:
            by_handle = {str(block["handle"]): block for block in blocks}
            seen: set[str] = set()
            for element in elements:
                handle = _oneline(element.get("handle"))
                if not handle or handle in seen:
                    raise ProtocolError(
                        f"user script section {section_id!r} has a missing or duplicate element handle"
                    )
                seen.add(handle)
                block = by_handle.get(handle)
                if block is None:
                    raise ProtocolError(
                        f"user script section {section_id!r} references unknown handle {handle!r}"
                    )
                legacy_raw_script = str(
                    element.get("script") or element.get("text") or ""
                ).strip()
                legacy_animation_names = [
                    _oneline(match.group(1))
                    for match in MARKER_RE.finditer(legacy_raw_script)
                    if _oneline(match.group(1)) != "Spotlight"
                ]
                raw_script = canonicalize_user_script(legacy_raw_script)
                transcript, markers = parse_marked_transcript(raw_script)
                if not transcript:
                    raise ProtocolError(
                        f"user script section {section_id!r} handle {handle!r} has empty narration"
                    )
                block["raw_transcript"] = raw_script
                block["transcript"] = transcript
                block["script_source"] = "user_script"
                if markers or legacy_animation_names:
                    merged_markers, merged_effects, conflicts = _merge_effect_intents(
                        shape_id=str(block["shape_id"]),
                        handle=handle,
                        native_effects=_native_rows_from_block(block),
                        markers=markers,
                        animation_names=legacy_animation_names,
                        authority="user_script",
                        synthetic_order_base=(
                            2_000_000 + int(slide["index"]) * 10_000 + len(overridden_handles) * 100
                        ),
                    )
                    block["markers"] = merged_markers
                    block["effects"] = merged_effects
                    slide.setdefault("effect_conflicts", []).extend(conflicts)
                else:
                    block["markers"] = [
                        _default_marker(str(effect["name"]), "user_script_default")
                        for effect in block.get("effects") or []
                    ]
                overridden_handles.append(handle)
        else:
            text = _oneline(section.get("text"))
            if not text:
                raise ProtocolError(f"user script section {section_id!r} has empty text")
            clean_text, section_markers = parse_marked_transcript(text)
            if section_markers:
                raise ProtocolError(
                    f"user script section {section_id!r} has animation markers in slide-level "
                    "text; use handle-addressed elements for precise marker targets"
                )
            text = clean_text
            if not blocks:
                raise ProtocolError(f"PPTX section {section_id!r} has no protocol blocks")
            for block_index, block in enumerate(blocks):
                block["raw_transcript"] = text if block_index == 0 else ""
                block["transcript"] = text if block_index == 0 else ""
                block["script_source"] = "user_script"
                block["markers"] = [
                    _default_marker(str(effect["name"]), "user_script_default")
                    for effect in block.get("effects") or []
                ]
            overridden_handles = [str(blocks[0]["handle"])]

        transcript = _oneline(" ".join(str(block.get("transcript") or "") for block in blocks))
        slide_effect_count = sum(len(block.get("effects") or []) for block in blocks)
        if not transcript and not slide_effect_count:
            raise ProtocolError(f"user script left PPTX section {section_id!r} empty")
        slide["transcript"] = transcript
        slide["script_sources"] = list(
            OrderedDict(
                (str(block["script_source"]), None)
                for block in blocks
                if str(block.get("transcript") or "")
            )
        )
        slide["effect_count"] = slide_effect_count
        report_slides.append(
            {
                "index": int(slide["index"]),
                "id": section_id,
                "mode": mode,
                "overridden_handles": overridden_handles,
                "transcript": transcript,
            }
        )

    updated["script_sources"] = sorted(
        {
            str(source)
            for slide in updated.get("slides") or []
            for source in slide.get("script_sources") or []
        }
    )
    updated["effect_count"] = sum(
        int(slide.get("effect_count") or 0) for slide in updated.get("slides") or []
    )
    updated["user_script"] = str(script_json)
    report = {
        "schema_version": "paper2video_user_script_authority.v1",
        "created_at": _utc_now(),
        "script_json": str(script_json),
        "resolution": "user_script",
        "slide_count": len(report_slides),
        "slides": report_slides,
    }
    return updated, report


def write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _ensure_notes_parts(source_pptx: Path, directory: Path) -> tuple[Path, Path | None]:
    """Create missing notes parts through python-pptx while preserving slide XML."""
    with ZipFile(source_pptx) as archive:
        missing = any(
            _notes_part_optional(archive, str(ref["part"])) is None
            for ref in presentation_slides(archive)
        )
    if not missing:
        return source_pptx, None
    try:
        from pptx import Presentation
    except ImportError as exc:  # pragma: no cover - installer provides python-pptx
        raise ProtocolError(
            "python-pptx is required to create missing Author Notes parts"
        ) from exc
    with tempfile.NamedTemporaryFile(
        prefix=source_pptx.stem + ".notes.",
        suffix=".pptx",
        dir=directory,
        delete=False,
    ) as temporary:
        seeded = Path(temporary.name)
    presentation = Presentation(source_pptx)
    for slide in presentation.slides:
        _ = slide.notes_slide
    presentation.save(seeded)
    return seeded, seeded


def _script_fields_from(
    notes_block: dict[str, object],
    script_block: dict[str, object],
) -> dict[str, object]:
    """Keep the Notes handle/locator while accepting another script version."""
    resolved = dict(notes_block)
    for field in ("raw_transcript", "transcript", "markers", "animation_names"):
        resolved[field] = script_block.get(field, [] if field.endswith("s") else "")
    if not _oneline(resolved.get("semantic")):
        resolved["semantic"] = script_block.get("semantic") or ""
    if not _oneline(resolved.get("semantic_concise")):
        resolved["semantic_concise"] = (
            script_block.get("semantic_concise")
            or script_block.get("semantic")
            or ""
        )
    if not _oneline(resolved.get("semantic_detailed")):
        resolved["semantic_detailed"] = script_block.get("semantic_detailed") or ""
    return resolved


def _resolve_notes_alt_script(
    notes_block: dict[str, object],
    alt_block: dict[str, object] | None,
    *,
    baseline_hash: str | None,
    alt_kind: str | None,
) -> tuple[dict[str, object], dict[str, object]]:
    """Resolve Notes and Alt Text against the last system-synchronized hash."""
    notes_hash = script_sha256(notes_block.get("raw_transcript") or "")
    alt_hash = (
        script_sha256(alt_block.get("raw_transcript") or "")
        if alt_block is not None
        else None
    )
    notes_changed: bool | None = (
        notes_hash != baseline_hash if baseline_hash is not None else None
    )
    alt_changed: bool | None = (
        alt_hash != baseline_hash
        if baseline_hash is not None and alt_hash is not None
        else None
    )
    selected = notes_block
    selected_source = "author_notes"
    conflict = False

    if baseline_hash is None:
        if alt_block is None:
            resolution = "legacy_notes_only"
        elif alt_hash == notes_hash:
            resolution = "legacy_equal_versions"
        elif alt_kind == "plain":
            selected = _script_fields_from(notes_block, alt_block)
            selected_source = "alt_text"
            resolution = "legacy_plain_alt_text_user_edit"
        else:
            resolution = "legacy_conflict_notes_wins"
            conflict = True
    elif alt_block is None:
        resolution = "notes_user_edit" if notes_changed else "baseline_notes_only"
    elif not notes_changed and not alt_changed:
        resolution = "baseline_unchanged"
    elif notes_changed and not alt_changed:
        resolution = "notes_user_edit"
    elif not notes_changed and alt_changed:
        selected = _script_fields_from(notes_block, alt_block)
        selected_source = "alt_text"
        resolution = "alt_text_user_edit"
    elif notes_hash == alt_hash:
        resolution = "both_same_user_edit"
    else:
        resolution = "conflict_notes_wins"
        conflict = True

    selected_hash = script_sha256(selected.get("raw_transcript") or "")
    return selected, {
        "baseline_hash": baseline_hash,
        "notes_hash": notes_hash,
        "alt_text_hash": alt_hash,
        "notes_changed": notes_changed,
        "alt_text_changed": alt_changed,
        "selected_hash": selected_hash,
        "selected_source": selected_source,
        "resolution": resolution,
        "conflict": conflict,
        "legacy_alt_text_kind": alt_kind if baseline_hash is None else None,
    }


def _single_source_script_resolution(
    block: dict[str, object],
    *,
    source_name: str,
    baseline_hash: str | None,
    notes_present: bool,
) -> dict[str, object]:
    selected_hash = script_sha256(block.get("raw_transcript") or "")
    if source_name == "author_notes":
        notes_hash = selected_hash
        alt_hash = None
    elif source_name == "alt_text":
        notes_hash = None
        alt_hash = selected_hash
    else:
        notes_hash = None
        alt_hash = None
    return {
        "baseline_hash": baseline_hash,
        "notes_hash": notes_hash,
        "alt_text_hash": alt_hash,
        "notes_changed": (
            notes_hash != baseline_hash
            if baseline_hash is not None and notes_hash is not None
            else None
        ),
        "alt_text_changed": (
            alt_hash != baseline_hash
            if baseline_hash is not None and alt_hash is not None
            else None
        ),
        "selected_hash": selected_hash,
        "selected_source": source_name,
        "resolution": (
            f"{source_name}_only"
            if baseline_hash is not None
            else (
                (
                    "legacy_alt_text_insert"
                    if notes_present
                    else "legacy_alt_text_backfill"
                )
                if source_name == "alt_text"
                else "legacy_animation_pane_backfill"
            )
        ),
        "conflict": False,
        "legacy_alt_text_kind": None,
    }


def normalize_author_notes_authority(
    source_pptx: Path,
    output_pptx: Path,
    *,
    narration_order_policy: str = "geometry",
) -> dict[str, object]:
    """Normalize handles, compact Alt Text, and canonical Notes in one PPTX copy.

    Existing canonical Notes blocks keep their relative narration order. By
    default, current canvas geometry inserts only blocks absent from Notes.
    ``author_notes`` remains the explicit Notes authority policy.
    """
    source_pptx = source_pptx.resolve()
    output_pptx = output_pptx.resolve()
    narration_order_policy = normalize_narration_order_policy(
        narration_order_policy
    )
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    working_source, seeded_source = _ensure_notes_parts(source_pptx, output_pptx.parent)
    replacements: dict[str, bytes] = {}
    slides_out: list[dict[str, object]] = []
    try:
        with ZipFile(working_source) as source:
            slide_refs = presentation_slides(source)
            slide_width, slide_height = presentation_size(source)
            for ref in slide_refs:
                index = int(ref["index"])
                slide_part = str(ref["part"])
                slide_root = etree.fromstring(source.read(slide_part))
                effects = _native_effects(slide_root)
                effects_by_shape: dict[str, list[dict[str, object]]] = OrderedDict()
                for effect in effects:
                    effects_by_shape.setdefault(str(effect["shape_id"]), []).append(effect)
                ordered_shape_ids = list(effects_by_shape)
                by_shape, by_handle = _shape_maps(
                    slide_root,
                    slide_width=slide_width,
                    slide_height=slide_height,
                )
                blocks = _canonical_notes_blocks(
                    _notes_text_optional(source, slide_part)
                )
                input_had_notes = bool(blocks)
                alt_by_shape: dict[str, dict[str, object]] = {}
                for shape_id, shape in by_shape.items():
                    if not shape.get("top_level"):
                        continue
                    alt = parse_alt_protocol(str(shape.get("description") or ""))
                    if alt is not None:
                        alt_by_shape[shape_id] = alt

                records: list[dict[str, object]] = []
                removed_stale_notes: list[dict[str, object]] = []
                represented: set[str] = set()
                used_handles: set[str] = set()

                def add_record(
                    block: dict[str, object],
                    shape: dict[str, object],
                    *,
                    source_name: str,
                    shape_resolution: str,
                    order_source: str,
                    script_resolution: dict[str, object],
                    notes_anchor: bool = False,
                ) -> None:
                    _hydrate_block_semantics(block, shape)
                    shape_id = str(shape["shape_id"])
                    handle = str(block["handle"])
                    if not shape.get("top_level"):
                        raise ProtocolError(
                            f"slide {index} handle {handle!r} targets a nested shape"
                        )
                    if shape_id in represented or handle in used_handles:
                        raise ProtocolError(
                            f"slide {index} has a duplicate shape or handle for {handle!r}"
                        )
                    represented.add(shape_id)
                    used_handles.add(handle)
                    records.append(
                        {
                            "block": block,
                            "shape": shape,
                            "source": source_name,
                            "shape_resolution": shape_resolution,
                            "order_source": order_source,
                            "notes_anchor": notes_anchor,
                            "script_resolution": script_resolution,
                        }
                    )

                shapes_by_name: dict[str, list[dict[str, object]]] = {}
                for candidate in by_shape.values():
                    if candidate.get("top_level"):
                        shapes_by_name.setdefault(
                            _oneline(candidate.get("shape_name")), []
                        ).append(candidate)
                pre_resolved: dict[int, tuple[dict[str, object], str]] = {}
                claimed_shape_ids: set[str] = set()
                unresolved_block_indexes: list[int] = []
                for block_index, block in enumerate(blocks):
                    handle = str(block["handle"])
                    shape = by_handle.get(handle)
                    resolution = "alt_text_handle"
                    if shape is None:
                        name_matches = shapes_by_name.get(_oneline(handle)) or []
                        if len(name_matches) == 1:
                            shape = name_matches[0]
                            resolution = "shape_name"
                    if shape is None:
                        unresolved_block_indexes.append(block_index)
                        continue
                    shape_id = str(shape["shape_id"])
                    if shape_id in claimed_shape_ids:
                        raise ProtocolError(
                            f"slide {index} Notes handles resolve more than once to shape {shape_id!r}"
                        )
                    claimed_shape_ids.add(shape_id)
                    pre_resolved[block_index] = (shape, resolution)

                removed_stale_notes.extend(
                    {
                        "block_index": block_index,
                        "handle": str(blocks[block_index]["handle"]),
                        "resolution": "removed_missing_shape",
                    }
                    for block_index in unresolved_block_indexes
                )
                notes_reordered = _notes_sequence_was_reordered(
                    blocks,
                    pre_resolved,
                )

                for block_index, block in enumerate(blocks):
                    if block_index not in pre_resolved:
                        continue
                    handle = str(block["handle"])
                    shape, resolution = pre_resolved[block_index]
                    description = str(shape.get("description") or "")
                    managed_alt = alt_by_shape.get(str(shape["shape_id"]))
                    plain_alt_script = (
                        _legacy_plain_alt_script(description)
                        if managed_alt is None
                        else ""
                    )
                    if managed_alt is not None:
                        alt_candidate = managed_alt
                        alt_kind = "managed"
                    elif plain_alt_script:
                        transcript, markers = parse_marked_transcript(plain_alt_script)
                        alt_candidate = {
                            "handle": handle,
                            "semantic": "",
                            "semantic_concise": "",
                            "semantic_detailed": "",
                            "raw_transcript": plain_alt_script,
                            "transcript": transcript,
                            "markers": markers,
                            "animation_names": [],
                        }
                        alt_kind = "plain"
                    else:
                        alt_candidate = None
                        alt_kind = None
                    resolved_block, script_resolution = _resolve_notes_alt_script(
                        block,
                        alt_candidate,
                        baseline_hash=(
                            str(shape["script_baseline_sha256"])
                            if shape.get("script_baseline_sha256")
                            else None
                        ),
                        alt_kind=alt_kind,
                    )
                    add_record(
                        resolved_block,
                        shape,
                        source_name=str(script_resolution["selected_source"]),
                        shape_resolution=resolution,
                        order_source=(
                            "author_notes"
                            if narration_order_policy == "author_notes"
                            else _notes_block_order_source(
                                block,
                                shape,
                                handle_resolution=resolution,
                                user_reordered=notes_reordered,
                            )
                        ),
                        script_resolution=script_resolution,
                        notes_anchor=True,
                    )

                for shape_id in ordered_shape_ids:
                    if shape_id in represented:
                        continue
                    shape = by_shape.get(shape_id)
                    if shape is None:
                        raise ProtocolError(
                            f"slide {index} native animation targets missing shape {shape_id!r}"
                        )
                    alt = alt_by_shape.get(shape_id)
                    if alt is None:
                        handle = _fallback_handle(shape, used_handles)
                        plain_script = _legacy_plain_alt_script(
                            str(shape.get("description") or "")
                        )
                        transcript, markers = parse_marked_transcript(plain_script)
                        alt = {
                            "handle": handle,
                            "semantic": "",
                            "semantic_concise": "",
                            "semantic_detailed": "",
                            "raw_transcript": plain_script,
                            "transcript": transcript,
                            "markers": markers,
                            "animation_names": [],
                        }
                        source_name = "alt_text" if plain_script else "animation_pane"
                        resolution = "generated_handle"
                    else:
                        source_name = "alt_text"
                        resolution = "alt_text_handle"
                    add_record(
                        alt,
                        shape,
                        source_name=source_name,
                        shape_resolution=resolution,
                        order_source="geometry",
                        script_resolution=_single_source_script_resolution(
                            alt,
                            source_name=source_name,
                            baseline_hash=(
                                str(shape["script_baseline_sha256"])
                                if shape.get("script_baseline_sha256")
                                else None
                            ),
                            notes_present=input_had_notes,
                        ),
                    )

                remaining_alt = _spatially_ordered(
                    [
                        by_shape[shape_id]
                        for shape_id in alt_by_shape
                        if shape_id not in represented
                    ]
                )
                for shape in remaining_alt:
                    shape_id = str(shape["shape_id"])
                    add_record(
                        alt_by_shape[shape_id],
                        shape,
                        source_name="alt_text",
                        shape_resolution="alt_text_handle",
                        order_source="geometry",
                        script_resolution=_single_source_script_resolution(
                            alt_by_shape[shape_id],
                            source_name="alt_text",
                            baseline_hash=(
                                str(shape["script_baseline_sha256"])
                                if shape.get("script_baseline_sha256")
                                else None
                            ),
                            notes_present=input_had_notes,
                        ),
                    )

                records = _order_records_by_narration_policy(
                    records,
                    narration_order_policy=narration_order_policy,
                )
                inserted_blocks: list[dict[str, object]] = []
                for record_index, record in enumerate(records):
                    block = record["block"]
                    shape = record["shape"]
                    assert isinstance(block, dict) and isinstance(shape, dict)
                    shape_id = str(shape["shape_id"])
                    if shape_id in claimed_shape_ids:
                        continue
                    native_rows = effects_by_shape.get(shape_id) or []
                    inserted_blocks.append(
                        {
                            "shape_id": shape_id,
                            "handle": str(block["handle"]),
                            "shape_name": str(shape.get("shape_name") or ""),
                            "visible_texts": list(shape.get("visible_texts") or []),
                            "transcript": _oneline(block.get("transcript")),
                            "semantic_concise": str(
                                block.get("semantic_concise") or ""
                            ),
                            "semantic_detailed": str(
                                block.get("semantic_detailed") or ""
                            ),
                            "bbox": shape.get("bbox"),
                            "order_source": str(record.get("order_source") or "geometry"),
                            "order_index": record_index,
                            "reading_model": record.get("reading_model"),
                            "native_entrance_rows": [
                                {
                                    "native_order": int(effect["native_order"]),
                                    "name": str(effect["name"]),
                                    "trigger": str(effect.get("trigger") or ""),
                                    "pane_start_seconds": float(
                                        effect.get("pane_start_seconds") or 0.0
                                    ),
                                    "simultaneous_group": int(
                                        effect.get("simultaneous_group")
                                        or effect["native_order"]
                                    ),
                                }
                                for effect in native_rows
                                if str(effect.get("kind") or "entrance") == "entrance"
                            ],
                        }
                    )
                changes: list[dict[str, object]] = []
                for record_index, record in enumerate(records):
                    block = record["block"]
                    shape = record["shape"]
                    assert isinstance(block, dict) and isinstance(shape, dict)
                    shape_id = str(shape["shape_id"])
                    handle = str(block["handle"])
                    source_name = str(record["source"])
                    matches = slide_root.xpath(
                        f'.//p:spTree//p:cNvPr[@id="{shape_id}"]', namespaces=NS
                    )
                    if len(matches) != 1:
                        raise ProtocolError(
                            f"slide {index} shape {shape_id!r} resolved to {len(matches)} nodes"
                        )
                    previous_description = str(matches[0].get("descr") or "")
                    previous_baseline_hash = _script_baseline_from_cnvpr(matches[0])
                    previous_semantics = _semantic_provenance_from_cnvpr(matches[0]) or {}
                    script_resolution = record["script_resolution"]
                    assert isinstance(script_resolution, dict)
                    selected_baseline_hash = str(script_resolution["selected_hash"])
                    _, effective_effects, _ = _merge_effect_intents(
                        shape_id=shape_id,
                        handle=handle,
                        native_effects=effects_by_shape.get(shape_id) or [],
                        markers=list(block.get("markers") or []),
                        animation_names=[
                            str(name) for name in block.get("animation_names") or []
                        ],
                        authority=source_name,
                        synthetic_order_base=(
                            1_000_000 + index * 10_000 + record_index * 100
                        ),
                    )
                    generated = build_system_alt_text(
                        handle=handle,
                        animation_names=[
                            str(effect["name"]) for effect in effective_effects
                        ],
                        raw_script=str(block.get("raw_transcript") or ""),
                        shape_name=(
                            _oneline(block.get("semantic_concise"))
                            or _oneline(block.get("semantic"))
                            or str(shape.get("shape_name") or handle)
                        ),
                        shape_id=shape_id,
                        slide_index=index,
                        existing_description=previous_description,
                        order_source=str(record["order_source"]),
                        baseline_hash=selected_baseline_hash,
                    )
                    matches[0].set("descr", generated)
                    written_baseline_hash = _set_script_baseline_on_cnvpr(
                        matches[0],
                        block.get("raw_transcript") or "",
                        order_source=str(record["order_source"]),
                        order_index=record_index,
                    )
                    _set_semantic_provenance_on_cnvpr(
                        matches[0],
                        semantic_concise=block.get("semantic_concise") or "",
                        semantic_detailed=block.get("semantic_detailed") or "",
                    )
                    if written_baseline_hash != selected_baseline_hash:
                        raise ProtocolError(
                            f"slide {index} shape {shape_id!r} script baseline write mismatch"
                        )
                    if (
                        previous_description != generated
                        or previous_baseline_hash != written_baseline_hash
                        or previous_semantics.get("semantic_concise")
                        != block.get("semantic_concise")
                        or previous_semantics.get("semantic_detailed")
                        != block.get("semantic_detailed")
                    ):
                        changes.append(
                            {
                                "shape_id": shape_id,
                                "shape_name": shape["shape_name"],
                                "previous_handle": parse_alt_id(previous_description),
                                "alt_text_handle": handle,
                                "metadata_refreshed": True,
                                "shape_resolution": record["shape_resolution"],
                                "order_source": record["order_source"],
                                "resolution": source_name,
                                "semantic_concise": block.get("semantic_concise"),
                                "semantic_detailed": block.get("semantic_detailed"),
                                **script_resolution,
                            }
                        )

                if records:
                    replacements[slide_part] = etree.tostring(
                        slide_root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )

                notes_backfilled = not input_had_notes and bool(records)
                notes_records = records
                notes_lines: list[str] = []
                for record_index, record in enumerate(notes_records):
                    block = record["block"]
                    shape = record["shape"]
                    assert isinstance(block, dict) and isinstance(shape, dict)
                    if record_index:
                        notes_lines.append("")
                    semantic = (
                        _oneline(block.get("semantic_concise"))
                        or _oneline(block.get("semantic"))
                        or _oneline(shape.get("shape_name"))
                    )
                    header = f"## [{block['handle']}]"
                    if semantic:
                        header += f" {semantic}"
                    notes_lines.append(header)
                    raw_script = str(block.get("raw_transcript") or "").strip()
                    if raw_script:
                        notes_lines.extend(raw_script.splitlines())
                canonical_notes = "\n".join(notes_lines)
                original_handles = [str(block["handle"]) for block in blocks]
                canonical_handles = [
                    str(record["block"]["handle"]) for record in notes_records
                ]
                notes_order_changed = canonical_handles != original_handles
                notes_sync_requested = (
                    notes_backfilled
                    or notes_order_changed
                    or bool(removed_stale_notes)
                    or bool(changes)
                    or any(
                    str(record["script_resolution"]["resolution"])
                    in {
                        "alt_text_user_edit",
                        "legacy_plain_alt_text_user_edit",
                    }
                    for record in notes_records
                    )
                )
                notes_rewritten = notes_sync_requested and (
                    _notes_text_optional(source, slide_part).strip() != canonical_notes.strip()
                )
                if notes_rewritten:
                    notes_part = _notes_part_optional(source, slide_part)
                    if notes_part is None:
                        raise ProtocolError(
                            f"slide {index} could not create an Author Notes part"
                        )
                    notes_root = etree.fromstring(source.read(notes_part))
                    _replace_notes_body(notes_root, notes_lines)
                    replacements[notes_part] = etree.tostring(
                        notes_root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )

                slides_out.append(
                    {
                        "index": index,
                        "target_count": len(ordered_shape_ids),
                        "notes_block_count": len(blocks),
                        "notes_backfilled": notes_backfilled,
                        "notes_rewritten": notes_rewritten,
                        "generated_notes_block_count": max(
                            0, len(notes_records) - len(blocks)
                        ),
                        "inserted_blocks": inserted_blocks,
                        "notes_order_changed": notes_order_changed,
                        "removed_stale_notes": removed_stale_notes,
                        "script_source": (
                            "author_notes"
                            if input_had_notes
                            else ("alt_text" if alt_by_shape else "animation_pane")
                        ),
                        "narration_order_policy": narration_order_policy,
                        "alt_text_changes": changes,
                        "script_resolutions": [
                            {
                                "shape_id": str(record["shape"]["shape_id"]),
                                "handle": str(record["block"]["handle"]),
                                "semantic_concise": str(
                                    record["block"].get("semantic_concise") or ""
                                ),
                                "semantic_detailed": str(
                                    record["block"].get("semantic_detailed") or ""
                                ),
                                **record["script_resolution"],
                            }
                            for record in records
                        ],
                    }
                )

            with tempfile.NamedTemporaryFile(
                prefix=output_pptx.stem + ".",
                suffix=".pptx",
                dir=output_pptx.parent,
                delete=False,
            ) as temporary:
                temporary_path = Path(temporary.name)
            try:
                with ZipFile(temporary_path, "w") as destination:
                    for info in source.infolist():
                        data = replacements.get(info.filename, source.read(info.filename))
                        destination.writestr(_copy_zipinfo(info), data)
                shutil.move(temporary_path, output_pptx)
            finally:
                temporary_path.unlink(missing_ok=True)
    finally:
        if seeded_source is not None:
            seeded_source.unlink(missing_ok=True)

    return {
        "schema_version": "paper2video_author_notes_authority.v5",
        "created_at": _utc_now(),
        "source_pptx": str(source_pptx),
        "source_sha256": file_sha256(source_pptx),
        "output_pptx": str(output_pptx),
        "output_sha256": file_sha256(output_pptx),
        "slide_count": len(slides_out),
        "narration_order_policy": narration_order_policy,
        "alt_text_change_count": sum(
            len(slide["alt_text_changes"]) for slide in slides_out
        ),
        "notes_backfill_count": sum(
            1 for slide in slides_out if slide["notes_backfilled"]
        ),
        "notes_rewrite_count": sum(
            1 for slide in slides_out if slide["notes_rewritten"]
        ),
        "inserted_block_count": sum(
            len(slide["inserted_blocks"]) for slide in slides_out
        ),
        "script_conflict_count": sum(
            1
            for slide in slides_out
            for resolution in slide["script_resolutions"]
            if resolution["conflict"]
        ),
        "slides": slides_out,
    }


def _shape_change_inventory(pptx_path: Path) -> dict[tuple[str, str], dict[str, object]]:
    inventory: dict[tuple[str, str], dict[str, object]] = {}
    with ZipFile(pptx_path) as archive:
        slide_width, slide_height = presentation_size(archive)
        for ref in presentation_slides(archive):
            slide_root = etree.fromstring(archive.read(str(ref["part"])))
            effects = _native_effects(slide_root)
            effects_by_shape: dict[str, list[dict[str, object]]] = OrderedDict()
            for effect in effects:
                effects_by_shape.setdefault(str(effect["shape_id"]), []).append(effect)
            by_shape, _ = _shape_maps(
                slide_root,
                slide_width=slide_width,
                slide_height=slide_height,
            )
            for shape_id, shape in by_shape.items():
                if not shape.get("top_level"):
                    continue
                nodes = slide_root.xpath(
                    f'.//p:spTree/*[p:nvSpPr/p:cNvPr[@id="{shape_id}"] '
                    f'or p:nvPicPr/p:cNvPr[@id="{shape_id}"] '
                    f'or p:nvGraphicFramePr/p:cNvPr[@id="{shape_id}"] '
                    f'or p:nvGrpSpPr/p:cNvPr[@id="{shape_id}"]]',
                    namespaces=NS,
                )
                if len(nodes) != 1:
                    continue
                canonical = deepcopy(nodes[0])
                for non_visual in canonical.xpath(".//p:cNvPr", namespaces=NS):
                    non_visual.attrib.pop("descr", None)
                shape_effects = [
                    {
                        "name": effect.get("name"),
                        "kind": effect.get("kind"),
                        "trigger": effect.get("trigger"),
                        "delay_seconds": effect.get("delay_seconds"),
                        "duration_seconds": effect.get("duration_seconds"),
                        "native_order": effect.get("native_order"),
                    }
                    for effect in effects_by_shape.get(shape_id) or []
                ]
                digest = hashlib.sha256()
                digest.update(etree.tostring(canonical, method="c14n"))
                digest.update(
                    json.dumps(shape_effects, sort_keys=True).encode("utf-8")
                )
                handle = parse_alt_id(str(shape.get("description") or ""))
                inventory[(str(ref["stable_id"]), shape_id)] = {
                    "slide_index": int(ref["index"]),
                    "stable_slide_id": str(ref["stable_id"]),
                    "shape_id": shape_id,
                    "handle": handle,
                    "shape_name": str(shape.get("shape_name") or ""),
                    "text": _oneline(" ".join(nodes[0].xpath(".//a:t/text()", namespaces=NS))),
                    "bbox": shape.get("bbox"),
                    "animations": [str(effect.get("name") or "") for effect in shape_effects],
                    "fingerprint": digest.hexdigest(),
                }
    return inventory


def detect_pptx_changes(
    baseline_pptx: Path,
    edited_pptx: Path,
) -> dict[str, object]:
    """Find user-visible shape and native-animation changes between two decks."""
    baseline_pptx = baseline_pptx.resolve()
    edited_pptx = edited_pptx.resolve()
    before = _shape_change_inventory(baseline_pptx)
    after = _shape_change_inventory(edited_pptx)
    changes: list[dict[str, object]] = []
    for key in sorted(set(before) | set(after)):
        old = before.get(key)
        new = after.get(key)
        if old is None:
            kind = "added"
        elif new is None:
            kind = "removed"
        elif old["fingerprint"] != new["fingerprint"]:
            kind = "modified"
        else:
            continue
        current = new or old
        assert current is not None
        changes.append(
            {
                "kind": kind,
                "slide_index": current["slide_index"],
                "stable_slide_id": current["stable_slide_id"],
                "shape_id": current["shape_id"],
                "handle": (new or {}).get("handle") or (old or {}).get("handle"),
                "before": old,
                "after": new,
            }
        )
    return {
        "schema_version": "paper2video_pptx_changes.v1",
        "created_at": _utc_now(),
        "baseline_pptx": str(baseline_pptx),
        "baseline_sha256": file_sha256(baseline_pptx),
        "edited_pptx": str(edited_pptx),
        "edited_sha256": file_sha256(edited_pptx),
        "change_count": len(changes),
        "changes": changes,
    }


def write_protocol_to_pptx(
    source_pptx: Path,
    protocol: dict[str, object],
    output_pptx: Path,
) -> dict[str, object]:
    """Persist narration into Notes, compact Alt Text, and hidden OOXML provenance."""
    source_pptx = source_pptx.resolve()
    output_pptx = output_pptx.resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    replacements: dict[str, bytes] = {}
    slides_out: list[dict[str, object]] = []
    with ZipFile(source_pptx) as source:
        refs = presentation_slides(source)
        slides = list(protocol.get("slides") or [])
        if len(refs) != len(slides):
            raise ProtocolError(
                f"protocol slide count {len(slides)} != PPTX slide count {len(refs)}"
            )
        for ref, slide in zip(refs, slides):
            index = int(ref["index"])
            if int(slide.get("index") or 0) != index:
                raise ProtocolError(f"protocol slide order mismatch at slide {index}")
            slide_part = str(ref["part"])
            slide_root = etree.fromstring(source.read(slide_part))
            notes_part = _notes_part_optional(source, slide_part)
            if notes_part is None:
                raise ProtocolError(f"slide {index} has no Author Notes part")
            notes_root = etree.fromstring(source.read(notes_part))
            notes_lines: list[str] = []
            alt_change_count = 0
            for block_index, block in enumerate(slide.get("blocks") or []):
                _hydrate_block_semantics(block, block)
                if block_index:
                    notes_lines.append("")
                semantic = (
                    _oneline(block.get("semantic_concise"))
                    or _oneline(block.get("semantic"))
                    or _oneline(block.get("shape_name"))
                )
                header = f"## [{block['handle']}]"
                if semantic:
                    header += f" {semantic}"
                notes_lines.append(header)
                raw_script = canonicalize_user_script(
                    str(block.get("raw_transcript") or "").strip()
                )
                block["raw_transcript"] = raw_script
                if raw_script:
                    notes_lines.extend(raw_script.splitlines())

                shape_id = str(block["shape_id"])
                nodes = slide_root.xpath(
                    f'.//p:spTree//p:cNvPr[@id="{shape_id}"]', namespaces=NS
                )
                if len(nodes) != 1:
                    raise ProtocolError(
                        f"slide {index} protocol shape {shape_id!r} resolved to {len(nodes)} nodes"
                    )
                previous = str(nodes[0].get("descr") or "")
                resolution = str(block.get("handle_resolution") or "")
                order_source = str(block.get("order_source") or "")
                if order_source not in {"author_notes", "animation_pane", "geometry"}:
                    if str(block.get("script_source") or "") in {
                        "author_notes",
                        "user_script",
                        "llm_regeneration",
                    }:
                        order_source = "author_notes"
                    elif resolution == "geometry_order":
                        order_source = "geometry"
                    else:
                        order_source = "animation_pane"
                generated = build_system_alt_text(
                    handle=str(block["handle"]),
                    animation_names=[
                        str(effect["name"]) for effect in block.get("effects") or []
                    ],
                    raw_script=raw_script,
                    shape_name=(
                        _oneline(block.get("semantic_concise"))
                        or _oneline(block.get("semantic"))
                        or str(block.get("shape_name") or block["handle"])
                    ),
                    shape_id=shape_id,
                    slide_index=index,
                    existing_description=previous,
                    order_source=order_source,
                )
                nodes[0].set("descr", generated)
                _set_script_baseline_on_cnvpr(
                    nodes[0],
                    raw_script,
                    order_source=order_source,
                    order_index=block_index,
                )
                _set_semantic_provenance_on_cnvpr(
                    nodes[0],
                    semantic_concise=block.get("semantic_concise") or "",
                    semantic_detailed=block.get("semantic_detailed") or "",
                )
                if previous != generated:
                    alt_change_count += 1
            _replace_notes_body(notes_root, notes_lines)
            replacements[slide_part] = etree.tostring(
                slide_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            replacements[notes_part] = etree.tostring(
                notes_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            slides_out.append(
                {
                    "index": index,
                    "block_count": len(slide.get("blocks") or []),
                    "alt_text_change_count": alt_change_count,
                    "blocks": [
                        {
                            "shape_id": str(block.get("shape_id") or ""),
                            "handle": str(block.get("handle") or ""),
                            "semantic_concise": str(
                                block.get("semantic_concise") or ""
                            ),
                            "semantic_detailed": str(
                                block.get("semantic_detailed") or ""
                            ),
                        }
                        for block in slide.get("blocks") or []
                    ],
                }
            )
        with tempfile.NamedTemporaryFile(
            prefix=output_pptx.stem + ".protocol.",
            suffix=".pptx",
            dir=output_pptx.parent,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with ZipFile(temporary_path, "w") as destination:
                for info in source.infolist():
                    destination.writestr(
                        _copy_zipinfo(info),
                        replacements.get(info.filename, source.read(info.filename)),
                    )
            shutil.move(temporary_path, output_pptx)
        finally:
            temporary_path.unlink(missing_ok=True)
    return {
        "schema_version": "paper2video_protocol_writeback.v1",
        "created_at": _utc_now(),
        "source_pptx": str(source_pptx),
        "output_pptx": str(output_pptx),
        "output_sha256": file_sha256(output_pptx),
        "slides": slides_out,
    }


def _split_words(text: str, count: int) -> list[str]:
    words = _oneline(text).split()
    if count <= 0:
        raise ProtocolError("cannot split narration across zero animated elements")
    if len(words) < count:
        raise ProtocolError(
            f"cannot split {len(words)} narration words across {count} animated elements"
        )
    chunks: list[str] = []
    cursor = 0
    for index in range(count):
        remaining_words = len(words) - cursor
        remaining_parts = count - index
        width = (remaining_words + remaining_parts - 1) // remaining_parts
        chunks.append(" ".join(words[cursor:cursor + width]))
        cursor += width
    return chunks


def _replace_notes_body(notes_root: etree._Element, lines: list[str]) -> None:
    bodies = notes_root.xpath(
        './/p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]', namespaces=NS
    )
    if len(bodies) != 1:
        raise ProtocolError(
            f"notes slide must contain one body placeholder, found {len(bodies)}"
        )
    text_body = bodies[0].find(f"{{{P_NS}}}txBody")
    if text_body is None:
        raise ProtocolError("notes body placeholder has no p:txBody")
    for paragraph in list(text_body.findall(f"{{{A_NS}}}p")):
        text_body.remove(paragraph)
    for line in lines:
        paragraph = etree.SubElement(text_body, f"{{{A_NS}}}p")
        if line:
            run = etree.SubElement(paragraph, f"{{{A_NS}}}r")
            etree.SubElement(run, f"{{{A_NS}}}rPr", lang="en-US", dirty="0")
            text = etree.SubElement(run, f"{{{A_NS}}}t")
            text.text = line
        etree.SubElement(paragraph, f"{{{A_NS}}}endParaRPr", lang="en-US", dirty="0")


def bootstrap_protocol(
    source_pptx: Path,
    script_json: Path,
    output_pptx: Path,
) -> dict[str, object]:
    """Seed canonical Notes and Alt Text from native animations and narration.

    This is a one-time deterministic bridge for an ordinary animated deck. It
    preserves slide timing trees and visible content, assigns stable handles to
    animated top-level elements, and divides each slide narration across those
    elements in row-aware canvas order while preserving each target's native
    Animation Pane effects and trigger metadata.
    """
    try:
        script_payload = json.loads(script_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"could not read script {script_json}: {exc}") from exc
    sections = script_payload.get("sections") or []
    if not isinstance(sections, list) or not sections:
        raise ProtocolError("bootstrap script has no sections")

    source_pptx = source_pptx.resolve()
    with ZipFile(source_pptx) as source:
        slide_refs = presentation_slides(source)
        slide_width, slide_height = presentation_size(source)
        if len(slide_refs) != len(sections):
            raise ProtocolError(
                f"script section count {len(sections)} != PPTX slide count {len(slide_refs)}"
            )
        replacements: dict[str, bytes] = {}
        report_slides: list[dict[str, object]] = []
        for ref, section in zip(slide_refs, sections):
            index = int(ref["index"])
            slide_part = str(ref["part"])
            notes_part = _notes_part(source, slide_part)
            slide_root = etree.fromstring(source.read(slide_part))
            notes_root = etree.fromstring(source.read(notes_part))
            effects = _native_effects(slide_root)
            entrance_effects = [
                effect for effect in effects if effect.get("kind") == "entrance"
            ]
            if not entrance_effects:
                raise ProtocolError(f"slide {index} has no native entrance effects to bootstrap")
            unsupported = [
                effect["name"]
                for effect in entrance_effects
                if effect["name"] not in VIDEO_EFFECT_SECONDS
            ]
            if unsupported:
                raise ProtocolError(
                    f"slide {index} uses video-unsupported native effects {unsupported!r}"
                )
            by_shape, _ = _shape_maps(
                slide_root,
                slide_width=slide_width,
                slide_height=slide_height,
            )
            grouped: OrderedDict[str, list[dict[str, object]]] = OrderedDict()
            for effect in effects:
                grouped.setdefault(str(effect["shape_id"]), []).append(effect)
            spatial_targets = _spatially_ordered(
                [
                    {
                        **by_shape[shape_id],
                        "shape_effects": shape_effects,
                    }
                    for shape_id, shape_effects in grouped.items()
                ]
            )
            chunks = _split_words(
                str(section.get("text") or ""), len(spatial_targets)
            )
            used_handles: set[str] = set()
            lines: list[str] = []
            blocks: list[dict[str, object]] = []
            for block_index, (target, transcript) in enumerate(
                zip(spatial_targets, chunks)
            ):
                shape_id = str(target["shape_id"])
                shape_effects = target["shape_effects"]
                assert isinstance(shape_effects, list)
                shape = by_shape.get(shape_id)
                if shape is None:
                    raise ProtocolError(
                        f"slide {index} animation target {shape_id!r} has no shape"
                    )
                if not shape.get("top_level"):
                    raise ProtocolError(
                        f"slide {index} animation target {shape_id!r} is nested; group it first"
                    )
                existing = parse_alt_id(str(shape.get("description") or ""))
                base = existing or _oneline(shape.get("shape_name")) or f"Shape {shape_id}"
                handle = base
                if handle in used_handles:
                    handle = f"{base} #{shape_id}"
                if handle in used_handles:
                    raise ProtocolError(
                        f"slide {index} could not create a unique handle for shape {shape_id}"
                    )
                used_handles.add(handle)
                matches = slide_root.xpath(
                    f'.//p:spTree//p:cNvPr[@id="{shape_id}"]', namespaces=NS
                )
                if len(matches) != 1:
                    raise ProtocolError(
                        f"slide {index} shape {shape_id!r} resolved to {len(matches)} nodes"
                    )
                if block_index:
                    lines.append("")
                raw_script = transcript
                semantic_concise, semantic_detailed = _semantic_values(
                    transcript=raw_script,
                    visible_texts=shape.get("visible_texts") or [],
                    shape_name=shape.get("shape_name"),
                    handle=handle,
                )
                lines.append(f"## [{handle}] {semantic_concise}")
                lines.append(raw_script)
                matches[0].set(
                    "descr",
                    build_system_alt_text(
                        handle=handle,
                        animation_names=[str(effect["name"]) for effect in shape_effects],
                        raw_script=raw_script,
                        shape_name=semantic_concise,
                        shape_id=shape_id,
                        slide_index=index,
                        existing_description=str(shape.get("description") or ""),
                        order_source="geometry",
                    ),
                )
                _set_script_baseline_on_cnvpr(
                    matches[0],
                    raw_script,
                    order_source="geometry",
                    order_index=block_index,
                )
                _set_semantic_provenance_on_cnvpr(
                    matches[0],
                    semantic_concise=semantic_concise,
                    semantic_detailed=semantic_detailed,
                )
                blocks.append(
                    {
                        "handle": handle,
                        "shape_id": shape_id,
                        "semantic": semantic_concise,
                        "semantic_concise": semantic_concise,
                        "semantic_detailed": semantic_detailed,
                        "effects": [effect["name"] for effect in shape_effects],
                        "transcript": transcript,
                    }
                )
            _replace_notes_body(notes_root, lines)
            replacements[slide_part] = etree.tostring(
                slide_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            replacements[notes_part] = etree.tostring(
                notes_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            report_slides.append(
                {
                    "index": index,
                    "section_id": str(section.get("id") or f"slide-{ref['stable_id']}"),
                    "block_count": len(blocks),
                    "effect_count": sum(len(block["effects"]) for block in blocks),
                    "blocks": blocks,
                }
            )

        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=output_pptx.stem + ".",
            suffix=".pptx",
            dir=output_pptx.parent,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with ZipFile(temporary_path, "w") as destination:
                for info in source.infolist():
                    data = (
                        replacements[info.filename]
                        if info.filename in replacements
                        else source.read(info.filename)
                    )
                    destination.writestr(_copy_zipinfo(info), data)
            shutil.move(temporary_path, output_pptx)
        finally:
            temporary_path.unlink(missing_ok=True)

    validated = extract_protocol(
        output_pptx,
        section_ids=[str(section["section_id"]) for section in report_slides],
    )
    return {
        "schema_version": "paper2video_editable_pptx_bootstrap.v1",
        "created_at": _utc_now(),
        "source_pptx": str(source_pptx),
        "output_pptx": str(output_pptx.resolve()),
        "slide_count": len(report_slides),
        "effect_count": validated["effect_count"],
        "slides": report_slides,
        "validated_source_sha256": validated["source_sha256"],
    }


def _align_block(
    words: list[dict[str, object]], cursor: int, transcript: str
) -> tuple[int, int]:
    target = _normalized_chars(transcript)
    if not target:
        raise ProtocolError("animation transcript block is empty")
    combined = ""
    end = cursor
    while end < len(words) and len(combined) < len(target):
        combined += _normalized_chars(words[end].get("text"))
        end += 1
    if combined != target:
        raise ProtocolError(
            f"could not align transcript {transcript!r} at timing word {cursor}; "
            f"received {combined!r}, expected {target!r}"
        )
    return cursor, end


def _marker_timing(
    words: list[dict[str, object]],
    start: int,
    end: int,
    normalized_char: int,
) -> tuple[float, int]:
    if normalized_char == 0:
        return float(words[start].get("start") or 0.0), start
    consumed = 0
    for index in range(start, end):
        token_length = len(_normalized_chars(words[index].get("text")))
        if consumed == normalized_char:
            return float(words[index].get("start") or 0.0), index
        consumed += token_length
        if consumed > normalized_char:
            raise ProtocolError("animation marker is not positioned on a spoken word boundary")
    if consumed == normalized_char:
        return float(words[end - 1].get("end") or 0.0), end - 1
    raise ProtocolError("animation marker position exceeds its transcript block")


def _marker_end_timing(
    words: list[dict[str, object]],
    start: int,
    end: int,
    normalized_end_char: int,
) -> tuple[float, int]:
    """Resolve a spoken span's exclusive character end to a word end boundary."""
    if normalized_end_char <= 0:
        raise ProtocolError("Spotlight spoken scope end must follow its start")
    consumed = 0
    for index in range(start, end):
        consumed += len(_normalized_chars(words[index].get("text")))
        if consumed == normalized_end_char:
            return float(words[index].get("end") or 0.0), index
        if consumed > normalized_end_char:
            raise ProtocolError(
                "Spotlight spoken scope does not end on a spoken word boundary"
            )
    raise ProtocolError("Spotlight spoken scope end exceeds its transcript block")


def _finite_seconds(raw: object, *, label: str) -> float:
    try:
        value = float(raw)
    except (TypeError, ValueError) as exc:
        raise ProtocolError(f"{label} must be a finite number") from exc
    if not math.isfinite(value):
        raise ProtocolError(f"{label} must be a finite number")
    return value


def _effect_duration_seconds(effect: dict[str, object]) -> float:
    name = str(effect.get("name") or "")
    kind = str(effect.get("kind") or "entrance")
    default_duration = VIDEO_EFFECT_SECONDS.get(name, 0.48) if kind == "entrance" else 2.4
    return max(
        0.0 if kind == "entrance" else 0.2,
        float(effect.get("duration_seconds") or default_duration),
    )


def _effect_pane_end_seconds(effect: dict[str, object]) -> float:
    pane_start = float(effect.get("pane_start_seconds") or 0.0)
    pane_end = effect.get("pane_end_seconds")
    if pane_end is not None:
        return float(pane_end)
    return pane_start + _effect_duration_seconds(effect)


def _align_timing_parts(
    slide: dict[str, object],
    aligned_blocks: list[dict[str, object]],
    timing_section: dict[str, object] | None,
) -> dict[int, dict[str, object]]:
    """Validate explicit output-audio windows against aligned Notes blocks."""
    if timing_section is None:
        return {}
    raw_parts = timing_section.get("tts_parts")
    if raw_parts is None:
        # Transitional compatibility for timing payloads produced while the
        # serialized contract was being introduced. New output writes tts_parts.
        raw_parts = timing_section.get("parts")
    if raw_parts is None:
        return {}
    if not isinstance(raw_parts, list):
        raise ProtocolError(
            f"timing section {slide['section_id']!r} parts must be an array"
        )
    narrated = [
        aligned
        for aligned in aligned_blocks
        if str(aligned["block"].get("transcript") or "")
    ]
    if len(raw_parts) != len(narrated):
        raise ProtocolError(
            f"timing section {slide['section_id']!r} has {len(raw_parts)} part windows "
            f"for {len(narrated)} narrated blocks"
        )

    by_block: dict[int, dict[str, object]] = {}
    for raw_part, aligned in zip(raw_parts, narrated, strict=True):
        if not isinstance(raw_part, dict):
            raise ProtocolError(
                f"timing section {slide['section_id']!r} contains a non-object part"
            )
        block = aligned["block"]
        block_index = int(aligned["index"])
        part_id = str(raw_part.get("id") or "")
        handle = str(block.get("handle") or "")
        if part_id != handle:
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} does not "
                f"match narrated block {handle!r}"
            )
        values = {
            key: _finite_seconds(
                raw_part.get(key),
                label=(
                    f"timing section {slide['section_id']!r} part {part_id!r} {key}"
                ),
            )
            for key in (
                "pre_roll_seconds",
                "pre_roll_start",
                "pre_roll_end",
                "audio_start",
                "narration_start",
                "narration_end",
                "audio_end",
            )
        }
        if min(values.values()) < -0.0005:
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} has a "
                "negative output-audio timestamp"
            )
        if abs(values["pre_roll_start"] - values["audio_start"]) > 0.002:
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} pre-roll "
                "does not begin at audio_start"
            )
        if abs(
            values["pre_roll_end"]
            - values["pre_roll_start"]
            - values["pre_roll_seconds"]
        ) > 0.002:
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} has an "
                "inconsistent pre-roll window"
            )
        if not (
            values["audio_start"]
            <= values["pre_roll_end"] + 0.002
            <= values["narration_start"] + 0.002
            <= values["narration_end"] + 0.002
            <= values["audio_end"] + 0.002
        ):
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} has an "
                "invalid audio/pre-roll/narration window"
            )
        if (
            abs(values["narration_start"] - float(aligned["narration_start"])) > 0.002
            or abs(values["narration_end"] - float(aligned["spoken_end"])) > 0.002
        ):
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} does not "
                "match its Edge word boundaries"
            )
        try:
            part_word_start = int(raw_part.get("word_start"))
            part_word_end = int(raw_part.get("word_end"))
        except (TypeError, ValueError) as exc:
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} has "
                "invalid word indices"
            ) from exc
        if (
            part_word_start != int(aligned["word_start"])
            or part_word_end != int(aligned["word_end"]) - 1
        ):
            raise ProtocolError(
                f"timing section {slide['section_id']!r} part {part_id!r} word "
                "indices do not match its transcript"
            )
        normalized = {**raw_part, **values}
        for key in ("pane_anchor_start_seconds", "pane_anchor_end_seconds"):
            if raw_part.get(key) is not None:
                normalized[key] = _finite_seconds(
                    raw_part.get(key),
                    label=(
                        f"timing section {slide['section_id']!r} part "
                        f"{part_id!r} {key}"
                    ),
                )
        by_block[block_index] = normalized
    return by_block


def _serialized_native_schedule(
    slide: dict[str, object],
    aligned_blocks: list[dict[str, object]],
    native_pairs: list[
        tuple[dict[str, object], dict[str, object], dict[str, object]]
    ],
    timing_parts: dict[int, dict[str, object]],
    timing_section: dict[str, object] | None,
) -> dict[int, dict[str, object]]:
    """Place Pane-relative native phases between their narration blocks.

    A phase ends at one narrated entrance simultaneous group. The complete
    Pane segment since the previous narrated group, including silent connector
    rows and Pane delays, is placed in that group's output-only audio pre-roll.
    Effects inside a With Previous group retain their original relative starts.
    """
    if not native_pairs:
        return {}

    narrated_groups: OrderedDict[int, dict[str, object]] = OrderedDict()
    all_group_positions: dict[int, list[int]] = {}
    for position, (aligned, _marker, effect) in enumerate(native_pairs):
        group = int(effect.get("simultaneous_group") or effect["native_order"])
        all_group_positions.setdefault(group, []).append(position)
        if (
            str(effect.get("kind") or "entrance") == "entrance"
            and str(aligned["block"].get("transcript") or "")
        ):
            record = narrated_groups.setdefault(
                group,
                {"group": group, "positions": [], "blocks": []},
            )
            record["positions"].append(position)
            if aligned not in record["blocks"]:
                record["blocks"].append(aligned)

    # A native-only slide keeps its original Pane clock.
    if not narrated_groups:
        return {
            int(effect["native_order"]): {
                "start": float(effect.get("pane_start_seconds") or 0.0),
                "end": _effect_pane_end_seconds(effect),
                "sequence_gate": 0.0,
                "phase_index": 0,
                "phase_audio_start": 0.0,
                "timing_resolution": "native_pane_authority",
            }
            for _aligned, _marker, effect in native_pairs
        }

    groups = list(narrated_groups.values())
    previous_max_block = 0
    for group in groups:
        block_indices = sorted(int(aligned["index"]) for aligned in group["blocks"])
        if block_indices[0] <= previous_max_block:
            raise ProtocolError(
                f"slide {slide['index']} narrated Animation Pane groups do not follow "
                "the selected narration block order required for serialized playback"
            )
        previous_max_block = block_indices[-1]
        group["block_indices"] = block_indices
        group["first_block"] = min(group["blocks"], key=lambda item: int(item["index"]))
        group["last_native_position"] = max(all_group_positions[int(group["group"])])
        group_effects = [
            native_pairs[position][2]
            for position in all_group_positions[int(group["group"])]
        ]
        group["pane_end"] = max(
            _effect_pane_end_seconds(effect) for effect in group_effects
        )

    schedule: dict[int, dict[str, object]] = {}
    previous_native_position = -1
    previous_pane_end = 0.0
    previous_narration_end = 0.0
    for phase_index, group in enumerate(groups, start=1):
        first_block = group["first_block"]
        first_block_index = int(first_block["index"])
        previous_narration_end = max(
            [
                float(aligned["spoken_end"])
                for aligned in aligned_blocks
                if int(aligned["index"]) < first_block_index
                and str(aligned["block"].get("transcript") or "")
            ]
            + [0.0]
        )
        part = timing_parts.get(first_block_index)
        phase_pane_start = previous_pane_end
        phase_pane_end = float(group["pane_end"])
        if phase_pane_end + 0.002 < phase_pane_start:
            raise ProtocolError(
                f"slide {slide['index']} narrated Pane phase {phase_index} ends before "
                "the previous narrated phase"
            )
        if part is not None:
            phase_audio_start = float(part["pre_roll_start"])
            explicit_start = part.get("pane_anchor_start_seconds")
            explicit_end = part.get("pane_anchor_end_seconds")
            if explicit_start is not None and abs(float(explicit_start) - phase_pane_start) > 0.002:
                raise ProtocolError(
                    f"slide {slide['index']} part {part['id']!r} Pane anchor start "
                    "does not match the serialized animation phase"
                )
            if explicit_end is not None and abs(float(explicit_end) - phase_pane_end) > 0.002:
                raise ProtocolError(
                    f"slide {slide['index']} part {part['id']!r} Pane anchor end does "
                    "not match the serialized animation phase"
                )
            pane_span = phase_pane_end - phase_pane_start
            if abs(float(part["pre_roll_seconds"]) - pane_span) > 0.003:
                raise ProtocolError(
                    f"slide {slide['index']} part {part['id']!r} pre-roll "
                    f"{float(part['pre_roll_seconds']):.3f}s does not match its "
                    f"{pane_span:.3f}s Animation Pane phase"
                )
            phase_audio_end = float(part["pre_roll_end"])
            timing_resolution = "native_pane_phase_in_block_preroll"
        else:
            phase_audio_start = previous_narration_end
            phase_audio_end = phase_audio_start + phase_pane_end - phase_pane_start
            timing_resolution = "native_pane_phase_in_word_timing_gap"

        if phase_audio_start + 0.002 < previous_narration_end:
            raise ProtocolError(
                f"slide {slide['index']} animation phase {phase_index} begins before "
                "the previous narrated block ends"
            )
        last_native_position = int(group["last_native_position"])
        for aligned, _marker, effect in native_pairs[
            previous_native_position + 1 : last_native_position + 1
        ]:
            pane_start = float(effect.get("pane_start_seconds") or 0.0)
            if pane_start + 0.002 < phase_pane_start:
                raise ProtocolError(
                    f"slide {slide['index']} native effect {effect['name']!r} crosses "
                    "a narrated phase boundary; keep With Previous effects in one "
                    "simultaneous group"
                )
            start = phase_audio_start + pane_start - phase_pane_start
            end = start + _effect_duration_seconds(effect)
            if start + 0.002 < previous_narration_end or end > phase_audio_end + 0.003:
                raise ProtocolError(
                    f"slide {slide['index']} native effect {effect['name']!r} does not "
                    "fit between the previous narration and the next block narration"
                )
            schedule[int(effect["native_order"])] = {
                "start": start,
                "end": end,
                "sequence_gate": previous_narration_end,
                "phase_index": phase_index,
                "phase_audio_start": phase_audio_start,
                "phase_target_block_indices": list(group["block_indices"]),
                "timing_resolution": timing_resolution,
            }

        group_narration_start = min(
            float(aligned["narration_start"]) for aligned in group["blocks"]
        )
        group_effect_end = max(
            schedule[int(native_pairs[position][2]["native_order"])]["end"]
            for position in group["positions"]
        )
        if group_effect_end > group_narration_start + 0.002:
            raise ProtocolError(
                f"slide {slide['index']} narrated entrance group {group['group']} ends "
                f"at {group_effect_end:.3f}s after its narration begins at "
                f"{group_narration_start:.3f}s; regenerate block audio with animation "
                "pre-roll"
            )
        previous_native_position = last_native_position
        previous_pane_end = phase_pane_end
        previous_narration_end = max(
            float(aligned["spoken_end"])
            for aligned in aligned_blocks
            if int(aligned["index"]) <= max(group["block_indices"])
        )

    # Native-only rows after the final narrated group belong in section post-roll.
    if previous_native_position + 1 < len(native_pairs):
        if timing_parts:
            last_part = max(timing_parts.values(), key=lambda part: float(part["audio_end"]))
            phase_audio_start = float(last_part["audio_end"])
        else:
            phase_audio_start = previous_narration_end
        post_roll = (
            float((timing_section or {}).get("post_roll_seconds") or 0.0)
            if timing_section is not None
            else 0.0
        )
        post_roll_end = phase_audio_start + post_roll
        for _aligned, _marker, effect in native_pairs[previous_native_position + 1 :]:
            start = (
                phase_audio_start
                + float(effect.get("pane_start_seconds") or 0.0)
                - previous_pane_end
            )
            end = start + _effect_duration_seconds(effect)
            if timing_parts and end > post_roll_end + 0.003:
                raise ProtocolError(
                    f"slide {slide['index']} trailing native effect {effect['name']!r} "
                    "does not fit in section animation post-roll"
                )
            schedule[int(effect["native_order"])] = {
                "start": start,
                "end": end,
                "sequence_gate": previous_narration_end,
                "phase_index": len(groups) + 1,
                "phase_audio_start": phase_audio_start,
                "phase_target_block_indices": [],
                "timing_resolution": "native_pane_phase_in_section_postroll",
            }
    return schedule


def _schedule_slide_effects(
    slide: dict[str, object],
    words: list[dict[str, object]],
    timing_section: dict[str, object] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], float]:
    """Resolve Pane effects and selected narration onto a serialized video clock.

    Narration follows the selected block order because Edge word timings are
    generated from that transcript. Native animation type, order, trigger and
    relative delay follow PowerPoint's Animation Pane. Each narrated entrance
    phase is placed before its block audio, and a later phase cannot start until
    the preceding narrated block has finished. Supplemental Spotlight cues alone
    follow their marker's Edge word boundary.
    """
    if str(slide.get("transcript") or "") and not words:
        raise ProtocolError(
            f"timing section {slide['section_id']!r} has narration but no word boundaries"
        )

    cursor = 0
    aligned_blocks: list[dict[str, object]] = []
    for block_index, block in enumerate(slide.get("blocks") or [], start=1):
        block_transcript = str(block.get("transcript") or "")
        if block_transcript:
            word_start, word_end = _align_block(words, cursor, block_transcript)
            cursor = word_end
            narration_start = float(words[word_start].get("start") or 0.0)
            spoken_end = float(words[word_end - 1].get("end") or narration_start)
        else:
            word_start = word_end = cursor
            narration_start = (
                float(words[cursor].get("start") or 0.0)
                if cursor < len(words)
                else (float(words[-1].get("end") or 0.0) if words else 0.0)
            )
            spoken_end = narration_start
        effects = list(block.get("effects") or [])
        markers = list(block.get("markers") or [])
        if len(effects) != len(markers):
            raise ProtocolError(
                f"slide {slide['index']} handle {block['handle']!r} has "
                f"{len(effects)} effects but {len(markers)} timing markers"
            )
        aligned_blocks.append(
            {
                "index": block_index,
                "block": block,
                "word_start": word_start,
                "word_end": word_end,
                "narration_start": narration_start,
                "spoken_end": spoken_end,
                "pairs": list(zip(markers, effects)),
            }
        )

    if cursor != len(words):
        raise ProtocolError(
            f"slide {slide['index']} protocol consumed {cursor}/{len(words)} timing words"
        )

    native_pairs: list[tuple[dict[str, object], dict[str, object], dict[str, object]]] = []
    supplemental_pairs: list[tuple[dict[str, object], dict[str, object], dict[str, object]]] = []
    for aligned in aligned_blocks:
        for marker, effect in aligned["pairs"]:
            target = native_pairs if effect.get("native_present") else supplemental_pairs
            target.append((aligned, marker, effect))
    native_pairs.sort(key=lambda pair: int(pair[2]["native_order"]))
    timing_parts = _align_timing_parts(slide, aligned_blocks, timing_section)
    native_schedule = _serialized_native_schedule(
        slide,
        aligned_blocks,
        native_pairs,
        timing_parts,
        timing_section,
    )

    scheduled: list[dict[str, object]] = []
    effects_by_block: dict[int, list[dict[str, object]]] = {
        int(aligned["index"]): [] for aligned in aligned_blocks
    }
    for aligned, marker, effect in native_pairs + supplemental_pairs:
        block = aligned["block"]
        name = str(effect["name"])
        kind = str(effect.get("kind") or "entrance")
        if kind == "entrance" and name not in VIDEO_EFFECT_SECONDS:
            raise ProtocolError(
                f"slide {slide['index']} uses native effect {name!r} without a video strategy"
            )
        native_present = bool(effect.get("native_present"))
        duration = _effect_duration_seconds(effect)
        trigger = str(effect.get("trigger") or "")
        delay = float(effect.get("delay_seconds") or 0.0)
        pane_start = float(effect.get("pane_start_seconds") or 0.0)
        word_start = int(aligned["word_start"])
        word_end = int(aligned["word_end"])
        block_transcript = str(block.get("transcript") or "")
        scope_text = _oneline(marker.get("scope_text"))
        scope_word_start: int | None = None
        scope_word_end: int | None = None

        if native_present:
            serialized = native_schedule[int(effect["native_order"])]
            start_time = float(serialized["start"])
            end_time = float(serialized["end"])
            requested_start = pane_start
            timing_word = word_start
            timing_source = "animation_pane"
            timing_resolution = str(serialized["timing_resolution"])
            duration_source = (
                "animation_pane"
                if effect.get("duration_seconds") is not None
                else "renderer_default"
            )
        else:
            if name != "Spotlight" or kind != "emphasis":
                raise ProtocolError(
                    f"slide {slide['index']} has non-native non-Spotlight effect {name!r}"
                )
            if not block_transcript:
                raise ProtocolError("supplemental Spotlight requires narrated text")
            requested_start, timing_word = _marker_timing(
                words,
                word_start,
                word_end,
                int(marker["normalized_char"]),
            )
            start_time = requested_start
            timing_source = "edge_word_alignment"
            timing_resolution = "supplemental_spotlight_marker"
            duration_source = "script_scope" if "normalized_end_char" in marker else "renderer_default"

        if "normalized_end_char" in marker:
            if native_present or name != "Spotlight" or kind != "emphasis":
                raise ProtocolError(
                    "spoken-span timing is supported only for supplemental Spotlight"
                )
            scope_end, scope_word_end = _marker_end_timing(
                words,
                word_start,
                word_end,
                int(marker["normalized_end_char"]),
            )
            scope_word_start = timing_word
            if scope_end <= start_time + 0.0005:
                raise ProtocolError("Spotlight spoken scope must end after its start")
            duration = scope_end - start_time
            end_time = scope_end
        elif not native_present:
            end_time = start_time + duration

        scheduled_effect = {
            "block_index": int(aligned["index"]),
            "native_order": int(effect["native_order"]),
            "shape_id": str(block["shape_id"]),
            "handle": str(block["handle"]),
            "name": name,
            "kind": kind,
            "start": round(start_time, 3),
            "end": round(end_time, 3),
            "duration": round(duration, 3),
            "duration_source": duration_source,
            "pane_start": round(pane_start, 3),
            "pane_trigger": trigger,
            "pane_delay": round(delay, 3),
            "sequence_gate": round(
                float(serialized["sequence_gate"]) if native_present else 0.0,
                3,
            ),
            "phase_index": int(serialized["phase_index"]) if native_present else None,
            "phase_audio_start": (
                round(float(serialized["phase_audio_start"]), 3)
                if native_present
                else None
            ),
            "phase_target_block_indices": (
                list(serialized.get("phase_target_block_indices") or [])
                if native_present
                else []
            ),
            "requested_start": round(requested_start, 3),
            "timing_source": timing_source,
            "timing_resolution": timing_resolution,
            "intent_source": marker.get("source"),
            "word_start": timing_word,
            "word_end": max(0, word_end - 1),
            "scope_text": scope_text or None,
            "scope_word_start": scope_word_start,
            "scope_word_end": scope_word_end,
            "spoken_end": round(float(aligned["spoken_end"]), 3),
            "simultaneous_group": int(
                effect.get("simultaneous_group") or effect["native_order"]
            ),
            "click_group": int(effect.get("click_group") or effect["native_order"]),
        }
        scheduled.append(scheduled_effect)
        effects_by_block[int(aligned["index"])].append(scheduled_effect)

    sequence_blocks: list[dict[str, object]] = []
    for aligned in aligned_blocks:
        block = aligned["block"]
        block_effects = effects_by_block[int(aligned["index"])]
        timing_part = timing_parts.get(int(aligned["index"]))
        sequence_blocks.append(
            {
                "index": int(aligned["index"]),
                "handle": str(block["handle"]),
                "shape_id": str(block["shape_id"]),
                "narrated": bool(str(block.get("transcript") or "")),
                "role": (
                    "narrated_element"
                    if str(block.get("transcript") or "")
                    else "silent_native_element"
                ),
                "word_start": int(aligned["word_start"]),
                "word_end": max(0, int(aligned["word_end"]) - 1),
                "narration_start": round(float(aligned["narration_start"]), 3),
                "spoken_end": round(float(aligned["spoken_end"]), 3),
                "release_before": round(float(aligned["narration_start"]), 3),
                "release": round(float(aligned["spoken_end"]), 3),
                "pre_roll_seconds": (
                    round(float(timing_part["pre_roll_seconds"]), 3)
                    if timing_part is not None
                    else None
                ),
                "pre_roll_start": (
                    round(float(timing_part["pre_roll_start"]), 3)
                    if timing_part is not None
                    else None
                ),
                "pre_roll_end": (
                    round(float(timing_part["pre_roll_end"]), 3)
                    if timing_part is not None
                    else None
                ),
                "effect_count": len(block_effects),
                "effects": [
                    {
                        key: effect[key]
                        for key in (
                            "name",
                            "kind",
                            "start",
                            "end",
                            "pane_trigger",
                            "pane_delay",
                            "sequence_gate",
                            "phase_index",
                            "phase_audio_start",
                            "phase_target_block_indices",
                            "timing_source",
                            "timing_resolution",
                            "duration_source",
                        )
                    }
                    for effect in block_effects
                ],
            }
        )

    schedule_end = max(
        [float(item["spoken_end"]) for item in aligned_blocks]
        + [float(effect["end"]) for effect in scheduled]
        + [0.0]
    )
    return scheduled, sequence_blocks, round(schedule_end, 3)


def build_pptx_animation_manifest(
    protocol: dict[str, object],
    word_timings: Path,
) -> dict[str, object]:
    """Map the reconciled PPTX protocol to Edge word-aligned render effects."""
    try:
        timing_payload = json.loads(word_timings.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"could not read word timings {word_timings}: {exc}") from exc
    timing_sections = timing_payload.get("sections") or []
    slides = protocol.get("slides") or []
    if len(timing_sections) != len(slides):
        raise ProtocolError(
            f"word timing section count {len(timing_sections)} != slide count {len(slides)}"
        )

    manifest_slides: list[dict[str, object]] = []
    effect_count = 0
    for slide, section in zip(slides, timing_sections):
        section_id = str(slide["section_id"])
        if str(section.get("id") or "") != section_id:
            raise ProtocolError(
                f"timing section {section.get('id')!r} != PPTX section {section_id!r}"
            )
        words = section.get("words") or []
        if not isinstance(words, list):
            raise ProtocolError(f"timing section {section_id!r} words must be an array")
        scheduled, sequence_blocks, schedule_end = _schedule_slide_effects(
            slide,
            words,
            timing_section=section,
        )
        effects_out: list[dict[str, object]] = []
        for effect in scheduled:
            if effect["kind"] != "entrance":
                continue
            effects_out.append(
                {
                    **effect,
                    "order": len(effects_out) + 1,
                    "locator": effect["handle"],
                }
            )
            effect_count += 1
        manifest_slides.append(
            {
                "index": int(slide["index"]),
                "id": section_id,
                "stable_slide_id": slide["stable_slide_id"],
                "schedule_policy": "narrated_block_serial_v1",
                "schedule_end": schedule_end,
                "sequence_blocks": sequence_blocks,
                "effect_count": len(effects_out),
                "effects": effects_out,
            }
        )

    return {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "created_at": _utc_now(),
        "source_kind": "pptx",
        "source_pptx": protocol["source_pptx"],
        "source_sha256": protocol["source_sha256"],
        "protocol_schema_version": protocol["schema_version"],
        "word_timings": str(word_timings.resolve()),
        "slide_count": len(manifest_slides),
        "effect_count": effect_count,
        "timing_source": "animation_pane_serialized_to_edge_narration",
        "slides": manifest_slides,
    }


def build_pptx_visual_cues(
    protocol: dict[str, object],
    word_timings: Path,
) -> dict[str, object]:
    """Build deterministic spotlight cues from native emphasis markers."""
    try:
        timing_payload = json.loads(word_timings.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"could not read word timings {word_timings}: {exc}") from exc
    timing_sections = timing_payload.get("sections") or []
    slides = protocol.get("slides") or []
    if len(timing_sections) != len(slides):
        raise ProtocolError(
            f"word timing section count {len(timing_sections)} != slide count {len(slides)}"
        )

    cue_slides: list[dict[str, object]] = []
    cue_count = 0
    for slide, section in zip(slides, timing_sections):
        section_id = str(slide["section_id"])
        if str(section.get("id") or "") != section_id:
            raise ProtocolError(
                f"timing section {section.get('id')!r} != PPTX section {section_id!r}"
            )
        words = section.get("words") or []
        if not isinstance(words, list):
            raise ProtocolError(f"timing section {section_id!r} words must be an array")
        scheduled, _, _ = _schedule_slide_effects(
            slide,
            words,
            timing_section=section,
        )
        blocks_by_index = {
            index: block
            for index, block in enumerate(slide.get("blocks") or [], start=1)
        }
        cues: list[dict[str, object]] = []
        for scheduled_effect in scheduled:
            if scheduled_effect["kind"] != "emphasis":
                continue
            block = blocks_by_index[int(scheduled_effect["block_index"])]
            bbox = block.get("bbox")
            if not isinstance(bbox, list) or len(bbox) != 4:
                raise ProtocolError(
                    f"slide {slide['index']} spotlight target {block['handle']!r} "
                    "has no usable top-level PowerPoint geometry"
                )
            point = [
                round(float(bbox[0]) + float(bbox[2]) / 2.0, 6),
                round(float(bbox[1]) + float(bbox[3]) / 2.0, 6),
            ]
            target = f"pptx:s{int(slide['index']):02d}_sh{block['shape_id']}"
            cues.append(
                {
                    "start": scheduled_effect["start"],
                    "end": scheduled_effect["end"],
                    "type": "highlight",
                    "box": [round(float(value), 6) for value in bbox],
                    "point": point,
                    "target": target,
                    "target_role": "content",
                    "target_source": "pptx",
                    "semantic_target": target,
                    "semantic_source": "pptx",
                    "semantic_box": [round(float(value), 6) for value in bbox],
                    "geometry_target": target,
                    "geometry_source": "pptx",
                    "geometry_box": [round(float(value), 6) for value in bbox],
                    "geometry_matched": True,
                    "geometry_match_iou": 1.0,
                    "confidence": 1.0,
                    "timing_source": scheduled_effect["timing_source"],
                    "timing_resolution": scheduled_effect["timing_resolution"],
                    "duration_source": scheduled_effect["duration_source"],
                    "intent_source": scheduled_effect["intent_source"],
                    "text": scheduled_effect.get("scope_text") or block["transcript"],
                    "scope_text": scheduled_effect.get("scope_text"),
                    "scope_word_start": scheduled_effect.get("scope_word_start"),
                    "scope_word_end": scheduled_effect.get("scope_word_end"),
                    "marker_name": "Spotlight",
                    "handle": block["handle"],
                    "shape_id": block["shape_id"],
                }
            )
            cue_count += 1
        cue_slides.append(
            {
                "index": int(slide["index"]),
                "id": section_id,
                "cues": cues,
            }
        )
    return {
        "schema_version": "paper2video_visual_cues.v3",
        "cue_shape": "semantic_box",
        "source_kind": "pptx_protocol",
        "source_sha256": protocol["source_sha256"],
        "cue_count": cue_count,
        "slides": cue_slides,
    }


def build_pptx_visual_cue_plan(
    visual_cues: dict[str, object],
) -> dict[str, object]:
    """Build a deterministic cue-plan sidecar for native emphasis effects.

    The semantic cue planner is intentionally not a runtime dependency of the
    editable PPTX route. Native emphasis rows already identify an exact PPTX
    shape, while Author Notes and Edge boundaries already identify the exact
    time. This adapter records that fully resolved mapping in the same plan
    shape consumed by ``build_timeline.py`` and strict QA.
    """
    slides_out: list[dict[str, object]] = []
    cue_count = 0
    for slide in visual_cues.get("slides") or []:
        chunks: list[dict[str, object]] = []
        for chunk_index, cue in enumerate(slide.get("cues") or [], start=1):
            start = float(cue["start"])
            end = float(cue["end"])
            timing_source = str(cue.get("timing_source") or "edge_word_alignment")
            target = str(cue.get("target") or "")
            box = cue.get("box")
            point = cue.get("point")
            chunks.append(
                {
                    "chunk_index": chunk_index,
                    "chunk_id": f"s{int(slide['index']):02d}_c{chunk_index:02d}",
                    "text": str(cue.get("text") or ""),
                    "start": round(start, 3),
                    "end": round(end, 3),
                    "seconds": round(end - start, 3),
                    "timing_source": timing_source,
                    "duration_source": cue.get("duration_source"),
                    "scope_text": cue.get("scope_text"),
                    "scope_word_start": cue.get("scope_word_start"),
                    "scope_word_end": cue.get("scope_word_end"),
                    "timing": {
                        "method": timing_source,
                        "score": 1.0,
                        "start": round(start, 3),
                        "end": round(end, 3),
                    },
                    "accepted": True,
                    "confidence": 1.0,
                    "reason": "native_pptx_emphasis",
                    "anchor_required": False,
                    "anchor_matched": True,
                    "target": target,
                    "target_role": cue.get("target_role") or "content",
                    "target_source": "pptx",
                    "semantic_target": cue.get("semantic_target") or target,
                    "semantic_role": cue.get("target_role") or "content",
                    "semantic_source": "pptx",
                    "semantic_box": cue.get("semantic_box") or box,
                    "geometry_target": cue.get("geometry_target") or target,
                    "geometry_role": cue.get("target_role") or "content",
                    "geometry_source": "pptx",
                    "geometry_box": cue.get("geometry_box") or box,
                    "geometry_matched": True,
                    "geometry_match_iou": 1.0,
                    "region_box": box,
                    "point": point,
                    "handle": cue.get("handle"),
                    "shape_id": cue.get("shape_id"),
                }
            )
            cue_count += 1
        slides_out.append(
            {
                "index": int(slide["index"]),
                "id": str(slide["id"]),
                "timing_source": (
                    chunks[0]["timing_source"] if chunks else "edge_word_alignment"
                ),
                "cue_count": len(chunks),
                "chunks": chunks,
            }
        )
    return {
        "schema_version": "paper2video_visual_cue_plan.v1",
        "source_kind": "pptx_protocol",
        "source_sha256": visual_cues.get("source_sha256"),
        "min_confidence": 1.0,
        "strict_gate": False,
        "cue_count": cue_count,
        "slides": slides_out,
    }


def _copy_zipinfo(info: ZipInfo) -> ZipInfo:
    copied = ZipInfo(info.filename, date_time=info.date_time)
    copied.compress_type = ZIP_DEFLATED
    copied.comment = info.comment
    copied.extra = info.extra
    copied.create_system = info.create_system
    copied.create_version = info.create_version
    copied.extract_version = info.extract_version
    copied.flag_bits = info.flag_bits
    copied.volume = info.volume
    copied.internal_attr = info.internal_attr
    copied.external_attr = info.external_attr
    return copied


def write_reveal_variant(
    source_pptx: Path,
    slide_shape_ids: list[list[str]],
    reveal_count: int,
    output_pptx: Path,
) -> None:
    """Write a temporary PPTX where each slide reveals its first N targets."""
    source_pptx = source_pptx.resolve()
    with ZipFile(source_pptx) as source:
        slide_refs = presentation_slides(source)
        if len(slide_refs) != len(slide_shape_ids):
            raise ProtocolError(
                f"reveal plan has {len(slide_shape_ids)} slides, PPTX has {len(slide_refs)}"
            )
        replacements: dict[str, bytes] = {}
        for ref, targets in zip(slide_refs, slide_shape_ids):
            part = str(ref["part"])
            root = etree.fromstring(source.read(part))
            remove_ids = set(targets[max(0, reveal_count):])
            for shape_id in remove_ids:
                matches = root.xpath(
                    f'.//p:spTree//p:cNvPr[@id="{shape_id}"]', namespaces=NS
                )
                if len(matches) != 1:
                    raise ProtocolError(
                        f"slide {ref['index']} reveal target {shape_id!r} resolved "
                        f"to {len(matches)} shapes"
                    )
                top = matches[0]
                while top.getparent() is not None and top.getparent().tag != f"{{{P_NS}}}spTree":
                    top = top.getparent()
                if top.getparent() is None:
                    raise ProtocolError(
                        f"slide {ref['index']} reveal target {shape_id!r} is not under p:spTree"
                    )
                top.getparent().remove(top)
            for timing in root.findall(f"{{{P_NS}}}timing"):
                root.remove(timing)
            replacements[part] = etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=output_pptx.stem + ".",
            suffix=".pptx",
            dir=output_pptx.parent,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with ZipFile(temporary_path, "w") as destination:
                for info in source.infolist():
                    destination.writestr(
                        _copy_zipinfo(info),
                        replacements.get(info.filename, source.read(info.filename)),
                    )
            shutil.move(temporary_path, output_pptx)
        finally:
            temporary_path.unlink(missing_ok=True)
