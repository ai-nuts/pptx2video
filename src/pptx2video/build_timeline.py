#!/usr/bin/env python3
"""Build the unified pptx2video timeline contract.

The timeline is the sidecar that ties narration text, audio windows, subtitle
cues, and visual-highlight targets to the same chunk ids. It does not render or
modify video; it only normalizes sidecar files produced by the existing
pptx2video pipeline into one auditable contract.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .audio_files import (
    is_portable_audio_filename,
    load_audio_manifest_file_map,
    resolve_audio_path,
)

SCHEMA_VERSION = "paper2video_timeline.v1"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"[build_timeline] file not found: {path}")
    except json.JSONDecodeError as exc:
        sys.exit(f"[build_timeline] invalid JSON {path}: {exc}")


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def safe_id(raw: object, fallback: str) -> str:
    text = str(raw or "").strip()
    if not text:
        text = fallback
    text = re.sub(r"[^A-Za-z0-9_.-]+", "-", text).strip("-")
    return text or fallback


def load_script_sections(path: Path) -> list[dict[str, Any]]:
    payload = read_json(path)
    sections = payload.get("sections") if isinstance(payload, dict) else None
    if not isinstance(sections, list) or not sections:
        sys.exit(f"[build_timeline] script JSON has no sections array: {path}")
    out: list[dict[str, Any]] = []
    for idx, sec in enumerate(sections, start=1):
        if not isinstance(sec, dict):
            sys.exit(f"[build_timeline] script section {idx} is not an object")
        sid = safe_id(sec.get("id"), f"slide-{idx:02d}")
        out.append({
            "index": idx,
            "id": sid,
            "heading": str(sec.get("heading") or sec.get("title") or sid),
            "text": str(sec.get("text") or sec.get("script") or "").strip(),
        })
    return out


def load_section_mapping(path: Path | None) -> tuple[dict[str, str], "OrderedDict[str, list[str]]"]:
    """Return primary slide-key -> section id plus explicit section groups.

    Accepted JSON shapes:

    1. {"problem": [3, 4], "method": ["05_architecture"]}
    2. {"03_sequence_evolution": "problem", "4": "problem"}

    Shape 1 may express overlapping poster sections, for example both
    "problem" and "motivation" can contain slide 3. Shape 2 is a single-owner
    mapping and is used only to give each slide/chunk a primary section id.
    """
    if path is None:
        return {}, OrderedDict()
    payload = read_json(path)
    if not isinstance(payload, dict):
        sys.exit(f"[build_timeline] section map must be a JSON object: {path}")
    primary: dict[str, str] = {}
    groups: "OrderedDict[str, list[str]]" = OrderedDict()
    for key, value in payload.items():
        if isinstance(value, list):
            section_id = safe_id(key, "section")
            groups.setdefault(section_id, [])
            for item in value:
                groups[section_id].append(str(item))
                primary.setdefault(str(item), section_id)
                try:
                    primary.setdefault(str(int(item)), section_id)
                except (TypeError, ValueError):
                    pass
        else:
            section_id = safe_id(value, "section")
            primary[str(key)] = section_id
            try:
                primary[str(int(key))] = section_id
            except (TypeError, ValueError):
                pass
    return primary, groups


def load_by_slide(path: Path | None, field: str) -> dict[int, dict[str, Any]]:
    if path is None:
        return {}
    payload = read_json(path)
    slides = payload.get("slides") if isinstance(payload, dict) else None
    if not isinstance(slides, list):
        sys.exit(f"[build_timeline] {field} must contain a slides array: {path}")
    out: dict[int, dict[str, Any]] = {}
    for slide in slides:
        if not isinstance(slide, dict):
            continue
        try:
            idx = int(slide.get("index"))
        except (TypeError, ValueError):
            continue
        out[idx] = slide
    return out


def parse_vtt_time(raw: str) -> float:
    match = re.match(r"(?:(\d+):)?(\d+):(\d+(?:\.\d+)?)", raw.strip())
    if not match:
        raise ValueError(raw)
    hours = int(match.group(1) or 0)
    minutes = int(match.group(2))
    seconds = float(match.group(3))
    return hours * 3600 + minutes * 60 + seconds


def load_vtt(path: Path | None) -> list[dict[str, Any]]:
    if path is None:
        return []
    if not path.is_file():
        sys.exit(f"[build_timeline] captions VTT not found: {path}")
    text = path.read_text(encoding="utf-8").replace("\r\n", "\n")
    cues: list[dict[str, Any]] = []
    for block in [part.strip() for part in text.split("\n\n") if part.strip()]:
        if block == "WEBVTT" or block.startswith("WEBVTT\n"):
            continue
        lines = block.splitlines()
        time_idx = next((i for i, line in enumerate(lines) if "-->" in line), None)
        if time_idx is None:
            continue
        left, right = lines[time_idx].split("-->", 1)
        try:
            start = parse_vtt_time(left.strip().split()[0])
            end = parse_vtt_time(right.strip().split()[0])
        except ValueError:
            continue
        body = "\n".join(lines[time_idx + 1:]).strip()
        if body and end > start:
            cues.append({"start": round(start, 3), "end": round(end, 3), "text": body})
    return cues


def overlap_captions(captions: list[dict[str, Any]], start: float, end: float) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for cue in captions:
        cue_start = float(cue["start"])
        cue_end = float(cue["end"])
        if cue_end <= start or cue_start >= end:
            continue
        out.append({
            "start": round(max(cue_start, start), 3),
            "end": round(min(cue_end, end), 3),
            "chunk_start": round(max(cue_start, start) - start, 3),
            "chunk_end": round(min(cue_end, end) - start, 3),
            "text": cue["text"],
        })
    return out


def cue_lookup(visual_slide: dict[str, Any] | None) -> dict[tuple[float, float, str], dict[str, Any]]:
    out: dict[tuple[float, float, str], dict[str, Any]] = {}
    if not visual_slide:
        return out
    for cue in visual_slide.get("cues") or []:
        if not isinstance(cue, dict):
            continue
        try:
            start = round(float(cue.get("start")), 3)
            end = round(float(cue.get("end")), 3)
        except (TypeError, ValueError):
            continue
        target = str(cue.get("target") or "")
        out[(start, end, target)] = cue
    return out


def slide_lookup_keys(slide: dict[str, Any]) -> list[str]:
    idx = int(slide["index"])
    return [str(slide.get("id") or ""), str(idx), f"{idx:02d}"]


def rebuild_sections_from_groups(
    groups: "OrderedDict[str, list[str]]",
    slides_out: list[dict[str, Any]],
) -> "OrderedDict[str, dict[str, Any]]":
    lookup: dict[str, dict[str, Any]] = {}
    for slide in slides_out:
        for key in slide_lookup_keys(slide):
            if key:
                lookup[key] = slide

    rebuilt: "OrderedDict[str, dict[str, Any]]" = OrderedDict()
    for section_id, keys in groups.items():
        selected: list[dict[str, Any]] = []
        seen: set[int] = set()
        for raw_key in keys:
            key = str(raw_key)
            slide = lookup.get(key)
            if slide is None:
                try:
                    slide = lookup.get(str(int(key)))
                except ValueError:
                    slide = None
            if slide is None:
                sys.exit(f"[build_timeline] section map references unknown slide {raw_key!r} for section {section_id!r}")
            idx = int(slide["index"])
            if idx in seen:
                continue
            seen.add(idx)
            selected.append(slide)
        if not selected:
            continue
        rebuilt[section_id] = {
            "id": section_id,
            "slide_indices": [int(slide["index"]) for slide in selected],
            "slide_ids": [str(slide["id"]) for slide in selected],
            "chunk_ids": [chunk["id"] for slide in selected for chunk in slide.get("chunks", [])],
            "start": round(min(float(slide["segment"]["start"]) for slide in selected), 3),
            "end": round(max(float(slide["segment"]["end"]) for slide in selected), 3),
        }
        rebuilt[section_id]["seconds"] = round(float(rebuilt[section_id]["end"]) - float(rebuilt[section_id]["start"]), 3)
    return rebuilt


def build_timeline(
    *,
    script_json: Path,
    duration_report: Path,
    visual_cue_plan: Path | None,
    visual_cues: Path | None,
    captions_vtt: Path | None,
    audio_dir: Path | None,
    video: Path | None,
    section_map_path: Path | None,
) -> dict[str, Any]:
    script_sections = load_script_sections(script_json)
    report = read_json(duration_report)
    report_slides = report.get("slides") if isinstance(report, dict) else None
    if not isinstance(report_slides, list) or not report_slides:
        sys.exit(f"[build_timeline] duration report has no slides array: {duration_report}")
    if len(report_slides) != len(script_sections):
        sys.exit(
            "[build_timeline] script/duration slide count mismatch: "
            f"{len(script_sections)} script sections vs {len(report_slides)} report slides"
        )

    cue_plan_by_slide = load_by_slide(visual_cue_plan, "visual cue plan")
    visual_cues_by_slide = load_by_slide(visual_cues, "visual cues")
    captions = load_vtt(captions_vtt)
    section_map, explicit_section_groups = load_section_mapping(section_map_path)
    try:
        audio_file_map = load_audio_manifest_file_map(audio_dir) if audio_dir else {}
    except ValueError as exc:
        sys.exit(f"[build_timeline] {exc}")

    start_pad = float(report.get("start_pad") or 0.0)
    pad_tail = float(report.get("pad_tail") or 0.0)
    cursor = start_pad
    slides_out: list[dict[str, Any]] = []
    flat_chunks: list[dict[str, Any]] = []
    sections: "OrderedDict[str, dict[str, Any]]" = OrderedDict()

    for idx, (script_sec, slide_report) in enumerate(zip(script_sections, report_slides), start=1):
        slide_index = int(slide_report.get("index") or idx)
        if slide_index != idx:
            sys.exit(f"[build_timeline] non-sequential slide index in duration report: {slide_index} at position {idx}")
        slide_id = script_sec["id"]
        report_audio = str(slide_report.get("audio") or "")
        if audio_dir is not None:
            if report_audio:
                if not is_portable_audio_filename(report_audio):
                    sys.exit(f"[build_timeline] non-portable audio filename: {report_audio!r}")
                audio_file = str((audio_dir / report_audio).resolve())
            else:
                try:
                    audio_file = str(
                        resolve_audio_path(
                            audio_dir,
                            str(slide_id),
                            file_map=audio_file_map,
                        ).resolve()
                    )
                except ValueError as exc:
                    sys.exit(f"[build_timeline] {exc}")
        else:
            audio_file = report_audio
        section_id = (
            section_map.get(slide_id)
            or section_map.get(str(slide_index))
            or section_map.get(f"{slide_index:02d}")
            or slide_id
        )
        audio_seconds = float(slide_report.get("audio_seconds") or 0.0)
        segment_seconds = float(slide_report.get("segment_seconds") or audio_seconds + pad_tail)
        segment_start = round(cursor, 3)
        segment_end = round(cursor + segment_seconds, 3)
        audio_start = segment_start
        audio_end = round(segment_start + audio_seconds, 3)
        cursor += segment_seconds

        plan_slide = cue_plan_by_slide.get(slide_index, {})
        chunks = plan_slide.get("chunks") if isinstance(plan_slide, dict) else None
        if not isinstance(chunks, list) or not chunks:
            chunks = [{
                "chunk_index": 1,
                "text": script_sec["text"],
                "start": 0.0,
                "end": audio_seconds,
                "seconds": audio_seconds,
                "timing_source": "slide_audio",
                "accepted": False,
                "reason": "visual_cue_plan_missing",
            }]

        visual_lookup = cue_lookup(visual_cues_by_slide.get(slide_index))
        chunk_entries: list[dict[str, Any]] = []
        for chunk in chunks:
            if not isinstance(chunk, dict):
                continue
            chunk_num = int(chunk.get("chunk_index") or len(chunk_entries) + 1)
            local_start = round(float(chunk.get("start") or 0.0), 3)
            local_end = round(float(chunk.get("end") or local_start), 3)
            local_start = max(0.0, min(local_start, segment_seconds))
            local_end = max(local_start, min(local_end, segment_seconds))
            global_start = round(segment_start + local_start, 3)
            global_end = round(segment_start + local_end, 3)
            target = str(chunk.get("target") or "")
            visual_cue = visual_lookup.get((round(local_start, 3), round(local_end, 3), target), {})
            visual_payload = None
            if chunk.get("accepted") or visual_cue:
                visual_payload = {
                    "accepted": bool(chunk.get("accepted")),
                    "anchor_id": chunk.get("anchor_id") or visual_cue.get("anchor_id"),
                    "anchor_matched": bool(chunk.get("anchor_matched") or visual_cue.get("anchor_matched")),
                    "target": target or visual_cue.get("target"),
                    "target_role": chunk.get("target_role") or visual_cue.get("target_role"),
                    "target_source": chunk.get("target_source") or visual_cue.get("target_source"),
                    "semantic_target": chunk.get("semantic_target") or visual_cue.get("semantic_target") or target or visual_cue.get("target"),
                    "semantic_original_target": chunk.get("semantic_original_target") or visual_cue.get("semantic_original_target"),
                    "semantic_promoted": bool(chunk.get("semantic_promoted") or visual_cue.get("semantic_promoted")),
                    "semantic_promotion": chunk.get("semantic_promotion") or visual_cue.get("semantic_promotion"),
                    "semantic_role": chunk.get("semantic_role") or visual_cue.get("semantic_role") or chunk.get("target_role") or visual_cue.get("target_role"),
                    "semantic_source": chunk.get("semantic_source") or visual_cue.get("semantic_source") or chunk.get("target_source") or visual_cue.get("target_source"),
                    "semantic_box": chunk.get("semantic_box") or visual_cue.get("semantic_box"),
                    "geometry_target": chunk.get("geometry_target") or visual_cue.get("geometry_target"),
                    "geometry_role": chunk.get("geometry_role") or visual_cue.get("geometry_role"),
                    "geometry_source": chunk.get("geometry_source") or visual_cue.get("geometry_source"),
                    "geometry_box": chunk.get("geometry_box") or visual_cue.get("geometry_box") or visual_cue.get("box"),
                    "geometry_matched": bool(chunk.get("geometry_matched") or visual_cue.get("geometry_matched")),
                    "geometry_match_score": chunk.get("geometry_match_score") or visual_cue.get("geometry_match_score"),
                    "geometry_match_iou": chunk.get("geometry_match_iou") or visual_cue.get("geometry_match_iou"),
                    "geometry_semantic_coverage": chunk.get("geometry_semantic_coverage") or visual_cue.get("geometry_semantic_coverage"),
                    "geometry_coverage": chunk.get("geometry_coverage") or visual_cue.get("geometry_coverage"),
                    "geometry_match_reason": chunk.get("geometry_match_reason") or visual_cue.get("geometry_match_reason"),
                    "point": chunk.get("point") or visual_cue.get("point"),
                    "region_box": chunk.get("geometry_box") or visual_cue.get("geometry_box") or chunk.get("region_box") or visual_cue.get("box"),
                    "confidence": chunk.get("confidence") or visual_cue.get("confidence"),
                    "reason": chunk.get("reason"),
                    "timing": chunk.get("timing"),
                    "cue": visual_cue or None,
                }
            chunk_id = f"{section_id}.s{slide_index:02d}.c{chunk_num:02d}"
            subtitle_cues = overlap_captions(captions, global_start, global_end)
            chunk_entry = {
                "id": chunk_id,
                "section_id": section_id,
                "slide_id": slide_id,
                "slide_index": slide_index,
                "chunk_index": chunk_num,
                "text": str(chunk.get("text") or "").strip(),
                "timing_source": chunk.get("timing_source") or plan_slide.get("timing_source") or "unknown",
                "local_start": local_start,
                "local_end": local_end,
                "start": global_start,
                "end": global_end,
                "seconds": round(global_end - global_start, 3),
                "audio": {
                    "file": audio_file,
                    "slide_start": audio_start,
                    "slide_end": audio_end,
                },
                "subtitles": subtitle_cues,
                "visual_cue": visual_payload,
            }
            chunk_entries.append(chunk_entry)
            flat_chunks.append(chunk_entry)

        slide_entry = {
            "index": slide_index,
            "id": slide_id,
            "heading": script_sec["heading"],
            "section_id": section_id,
            "audio": {
                "file": audio_file,
                "seconds": round(audio_seconds, 3),
                "start": audio_start,
                "end": audio_end,
            },
            "segment": {
                "start": segment_start,
                "end": segment_end,
                "seconds": round(segment_seconds, 3),
                "pad_tail": round(max(0.0, segment_seconds - audio_seconds), 3),
            },
            "text": script_sec["text"],
            "chunks": chunk_entries,
        }
        slides_out.append(slide_entry)

        sec = sections.setdefault(section_id, {
            "id": section_id,
            "slide_indices": [],
            "slide_ids": [],
            "chunk_ids": [],
            "start": segment_start,
            "end": segment_end,
        })
        sec["slide_indices"].append(slide_index)
        sec["slide_ids"].append(slide_id)
        sec["chunk_ids"].extend([chunk["id"] for chunk in chunk_entries])
        sec["start"] = min(float(sec["start"]), segment_start)
        sec["end"] = max(float(sec["end"]), segment_end)

    for sec in sections.values():
        sec["seconds"] = round(float(sec["end"]) - float(sec["start"]), 3)
        sec["start"] = round(float(sec["start"]), 3)
        sec["end"] = round(float(sec["end"]), 3)

    if explicit_section_groups:
        sections = rebuild_sections_from_groups(explicit_section_groups, slides_out)

    actual_seconds = report.get("actual_seconds")
    if actual_seconds is None:
        actual_seconds = round(cursor, 3)
    return {
        "schema_version": SCHEMA_VERSION,
        "created_at": utc_now(),
        "resources": {
            "script_json": str(script_json),
            "duration_report": str(duration_report),
            "visual_cue_plan": str(visual_cue_plan) if visual_cue_plan else None,
            "visual_cues": str(visual_cues) if visual_cues else None,
            "captions_vtt": str(captions_vtt) if captions_vtt else None,
            "audio_dir": str(audio_dir) if audio_dir else None,
            "video": str(video) if video else str(report.get("output") or ""),
            "section_map": str(section_map_path) if section_map_path else None,
        },
        "start_pad": round(start_pad, 3),
        "pad_tail": round(pad_tail, 3),
        "actual_seconds": round(float(actual_seconds), 3),
        "sections": list(sections.values()),
        "slides": slides_out,
        "chunks": flat_chunks,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--script-json", required=True, type=Path)
    parser.add_argument("--duration-report", required=True, type=Path)
    parser.add_argument("--visual-cue-plan", type=Path)
    parser.add_argument("--visual-cues", type=Path)
    parser.add_argument("--captions-vtt", type=Path)
    parser.add_argument("--audio-dir", type=Path)
    parser.add_argument("--video", type=Path)
    parser.add_argument("--section-map", type=Path,
                        help="Optional JSON that maps timeline slides to canonical poster/blog sections.")
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()

    timeline = build_timeline(
        script_json=args.script_json.resolve(),
        duration_report=args.duration_report.resolve(),
        visual_cue_plan=args.visual_cue_plan.resolve() if args.visual_cue_plan else None,
        visual_cues=args.visual_cues.resolve() if args.visual_cues else None,
        captions_vtt=args.captions_vtt.resolve() if args.captions_vtt else None,
        audio_dir=args.audio_dir.resolve() if args.audio_dir else None,
        video=args.video.resolve() if args.video else None,
        section_map_path=args.section_map.resolve() if args.section_map else None,
    )
    write_json(args.out.resolve(), timeline)
    print(
        f"[build_timeline] wrote {args.out.resolve()} "
        f"({len(timeline['sections'])} sections, {len(timeline['slides'])} slides, "
        f"{len(timeline['chunks'])} chunks)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
