#!/usr/bin/env python3
"""Render an edited protocol PPTX to a strictly checked video without an LLM."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from .audio_files import (
    legacy_audio_filename,
    load_audio_manifest_file_map,
    resolve_audio_path,
)
from .animation_order_authority import select_animation_order_policy

from .editable_pptx import (
    ProtocolError,
    analyze_animation_order_conflicts,
    apply_confirmed_animation_pane_order,
    apply_user_script,
    build_pptx_animation_manifest,
    build_pptx_visual_cue_plan,
    build_pptx_visual_cues,
    detect_pptx_changes,
    extract_protocol,
    file_sha256,
    normalize_animation_click_groups,
    normalize_author_notes_authority,
    script_from_protocol,
    write_protocol_to_pptx,
    write_json,
)
from .narration_regeneration import DEFAULT_MODEL, regenerate_changed_narration
from .render_parameters import (
    DEFAULT_FPS,
    DEFAULT_PAD_TAIL,
    DEFAULT_START_PAD,
    validate_render_timing,
)


def _module_command(module: str, *arguments: object) -> list[str]:
    """Build a subprocess command against this installed package."""
    return [
        sys.executable,
        "-m",
        f"pptx2video.{module}",
        *(str(argument) for argument in arguments),
    ]


def _run(command: list[str]) -> None:
    print("[render_edited_pptx] $ " + " ".join(command), flush=True)
    subprocess.run(command, check=True)


def _should_apply_confirmed_animation_pane_order(
    selected_policy: str,
    report: dict[str, object],
) -> bool:
    """Return whether the user explicitly chose Pane over a real conflict."""
    return (
        selected_policy == "animation-pane"
        and int(report.get("conflict_count") or 0) > 0
        and str(report.get("selection_source") or "")
        in {
            "command_line",
            "interactive",
            "command_line_identity_sequence",
        }
    )


def _animation_pane_slide_indices(report: dict[str, object]) -> set[int]:
    """Return conflict slides explicitly resolved in favor of the Pane.

    New per-slide decisions are authoritative. The conflict-slide fallback
    keeps reports produced by the older all-or-nothing resolver usable.
    """
    decisions = report.get("slide_decisions") or []
    if decisions:
        return {
            int(decision.get("slide_index") or 0)
            for decision in decisions
            if str(decision.get("selected_policy") or "") == "animation-pane"
            and int(decision.get("slide_index") or 0) > 0
        }
    return {
        int(slide.get("index") or 0)
        for slide in report.get("slides") or []
        if int(slide.get("index") or 0) > 0
    }


def _prepare_protocol_for_render(
    source_pptx: Path,
    staged_pptx: Path,
    *,
    ids_from_script: Path | None,
    script_json: Path | None,
    baseline_pptx: Path | None,
    narration_mode: str,
    regeneration_model: str,
    narration_order_policy: str = "geometry",
    click_group_policy: str = "normalize",
) -> dict[str, object]:
    """Resolve every narration authority in a temporary PPTX before prompting."""
    if narration_mode == "regenerate" and baseline_pptx is None:
        raise ProtocolError("--narration-mode regenerate requires --baseline-pptx")
    if narration_mode == "regenerate" and script_json is not None:
        raise ProtocolError(
            "--narration-mode regenerate and --script-json are mutually exclusive"
        )

    authority_report = normalize_author_notes_authority(
        source_pptx,
        staged_pptx,
        narration_order_policy=narration_order_policy,
    )
    protocol = extract_protocol(
        staged_pptx,
        ids_from_script=ids_from_script or script_json,
        narration_order_policy=narration_order_policy,
    )
    change_report = None
    if baseline_pptx is not None:
        baseline_pptx = baseline_pptx.resolve()
        if not baseline_pptx.is_file():
            raise ProtocolError(f"baseline PPTX not found: {baseline_pptx}")
        change_report = detect_pptx_changes(baseline_pptx, staged_pptx)

    regeneration_report = None
    if narration_mode == "regenerate":
        assert change_report is not None
        protocol, regeneration_report = regenerate_changed_narration(
            protocol,
            change_report,
            model=regeneration_model,
        )
        write_protocol_to_pptx(staged_pptx, protocol, staged_pptx)
        protocol = extract_protocol(
            staged_pptx,
            ids_from_script=ids_from_script,
            narration_order_policy=narration_order_policy,
        )
        script_authority = {
            "schema_version": "paper2video_user_script_authority.v1",
            "script_json": None,
            "resolution": "llm_regeneration",
            "model": regeneration_model,
            "changed_target_count": regeneration_report["target_count"],
            "updated_count": regeneration_report["updated_count"],
            "slide_count": protocol["slide_count"],
        }
    elif script_json is not None:
        protocol, script_authority = apply_user_script(protocol, script_json)
        write_protocol_to_pptx(staged_pptx, protocol, staged_pptx)
        protocol = extract_protocol(
            staged_pptx,
            ids_from_script=script_json,
            narration_order_policy=narration_order_policy,
        )
    else:
        script_authority = {
            "schema_version": "paper2video_user_script_authority.v1",
            "script_json": None,
            "resolution": "pptx_protocol",
            "script_sources": protocol.get("script_sources") or [],
            "slide_count": protocol["slide_count"],
            "baseline_pptx": str(baseline_pptx) if baseline_pptx else None,
            "detected_change_count": (
                int(change_report["change_count"])
                if change_report is not None
                else None
            ),
        }

    protocol_writeback = write_protocol_to_pptx(
        staged_pptx,
        protocol,
        staged_pptx,
    )
    click_group_report = normalize_animation_click_groups(
        staged_pptx,
        staged_pptx,
        policy=click_group_policy,
    )
    protocol = extract_protocol(
        staged_pptx,
        ids_from_script=ids_from_script or script_json,
        narration_order_policy=narration_order_policy,
    )
    return {
        "authority_report": authority_report,
        "protocol": protocol,
        "change_report": change_report,
        "regeneration_report": regeneration_report,
        "script_authority": script_authority,
        "protocol_writeback": protocol_writeback,
        "click_group_report": click_group_report,
    }


def _copy_audio_bundle(source: Path, destination: Path, section_ids: list[str]) -> None:
    source = source.resolve()
    destination.mkdir(parents=True, exist_ok=True)
    try:
        source_file_map = load_audio_manifest_file_map(source)
    except ValueError as exc:
        raise ProtocolError(str(exc)) from exc
    for section_id in section_ids:
        try:
            source_mp3 = resolve_audio_path(
                source,
                section_id,
                file_map=source_file_map,
            )
        except ValueError as exc:
            raise ProtocolError(str(exc)) from exc
        if not source_mp3.is_file():
            raise ProtocolError(f"prebuilt audio is missing {source_mp3}")
        destination_mp3 = destination / source_mp3.name
        if source_mp3.resolve() != destination_mp3.resolve():
            shutil.copy2(source_mp3, destination_mp3)
    timings = source / "word_timings.json"
    if not timings.is_file():
        raise ProtocolError(f"prebuilt audio is missing {timings}")
    if timings.resolve() != (destination / timings.name).resolve():
        shutil.copy2(timings, destination / timings.name)
    manifest = source / "manifest.json"
    if manifest.is_file() and manifest.resolve() != (destination / manifest.name).resolve():
        shutil.copy2(manifest, destination / manifest.name)


def _prune_orphan_audio(audio_dir: Path, section_ids: set[str]) -> None:
    try:
        audio_file_map = load_audio_manifest_file_map(audio_dir)
    except ValueError as exc:
        raise ProtocolError(str(exc)) from exc
    retained_filenames = {
        audio_file_map[section_id]
        for section_id in section_ids
        if section_id in audio_file_map
    }
    retained_filenames.update(
        filename
        for section_id in section_ids
        if section_id not in audio_file_map
        for filename in [legacy_audio_filename(section_id)]
        if filename is not None
    )
    for mp3 in audio_dir.glob("*.mp3"):
        if mp3.name not in retained_filenames:
            mp3.unlink()


def _attach_block_tts_parts(
    script: dict[str, object],
    protocol: dict[str, object],
) -> dict[str, object]:
    """Expose protocol blocks as independently cacheable, sequenced TTS units.

    ``pre_roll_seconds`` is deliberately not part of the TTS cache identity.  It
    describes output assembly only: the native entrance rows since the previous
    narrated target play in this silence, then this block's pristine cached
    speech begins.  Measuring between narrated target ``pane_end_seconds``
    anchors also accounts for intervening silent connector rows and Pane delays.
    """
    from .generate_edge_audio import normalize_script_text

    slides_by_section = {
        str(slide.get("section_id") or ""): slide
        for slide in protocol.get("slides") or []
    }
    for section in script.get("sections") or []:
        section_id = str(section.get("id") or "")
        slide = slides_by_section.get(section_id)
        if slide is None:
            continue
        parts: list[dict[str, object]] = []
        previous_narrated_pane_end = 0.0
        maximum_pane_end = 0.0
        slide_blocks = list(slide.get("blocks") or [])
        simultaneous_group_ends: dict[int, float] = {}
        for block in slide_blocks:
            for effect in block.get("effects") or []:
                if (
                    str(effect.get("authority") or "") != "animation_pane"
                    or str(effect.get("kind") or "") != "entrance"
                ):
                    continue
                group = int(
                    effect.get("simultaneous_group")
                    or effect.get("native_order")
                    or 0
                )
                simultaneous_group_ends[group] = max(
                    simultaneous_group_ends.get(group, 0.0),
                    float(effect.get("pane_end_seconds") or 0.0),
                )

        for index, block in enumerate(slide_blocks):
            transcript = normalize_script_text(str(block.get("transcript") or ""))
            native_entrances = [
                effect
                for effect in block.get("effects") or []
                if str(effect.get("authority") or "") == "animation_pane"
                and str(effect.get("kind") or "") == "entrance"
            ]
            pane_ends = [
                simultaneous_group_ends[
                    int(
                        effect.get("simultaneous_group")
                        or effect.get("native_order")
                        or 0
                    )
                ]
                for effect in native_entrances
            ]
            maximum_pane_end = max([maximum_pane_end, *pane_ends])
            if not transcript:
                continue

            current_pane_end = max(pane_ends) if pane_ends else None
            if current_pane_end is None:
                pre_roll_seconds = 0.0
            else:
                pre_roll_seconds = max(
                    0.0,
                    current_pane_end - previous_narrated_pane_end,
                )
                previous_narrated_pane_end = max(
                    previous_narrated_pane_end,
                    current_pane_end,
                )
            parts.append(
                {
                    "id": str(block.get("handle") or f"block-{index + 1}"),
                    "text": transcript,
                    "pre_roll_seconds": round(pre_roll_seconds, 3),
                    "pane_anchor_start_seconds": round(
                        previous_narrated_pane_end - pre_roll_seconds,
                        3,
                    ),
                    "pane_anchor_end_seconds": round(
                        previous_narrated_pane_end,
                        3,
                    ),
                }
            )
        section_text = normalize_script_text(str(section.get("text") or ""))
        reconstructed = normalize_script_text(" ".join(part["text"] for part in parts))
        if reconstructed != section_text:
            raise ProtocolError(
                f"section {section_id} protocol blocks do not reconstruct its narration"
            )
        if parts:
            section["tts_parts"] = parts
            section["post_roll_seconds"] = round(
                max(0.0, maximum_pane_end - previous_narrated_pane_end),
                3,
            )
    return script


def _pad_audio_for_sequence(
    manifest: dict[str, object],
    audio_dir: Path,
) -> dict[str, object]:
    """Keep the rendered segment alive through the resolved animation schedule."""
    from .generate_edge_audio import (
        ensure_minimum_audio_duration,
        probe_audio_duration,
    )

    entries: list[dict[str, object]] = []
    try:
        audio_file_map = load_audio_manifest_file_map(audio_dir)
    except ValueError as exc:
        raise ProtocolError(str(exc)) from exc
    for slide in manifest.get("slides") or []:
        section_id = str(slide["id"])
        try:
            audio_path = resolve_audio_path(
                audio_dir,
                section_id,
                file_map=audio_file_map,
            )
        except ValueError as exc:
            raise ProtocolError(str(exc)) from exc
        minimum = round(float(slide.get("schedule_end") or 0.0) + 0.05, 3)
        before = probe_audio_duration(audio_path)
        padded = ensure_minimum_audio_duration(audio_path, minimum)
        after = probe_audio_duration(audio_path)
        if after is None or after + 0.02 < minimum:
            raise ProtocolError(
                f"audio {audio_path} ends at {after!r}s but the resolved animation "
                f"sequence requires at least {minimum:.3f}s"
            )
        entries.append(
            {
                "id": section_id,
                "file": audio_path.name,
                "schedule_end": slide.get("schedule_end"),
                "minimum_audio_seconds": minimum,
                "before_seconds": round(before, 3) if before is not None else None,
                "after_seconds": round(after, 3),
                "padded": padded,
            }
        )

    audio_manifest_path = audio_dir / "manifest.json"
    if audio_manifest_path.is_file():
        try:
            audio_manifest = json.loads(audio_manifest_path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise ProtocolError(f"could not refresh {audio_manifest_path}: {exc}") from exc
        by_id = {str(entry["id"]): entry for entry in entries}
        for item in audio_manifest:
            section_id = str(item.get("id") or "")
            if section_id not in by_id:
                continue
            item["bytes"] = (audio_dir / str(item["file"])).stat().st_size
            item["sequence_minimum_seconds"] = by_id[section_id]["minimum_audio_seconds"]
            item["sequence_padding_applied"] = by_id[section_id]["padded"]
        audio_manifest_path.write_text(
            json.dumps(audio_manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    return {
        "schema_version": "paper2video_animation_sequence_audio.v1",
        "slide_count": len(entries),
        "padded_count": sum(1 for entry in entries if entry["padded"]),
        "slides": entries,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="Edited PPTX source of truth")
    parser.add_argument("outdir", type=Path, help="Paper2Video v2 output bundle")
    parser.add_argument(
        "--ids-from-script",
        type=Path,
        default=None,
        help="Preserve existing semantic slide IDs while rebuilding narration from the PPTX protocol.",
    )
    parser.add_argument(
        "--script-json",
        type=Path,
        default=None,
        help=(
            "Use a user-edited script.json as narration authority. Section text replaces "
            "slide narration; optional per-section elements preserve precise handle timing."
        ),
    )
    parser.add_argument(
        "--baseline-pptx",
        type=Path,
        default=None,
        help="Previous editable PPTX used to identify which elements changed.",
    )
    parser.add_argument(
        "--narration-mode",
        choices=("keep", "regenerate"),
        default="keep",
        help="Keep PPTX narration, or regenerate narration only for changed elements.",
    )
    parser.add_argument(
        "--narration-order-policy",
        choices=("geometry", "author-notes"),
        default="geometry",
        help=(
            "Preserve canonical Notes block order and insert only Notes-missing "
            "targets by deterministic slide geometry (default); author-notes is "
            "a compatibility alias."
        ),
    )
    parser.add_argument(
        "--regeneration-model",
        default=DEFAULT_MODEL,
        help="OpenAI model used only with --narration-mode regenerate.",
    )
    parser.add_argument("--voice", default=None, help="Edge TTS voice")
    parser.add_argument("--rate", default="+0%", help="Edge TTS rate")
    tts_cache_group = parser.add_mutually_exclusive_group()
    tts_cache_group.add_argument(
        "--tts-cache-dir",
        type=Path,
        default=None,
        help="Persistent pristine Edge TTS cache directory.",
    )
    tts_cache_group.add_argument(
        "--no-tts-cache",
        action="store_true",
        help="Disable persistent Edge TTS reuse.",
    )
    parser.add_argument(
        "--prebuilt-audio-dir",
        type=Path,
        default=None,
        help="Offline/test mode: use matching MP3s and word_timings.json instead of calling Edge TTS.",
    )
    parser.add_argument("--resolution", choices=("720p", "1080p", "1440p", "4k"), default="1080p")
    parser.add_argument("--fps", type=int, default=DEFAULT_FPS)
    parser.add_argument("--start-pad", type=float, default=DEFAULT_START_PAD)
    parser.add_argument("--pad-tail", type=float, default=DEFAULT_PAD_TAIL)
    parser.add_argument("--visual-cues", type=Path, default=None)
    parser.add_argument(
        "--visual-cue-plan",
        type=Path,
        default=None,
        help="Optional matching visual_cue_plan.json for an externally supplied --visual-cues file.",
    )
    parser.add_argument(
        "--highlight-style",
        default="spotlight_laser",
        choices=(
            "box", "spotlight", "cursor", "box_cursor", "spotlight_cursor",
            "laser", "box_laser", "spotlight_laser",
        ),
    )
    parser.add_argument(
        "--animation-order-policy",
        choices=("auto", "animation-pane", "reading-order"),
        default="auto",
        help=(
            "Resolve narration-order versus Animation Pane conflicts. "
            "auto asks in a TTY; non-TTY and reading-order choices exit 3 before "
            "creating the output bundle."
        ),
    )
    parser.add_argument(
        "--animation-order-sequence",
        action="append",
        default=[],
        metavar="[SLIDE=]ORDER",
        help=(
            "Deterministic Identity Map permutation such as CABGDEF. Repeat "
            "as SLIDE=ORDER when more than one conflict page exists. Requires "
            "--animation-order-policy auto."
        ),
    )
    parser.add_argument(
        "--semantic-profile",
        choices=("concise", "detailed"),
        default="concise",
        help=(
            "Choose the persisted semantic description shown beside A/B/C "
            "element identities in animation-order conflict previews."
        ),
    )
    parser.add_argument(
        "--click-group-policy",
        choices=("normalize", "preserve"),
        default="normalize",
        help=(
            "Normalize every slide so each non-With-Previous animation phase "
            "has a numbered PowerPoint click badge (default), or preserve the "
            "source mainSeq topology."
        ),
    )
    parser.add_argument(
        "--animation-order-report",
        type=Path,
        default=None,
        help=(
            "Decision report path used when preflight stops before creating the output "
            "bundle. Defaults to <outdir>.animation-order-authority.json."
        ),
    )
    parser.add_argument("--no-subtitles", action="store_true")
    parser.add_argument("--keep-temp", action="store_true")
    parser.add_argument(
        "--no-qa",
        action="store_true",
        help="Debug only: skip the final strict package gate.",
    )
    args = parser.parse_args()

    source_pptx = args.pptx.resolve()
    outdir = args.outdir.resolve()
    if not source_pptx.is_file():
        sys.exit(f"[render_edited_pptx] PPTX not found: {source_pptx}")
    if outdir.exists():
        sys.exit(
            f"[render_edited_pptx] output bundle already exists; choose a fresh path: {outdir}"
        )
    try:
        validate_render_timing(
            fps=args.fps,
            start_pad=args.start_pad,
            pad_tail=args.pad_tail,
        )
    except ValueError as exc:
        sys.exit(f"[render_edited_pptx] invalid render timing: {exc}")

    external_order_report_path = (
        args.animation_order_report.resolve()
        if args.animation_order_report is not None
        else Path(str(outdir) + ".animation-order-authority.json")
    )
    if external_order_report_path == outdir or external_order_report_path.is_relative_to(
        outdir
    ):
        sys.exit(
            "[render_edited_pptx] --animation-order-report must be outside the "
            "fresh output bundle path"
        )

    audio_dir = outdir / "assets" / "audio"
    captions_dir = outdir / "assets" / "captions"
    slides_dir = outdir / "assets" / "slides"
    clips_dir = outdir / "assets" / "clips"
    meta_dir = outdir / "assets" / "meta"
    reports_dir = meta_dir / "reports"
    frames_dir = slides_dir / "frames"
    script_path = audio_dir / "script.json"
    protocol_path = reports_dir / "editable_pptx_protocol.json"
    timings_path = audio_dir / "word_timings.json"
    animation_manifest_path = meta_dir / "animation_manifest.json"
    animation_report_path = reports_dir / "animation_render_report.json"
    author_cues_path = meta_dir / "editable_pptx_visual_cues.json"
    author_cue_plan_path = meta_dir / "editable_pptx_visual_cue_plan.json"
    duration_report_path = meta_dir / "video_duration_report.json"
    timeline_path = meta_dir / "timeline.json"
    raw_path = clips_dir / "video_raw.mp4"
    raw_delivery = outdir / "video_no_subtitles.mp4"
    final_path = outdir / "video.mp4"
    srt_path = captions_dir / "video.srt"
    vtt_path = captions_dir / "video.vtt"
    qa_path = reports_dir / "video_qa_report.json"
    authority_report_path = reports_dir / "author_notes_authority.json"
    script_authority_path = reports_dir / "script_authority.json"
    changes_report_path = reports_dir / "pptx_changes.json"
    regeneration_report_path = reports_dir / "narration_regeneration.json"
    sequence_audio_report_path = reports_dir / "animation_sequence_audio.json"
    protocol_writeback_path = reports_dir / "protocol_writeback.json"
    subtitle_timing_report_path = reports_dir / "subtitle_timing_alignment.json"
    animation_order_authority_path = reports_dir / "animation_order_authority.json"
    click_group_report_path = reports_dir / "click_group_normalization.json"
    delivered_pptx = outdir / "video.pptx"
    with tempfile.TemporaryDirectory(prefix="pptx2video_order_preflight_") as temp_dir:
        staged_pptx = Path(temp_dir) / "normalized.pptx"
        try:
            prepared = _prepare_protocol_for_render(
                source_pptx,
                staged_pptx,
                ids_from_script=args.ids_from_script,
                script_json=args.script_json,
                baseline_pptx=args.baseline_pptx,
                narration_mode=args.narration_mode,
                regeneration_model=args.regeneration_model,
                narration_order_policy=args.narration_order_policy,
                click_group_policy=args.click_group_policy,
            )
            authority_report = prepared["authority_report"]
            protocol = prepared["protocol"]
            assert isinstance(authority_report, dict) and isinstance(protocol, dict)
            decision_protocol = {
                **protocol,
                "source_pptx": str(source_pptx),
                "source_sha256": file_sha256(source_pptx),
            }
            order_report = analyze_animation_order_conflicts(
                decision_protocol,
                authority_report,
            )
            order_report["normalized_source_sha256"] = protocol.get(
                "source_sha256"
            )
            selected_animation_order, confirmed_order_report = (
                select_animation_order_policy(
                    args.animation_order_policy,
                    order_report,
                    report_path=external_order_report_path,
                    semantic_profile=args.semantic_profile,
                    identity_order_inputs=args.animation_order_sequence,
                )
            )
            if selected_animation_order is None:
                return 3

            for directory in (
                audio_dir,
                captions_dir,
                slides_dir,
                clips_dir,
                reports_dir,
            ):
                directory.mkdir(parents=True, exist_ok=True)
            shutil.copy2(staged_pptx, delivered_pptx)

            protocol = extract_protocol(
                delivered_pptx,
                ids_from_script=args.ids_from_script or args.script_json,
                narration_order_policy=args.narration_order_policy,
            )
            final_decision_protocol = {
                **protocol,
                "source_pptx": str(source_pptx),
                "source_sha256": file_sha256(source_pptx),
            }
            final_order_report = analyze_animation_order_conflicts(
                final_decision_protocol,
                authority_report,
            )
            final_order_report["normalized_source_sha256"] = protocol.get(
                "source_sha256"
            )
            if final_order_report.get("conflict_set_fingerprint") != (
                confirmed_order_report.get("conflict_set_fingerprint")
            ):
                raise ProtocolError(
                    "animation-order conflict set changed after the user decision"
                )
            final_order_report.update(
                {
                    "status": confirmed_order_report["status"],
                    "selected_policy": selected_animation_order,
                    "selection_source": confirmed_order_report["selection_source"],
                    "slide_decisions": list(
                        confirmed_order_report.get("slide_decisions") or []
                    ),
                    "confirmed_conflict_fingerprint": confirmed_order_report.get(
                        "conflict_fingerprint"
                    ),
                }
            )

            if _should_apply_confirmed_animation_pane_order(
                selected_animation_order,
                final_order_report,
            ):
                protocol, pane_order_application = (
                    apply_confirmed_animation_pane_order(
                        protocol,
                        slide_indices=_animation_pane_slide_indices(
                            final_order_report
                        ),
                    )
                )
                protocol_writeback = write_protocol_to_pptx(
                    delivered_pptx,
                    protocol,
                    delivered_pptx,
                )
                delivered_sha256 = file_sha256(delivered_pptx)
                protocol["source_pptx"] = str(delivered_pptx)
                protocol["source_sha256"] = delivered_sha256
                protocol_writeback["render_order_application"] = (
                    pane_order_application
                )
                final_order_report["render_order_application"] = (
                    pane_order_application
                )
            else:
                delivered_sha256 = file_sha256(delivered_pptx)
                protocol_writeback = prepared["protocol_writeback"]
                assert isinstance(protocol_writeback, dict)
                protocol_writeback.update(
                    {
                        "source_pptx": str(delivered_pptx),
                        "output_pptx": str(delivered_pptx),
                        "output_sha256": delivered_sha256,
                    }
                )
            authority_report.update(
                {
                    "output_pptx": str(delivered_pptx),
                    "output_sha256": delivered_sha256,
                }
            )
            final_order_report.update(
                {
                    "normalized_source_sha256": delivered_sha256,
                }
            )

            script_authority = prepared["script_authority"]
            click_group_report = prepared["click_group_report"]
            change_report = prepared["change_report"]
            regeneration_report = prepared["regeneration_report"]
            script = _attach_block_tts_parts(
                script_from_protocol(protocol, voice=args.voice),
                protocol,
            )
            write_json(authority_report_path, authority_report)
            assert isinstance(click_group_report, dict)
            click_group_report.update(
                {
                    "output_pptx": str(delivered_pptx),
                    "output_sha256": delivered_sha256,
                }
            )
            write_json(click_group_report_path, click_group_report)
            write_json(protocol_writeback_path, protocol_writeback)
            write_json(animation_order_authority_path, final_order_report)
            write_json(script_path, script)
            write_json(protocol_path, protocol)
            write_json(script_authority_path, script_authority)
            if isinstance(change_report, dict):
                write_json(changes_report_path, change_report)
            if isinstance(regeneration_report, dict):
                write_json(regeneration_report_path, regeneration_report)
        except (OSError, ProtocolError) as exc:
            sys.exit(
                f"[render_edited_pptx] PPTX protocol reconciliation failed: {exc}"
            )

    section_ids = [str(section["id"]) for section in script["sections"]]
    has_narration = any(str(section.get("text") or "").strip() for section in script["sections"])
    effective_no_subtitles = args.no_subtitles or not has_narration
    _prune_orphan_audio(audio_dir, set(section_ids))

    try:
        if args.prebuilt_audio_dir is not None:
            _copy_audio_bundle(args.prebuilt_audio_dir, audio_dir, section_ids)
        else:
            tts_command = _module_command(
                "generate_edge_audio",
                str(script_path),
                "--outdir",
                str(audio_dir),
                "--rate",
                args.rate,
                "--timings-out",
                str(timings_path),
            )
            if args.voice:
                tts_command.extend(["--voice", args.voice])
            if args.no_tts_cache:
                tts_command.append("--no-cache")
            elif args.tts_cache_dir is not None:
                tts_command.extend(["--cache-dir", str(args.tts_cache_dir.resolve())])
            _run(tts_command)

        manifest = build_pptx_animation_manifest(protocol, timings_path)
        write_json(animation_manifest_path, manifest)
        author_cues = build_pptx_visual_cues(protocol, timings_path)
        write_json(author_cues_path, author_cues)
        author_cue_plan = build_pptx_visual_cue_plan(author_cues)
        write_json(author_cue_plan_path, author_cue_plan)
        write_json(
            sequence_audio_report_path,
            _pad_audio_for_sequence(manifest, audio_dir),
        )
    except (OSError, ProtocolError, subprocess.CalledProcessError) as exc:
        sys.exit(f"[render_edited_pptx] audio/manifest stage failed: {exc}")

    render_command = _module_command(
        "render_video",
        str(outdir),
        "--pptx",
        str(delivered_pptx),
        "--audio-dir",
        str(audio_dir),
        "--script-json",
        str(script_path),
        "--frame-source",
        "pptx",
        "--animation-source",
        "pptx",
        "--animation-manifest",
        str(animation_manifest_path),
        "--animation-report-out",
        str(animation_report_path),
        "--duration-report-out",
        str(duration_report_path),
        "--resolution",
        args.resolution,
        "--fps",
        str(args.fps),
        "--start-pad",
        str(args.start_pad),
        "--pad-tail",
        str(args.pad_tail),
        "--frames-out",
        str(frames_dir),
        "--out",
        str(raw_path),
    )
    if int(manifest.get("effect_count") or 0) > 0:
        render_command.append("--require-animations")
    effective_visual_cues = (
        args.visual_cues.resolve()
        if args.visual_cues is not None
        else (author_cues_path if int(author_cues.get("cue_count") or 0) else None)
    )
    using_native_emphasis_cues = (
        args.visual_cues is None and effective_visual_cues == author_cues_path
    )
    effective_cue_plan = (
        args.visual_cue_plan.resolve()
        if args.visual_cue_plan is not None
        else (author_cue_plan_path if using_native_emphasis_cues else None)
    )
    if args.visual_cue_plan is not None and args.visual_cues is None:
        sys.exit("[render_edited_pptx] --visual-cue-plan requires --visual-cues")
    if effective_visual_cues is not None:
        render_command.extend(
            [
                "--attention-mode",
                "highlight",
                "--highlight-style",
                args.highlight_style,
                "--visual-cues",
                str(effective_visual_cues),
            ]
        )
    else:
        render_command.extend(["--attention-mode", "none"])
    if args.keep_temp:
        render_command.append("--keep-temp")

    try:
        _run(render_command)
        shutil.copy2(raw_path, raw_delivery)
        subtitle_command = _module_command(
            "add_subtitles",
            str(outdir),
            "--mp4",
            str(raw_path),
            "--audio-dir",
            str(audio_dir),
            "--script-json",
            str(script_path),
            "--word-timings",
            str(timings_path),
            "--require-word-timings",
            "--timing-report-out",
            str(subtitle_timing_report_path),
            "--start-pad",
            str(args.start_pad),
            "--pad-tail",
            str(args.pad_tail),
            "--srt-out",
            str(srt_path),
            "--vtt-out",
            str(vtt_path),
            "--out",
            str(final_path),
        )
        if effective_no_subtitles:
            subtitle_command.append("--no-subtitles")
        _run(subtitle_command)
        shutil.copy2(delivered_pptx, slides_dir / "slides.pptx")

        timeline_command = _module_command(
            "build_timeline",
            "--script-json",
            str(script_path),
            "--duration-report",
            str(duration_report_path),
            "--captions-vtt",
            str(vtt_path),
            "--audio-dir",
            str(audio_dir),
            "--video",
            str(raw_delivery),
            "--out",
            str(timeline_path),
        )
        if effective_visual_cues is not None:
            timeline_command.extend(["--visual-cues", str(effective_visual_cues)])
        if effective_cue_plan is not None:
            timeline_command.extend(["--visual-cue-plan", str(effective_cue_plan)])
        _run(timeline_command)
    except (OSError, subprocess.CalledProcessError) as exc:
        sys.exit(f"[render_edited_pptx] render stage failed: {exc}")

    if not args.no_qa:
        qa_command = _module_command(
            "check_video_package",
            str(outdir),
            "--pptx",
            str(outdir / "video.pptx"),
            "--script-json",
            str(script_path),
            "--audio-dir",
            str(audio_dir),
            "--frames-dir",
            str(frames_dir),
            "--mp4",
            str(final_path),
            "--raw-mp4",
            str(raw_delivery),
            "--subtitle-file",
            str(vtt_path),
            "--subtitle-timing-report",
            str(subtitle_timing_report_path),
            "--animation-manifest",
            str(animation_manifest_path),
            "--animation-report",
            str(animation_report_path),
            "--animation-order-authority",
            str(animation_order_authority_path),
            "--timeline",
            str(timeline_path),
            "--require-word-timings",
            "--require-timeline",
            "--strict",
            "--out",
            str(qa_path),
        )
        if int(manifest.get("effect_count") or 0) > 0:
            qa_command.extend(
                ["--require-animations", "--require-animation-order-authority"]
            )
        if not effective_no_subtitles:
            qa_command.extend(
                ["--require-subtitles", "--require-subtitle-word-alignment"]
            )
        if effective_visual_cues is not None:
            qa_command.extend(
                [
                    "--visual-cues",
                    str(effective_visual_cues),
                ]
            )
        if effective_cue_plan is not None:
            qa_command.extend(["--cue-plan", str(effective_cue_plan)])
        if args.visual_cues is not None and effective_cue_plan is not None:
            qa_command.extend(
                ["--strict-attention", "--require-visual-cues", "--require-cue-plan"]
            )
        else:
            # Native emphasis is optional. It is still rendered and audited
            # when present, but an editable deck is not required to spotlight
            # every narration chunk merely to pass media/protocol QA.
            qa_command.append("--allow-missing-attention")
        try:
            _run(qa_command)
        except subprocess.CalledProcessError as exc:
            sys.exit(f"[render_edited_pptx] strict QA failed with exit {exc.returncode}")

    print(
        f"[render_edited_pptx] DONE: {final_path} "
        f"({protocol['slide_count']} slides, {protocol['effect_count']} effects)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
