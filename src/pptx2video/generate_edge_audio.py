#!/usr/bin/env python3
"""
Generate MP3 narration from a pptx2video script JSON using edge-tts.

Logical section IDs remain in script/timing/manifest metadata. Physical MP3s
use portable index-and-hash filenames recorded by manifest.json.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from importlib.metadata import PackageNotFoundError, version as distribution_version
from pathlib import Path

from .audio_files import portable_audio_filename
from .ffmpeg import find_ffmpeg_pair as _shared_ffmpeg_pair

try:
    import edge_tts
    from edge_tts.exceptions import (
        NoAudioReceived,
        UnexpectedResponse,
        UnknownResponse,
        WebSocketError,
    )
except ImportError as exc:  # pragma: no cover - depends on local env
    raise SystemExit(
        "[generate_edge_audio] edge_tts is not installed in this Python env. "
        "Install the pptx2video package with its required dependencies."
    ) from exc

# Transport-level failures worth another attempt. A malformed request would
# fail identically every time, so the list stays narrow on purpose.
TRANSIENT_TTS_ERRORS: tuple[type[BaseException], ...] = (
    NoAudioReceived,
    WebSocketError,
    UnexpectedResponse,
    UnknownResponse,
    ConnectionError,
    TimeoutError,
    asyncio.TimeoutError,
    OSError,
)


DEFAULT_VOICE = "en-US-AriaNeural"

# Edge TTS is a remote service reached over a WebSocket, once per cache miss.
# A session that closes without delivering audio raises NoAudioReceived, which
# is transient far more often than it is a problem with the text: the same
# input usually succeeds moments later. Without a retry, one such session ends
# the whole run and discards every clip synthesized before it, so a deck fails
# after minutes of work for a reason that has nothing to do with the deck.
TTS_MAX_ATTEMPTS = max(1, int(os.environ.get("PPTX2VIDEO_TTS_ATTEMPTS", "4")))
TTS_RETRY_BASE_SECONDS = 1.5
TTS_RETRY_MAX_SECONDS = 20.0
TIMINGS_SCHEMA_VERSION = "paper2video_edge_word_boundaries.v1"
TTS_CACHE_SCHEMA_VERSION = "paper2video_tts_cache.v1"
TTS_PROVIDER = "edge-tts"


def _edge_tts_version() -> str:
    try:
        return distribution_version("edge-tts")
    except PackageNotFoundError:  # pragma: no cover - package metadata is normally present
        return str(getattr(edge_tts, "__version__", "unknown"))


TTS_PROVIDER_VERSION = _edge_tts_version()


def find_ffmpeg_pair(
    *, required: bool = True
) -> tuple[str | None, str | None]:
    """Return the shared doctor/runtime FFmpeg selection."""
    return _shared_ffmpeg_pair(
        required=required,
        component="generate_edge_audio",
    )


def normalize_script_text(text: str) -> str:
    """Normalize narration for stable cache identity and synthesis."""
    return " ".join(unicodedata.normalize("NFKC", str(text)).split())


def script_sha256(text: str) -> str:
    return hashlib.sha256(normalize_script_text(text).encode("utf-8")).hexdigest()


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_tts_cache_identity(
    text: str,
    *,
    voice: str,
    rate: str,
    pitch: str,
    provider: str = TTS_PROVIDER,
    provider_version: str = TTS_PROVIDER_VERSION,
) -> dict[str, str]:
    identity = {
        "schema_version": TTS_CACHE_SCHEMA_VERSION,
        "script_sha256": script_sha256(text),
        "voice": str(voice),
        "rate": str(rate),
        "pitch": str(pitch),
        "provider": str(provider),
        "provider_version": str(provider_version),
    }
    serialized = json.dumps(
        identity,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return {
        **identity,
        "cache_key": hashlib.sha256(serialized.encode("utf-8")).hexdigest(),
    }


def default_tts_cache_dir() -> Path:
    override = os.environ.get("PPTX2VIDEO_TTS_CACHE_DIR")
    if override:
        return Path(override).expanduser()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Caches"
    else:
        base = Path(os.environ.get("XDG_CACHE_HOME") or (Path.home() / ".cache"))
    return base / "pptx2video" / "tts"


def _cache_entry_dir(cache_dir: Path, cache_key: str) -> Path:
    return cache_dir / cache_key[:2] / cache_key


def _atomic_copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=destination.name + ".",
        suffix=".tmp",
        dir=destination.parent,
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        shutil.copyfile(source, temporary_path)
        temporary_path.replace(destination)
    finally:
        temporary_path.unlink(missing_ok=True)


def _atomic_write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=path.name + ".",
        suffix=".tmp",
        dir=path.parent,
        mode="w",
        encoding="utf-8",
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
        json.dump(payload, temporary, ensure_ascii=False, indent=2)
        temporary.write("\n")
    try:
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def restore_cached_tts(
    cache_dir: Path,
    identity: dict[str, str],
    out_path: Path,
    *,
    require_timings: bool,
) -> tuple[bool, list[dict]]:
    entry_dir = _cache_entry_dir(cache_dir, identity["cache_key"])
    metadata_path = entry_dir / "metadata.json"
    audio_path = entry_dir / "audio.mp3"
    timings_path = entry_dir / "word_boundaries.json"
    if not metadata_path.is_file() or not audio_path.is_file():
        return False, []
    if require_timings and not timings_path.is_file():
        return False, []
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        if any(metadata.get(key) != value for key, value in identity.items()):
            return False, []
        audio_sha256 = _file_sha256(audio_path)
        if metadata.get("audio_sha256") != audio_sha256:
            return False, []
        words: list[dict] = []
        if require_timings:
            timing_payload = json.loads(timings_path.read_text(encoding="utf-8"))
            if timing_payload.get("cache_key") != identity["cache_key"]:
                return False, []
            if timing_payload.get("audio_sha256") != audio_sha256:
                return False, []
            raw_words = timing_payload.get("words")
            if not isinstance(raw_words, list):
                return False, []
            words = raw_words
    except (OSError, ValueError):
        return False, []
    _atomic_copy(audio_path, out_path)
    return True, words


def store_cached_tts(
    cache_dir: Path,
    identity: dict[str, str],
    audio_path: Path,
    words: list[dict] | None,
) -> None:
    entry_dir = _cache_entry_dir(cache_dir, identity["cache_key"])
    _atomic_copy(audio_path, entry_dir / "audio.mp3")
    audio_sha256 = _file_sha256(audio_path)
    if words is not None:
        _atomic_write_json(
            entry_dir / "word_boundaries.json",
            {
                "schema_version": TIMINGS_SCHEMA_VERSION,
                "cache_key": identity["cache_key"],
                "audio_sha256": audio_sha256,
                "words": words,
            },
        )
    _atomic_write_json(
        entry_dir / "metadata.json",
        {**identity, "audio_sha256": audio_sha256},
    )


def load_rate_plan(path: Path, *, allow_unsafe: bool) -> str:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"[generate_edge_audio] rate plan not found: {path}")
    except json.JSONDecodeError as exc:
        sys.exit(f"[generate_edge_audio] invalid rate plan {path}: {exc}")
    if payload.get("schema_version") != "paper2video_tts_rate_plan.v1":
        sys.exit(f"[generate_edge_audio] unsupported rate plan schema: {payload.get('schema_version')}")
    status = str(payload.get("status") or "")
    safe = bool(payload.get("safe"))
    if not safe and not allow_unsafe:
        sys.exit(
            "[generate_edge_audio] rate plan is not safe for automatic TTS regeneration "
            f"(status={status}). Rewrite the narration script first, or pass "
            "--allow-unsafe-rate-plan only for an explicit experiment."
        )
    if status == "needs_script_rewrite" and not allow_unsafe:
        sys.exit("[generate_edge_audio] rate plan requires script rewrite; refusing to hide it with TTS rate.")
    rate = str(payload.get("recommended_edge_rate") or "+0%")
    if not rate.endswith("%") or not (rate.startswith("+") or rate.startswith("-")):
        sys.exit(f"[generate_edge_audio] invalid recommended_edge_rate in {path}: {rate!r}")
    return rate


async def synthesize_section(text: str, *, voice: str, rate: str, pitch: str, out_path: Path) -> None:
    communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate, pitch=pitch)
    await communicate.save(str(out_path))


def edge_ticks_to_seconds(raw: object) -> float:
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return 0.0
    # edge-tts WordBoundary offsets are 100ns ticks.
    return value / 10_000_000.0


async def synthesize_section_with_timings(
    text: str,
    *,
    voice: str,
    rate: str,
    pitch: str,
    out_path: Path,
) -> list[dict]:
    communicate = edge_tts.Communicate(
        text=text,
        voice=voice,
        rate=rate,
        pitch=pitch,
        boundary="WordBoundary",
    )
    words: list[dict] = []
    with out_path.open("wb") as fh:
        async for chunk in communicate.stream():
            kind = chunk.get("type")
            if kind == "audio":
                data = chunk.get("data")
                if data:
                    fh.write(data)
            elif kind == "WordBoundary":
                start = edge_ticks_to_seconds(chunk.get("offset"))
                duration = edge_ticks_to_seconds(chunk.get("duration"))
                words.append({
                    "text": str(chunk.get("text") or ""),
                    "start": round(start, 3),
                    "end": round(start + max(duration, 0.0), 3),
                    "duration": round(max(duration, 0.0), 3),
                })
    return words


def probe_audio_duration(path: Path) -> float | None:
    """Return decoded duration through the same binaries checked by doctor."""
    ffmpeg, ffprobe = find_ffmpeg_pair(required=False)
    if ffmpeg is None or ffprobe is None:
        return None
    if ffprobe != ffmpeg:
        probe = subprocess.run(
            [
                ffprobe,
                "-v",
                "error",
                "-show_entries",
                "format=duration",
                "-of",
                "default=noprint_wrappers=1:nokey=1",
                str(path),
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        try:
            if probe.returncode == 0:
                return float(probe.stdout.strip())
        except ValueError:
            pass
    probe = subprocess.run(
        [ffmpeg, "-i", str(path)],
        capture_output=True,
        text=True,
        check=False,
    )
    match = re.search(r"Duration:\s+(\d+):(\d+):([\d.]+)", probe.stderr)
    if match is None:
        return None
    return (
        int(match.group(1)) * 3600
        + int(match.group(2)) * 60
        + float(match.group(3))
    )


def ensure_minimum_audio_duration(path: Path, minimum_seconds: float) -> bool:
    """Pad a spoken clip with silence when native animation timing runs longer."""
    if minimum_seconds <= 0:
        return False
    ffmpeg, _ffprobe = find_ffmpeg_pair(required=False)
    current = probe_audio_duration(path)
    if ffmpeg is None or current is None:
        return False
    if current + 0.02 >= minimum_seconds:
        return False
    with tempfile.NamedTemporaryFile(
        prefix=path.stem + ".padded.",
        suffix=".mp3",
        dir=path.parent,
        delete=False,
    ) as temporary:
        padded = Path(temporary.name)
    try:
        subprocess.run(
            [
                ffmpeg,
                "-y",
                "-v",
                "error",
                "-i",
                str(path),
                "-af",
                "apad",
                "-t",
                str(minimum_seconds),
                "-c:a",
                "libmp3lame",
                str(padded),
            ],
            check=True,
        )
        padded.replace(path)
    finally:
        padded.unlink(missing_ok=True)
    return True


async def _synthesize_with_retry(
    text: str,
    *,
    voice: str,
    rate: str,
    pitch: str,
    out_path: Path,
    collect_timings: bool,
    label: str,
) -> list[dict]:
    """Synthesize one unit, retrying transient service failures.

    Returns word boundaries when requested. Raises the final exception with the
    unit named, so a persistent failure identifies which text to look at rather
    than only reporting that some request returned no audio.
    """
    last_error: BaseException | None = None
    for attempt in range(1, TTS_MAX_ATTEMPTS + 1):
        try:
            if collect_timings:
                return await synthesize_section_with_timings(
                    text, voice=voice, rate=rate, pitch=pitch, out_path=out_path
                )
            await synthesize_section(
                text, voice=voice, rate=rate, pitch=pitch, out_path=out_path
            )
            return []
        except TRANSIENT_TTS_ERRORS as exc:
            last_error = exc
            if attempt >= TTS_MAX_ATTEMPTS:
                break
            # Exponential backoff. Edge TTS shrugs off a brief pause far more
            # reliably than an immediate retry against the same endpoint.
            delay = min(
                TTS_RETRY_MAX_SECONDS,
                TTS_RETRY_BASE_SECONDS * (2 ** (attempt - 1)),
            )
            print(
                f"[edge-tts] {label}: attempt {attempt}/{TTS_MAX_ATTEMPTS} failed "
                f"({type(exc).__name__}: {exc}); retrying in {delay:.1f}s",
                file=sys.stderr,
                flush=True,
            )
            await asyncio.sleep(delay)

    raise RuntimeError(
        f"Edge TTS failed for {label} after {TTS_MAX_ATTEMPTS} attempt(s): "
        f"{type(last_error).__name__}: {last_error}. "
        f"Text ({len(text)} chars): {text[:120]!r}"
    ) from last_error


async def _materialize_tts_unit(
    text: str,
    *,
    voice: str,
    rate: str,
    pitch: str,
    out_path: Path,
    collect_timings: bool,
    cache_dir: Path | None,
    provider_version: str,
    label: str = "",
) -> tuple[list[dict], dict[str, str]]:
    normalized_text = normalize_script_text(text)
    identity = build_tts_cache_identity(
        normalized_text,
        voice=voice,
        rate=rate,
        pitch=pitch,
        provider_version=provider_version,
    )
    if cache_dir is not None:
        restored, words = restore_cached_tts(
            cache_dir,
            identity,
            out_path,
            require_timings=collect_timings,
        )
        if restored:
            print(f"[edge-tts] cache hit {identity['cache_key'][:12]} -> {out_path}")
            return words, {**identity, "cache_status": "hit"}

    words = await _synthesize_with_retry(
        normalized_text,
        voice=voice,
        rate=rate,
        pitch=pitch,
        out_path=out_path,
        collect_timings=collect_timings,
        label=label or identity["cache_key"][:12],
    )
    if cache_dir is not None:
        store_cached_tts(
            cache_dir,
            identity,
            out_path,
            words if collect_timings else None,
        )
        cache_status = "miss"
    else:
        cache_status = "disabled"
    return words, {**identity, "cache_status": cache_status}


def _nonnegative_seconds(raw: object, *, label: str) -> float:
    try:
        value = float(raw or 0.0)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a finite non-negative number") from exc
    if not math.isfinite(value) or value < 0:
        raise ValueError(f"{label} must be a finite non-negative number")
    return round(value, 3)


def _section_tts_parts(section: dict, section_id: str, text: str) -> list[dict[str, object]]:
    raw_parts = section.get("tts_parts")
    if raw_parts is None:
        return [
            {
                "id": section_id,
                "text": normalize_script_text(text),
                "pre_roll_seconds": 0.0,
            }
        ] if text else []
    if not isinstance(raw_parts, list):
        raise ValueError(f"section {section_id} tts_parts must be an array")
    parts: list[dict[str, object]] = []
    for index, raw_part in enumerate(raw_parts):
        if not isinstance(raw_part, dict):
            raise ValueError(f"section {section_id} tts_parts[{index}] must be an object")
        part_text = normalize_script_text(str(raw_part.get("text") or ""))
        if not part_text:
            continue
        part: dict[str, object] = {
            "id": str(raw_part.get("id") or f"part-{index + 1}"),
            "text": part_text,
            "pre_roll_seconds": _nonnegative_seconds(
                raw_part.get("pre_roll_seconds"),
                label=f"section {section_id} tts_parts[{index}].pre_roll_seconds",
            ),
        }
        for key in ("pane_anchor_start_seconds", "pane_anchor_end_seconds"):
            if raw_part.get(key) is not None:
                part[key] = _nonnegative_seconds(
                    raw_part.get(key),
                    label=f"section {section_id} tts_parts[{index}].{key}",
                )
        parts.append(part)
    if normalize_script_text(" ".join(part["text"] for part in parts)) != normalize_script_text(text):
        raise ValueError(
            f"section {section_id} tts_parts do not reconstruct the section text"
        )
    return parts


def _concatenate_audio_files(part_paths: list[Path], out_path: Path) -> None:
    if not part_paths:
        raise ValueError("cannot concatenate an empty audio part list")
    if len(part_paths) == 1:
        shutil.copy2(part_paths[0], out_path)
        return
    ffmpeg, _ffprobe = find_ffmpeg_pair(required=False)
    if ffmpeg is None:
        raise ValueError("ffmpeg is required to combine block-level TTS audio")
    command = [ffmpeg, "-y", "-v", "error"]
    for part_path in part_paths:
        command.extend(["-i", str(part_path)])
    inputs = "".join(f"[{index}:a]" for index in range(len(part_paths)))
    command.extend(
        [
            "-filter_complex",
            f"{inputs}concat=n={len(part_paths)}:v=0:a=1[out]",
            "-map",
            "[out]",
            "-c:a",
            "libmp3lame",
            str(out_path),
        ]
    )
    subprocess.run(command, check=True)


def _prefix_audio_with_silence(
    source_path: Path,
    out_path: Path,
    silence_seconds: float,
) -> None:
    """Create an output-only delayed copy without mutating pristine cache audio."""
    if silence_seconds <= 0:
        shutil.copy2(source_path, out_path)
        return
    ffmpeg, _ffprobe = find_ffmpeg_pair(required=False)
    if ffmpeg is None:
        raise ValueError("ffmpeg is required to add per-block animation pre-roll")
    delay_ms = int(round(silence_seconds * 1000.0))
    subprocess.run(
        [
            ffmpeg,
            "-y",
            "-v",
            "error",
            "-i",
            str(source_path),
            "-af",
            f"adelay=delays={delay_ms}:all=1",
            "-c:a",
            "libmp3lame",
            str(out_path),
        ],
        check=True,
    )


def _append_silence(path: Path, silence_seconds: float) -> None:
    """Append output-only silence to a completed section clip."""
    if silence_seconds <= 0:
        return
    current = probe_audio_duration(path)
    if current is None:
        raise ValueError("ffprobe is required to add section animation post-roll")
    minimum = current + silence_seconds
    if not ensure_minimum_audio_duration(path, minimum):
        after = probe_audio_duration(path)
        if after is None or after + 0.02 < minimum:
            raise ValueError(
                f"could not append {silence_seconds:.3f}s animation post-roll to {path}"
            )


def _merge_part_word_boundaries(
    part_paths: list[Path],
    part_words: list[list[dict]],
    parts: list[dict[str, object]],
) -> tuple[list[dict], list[dict[str, object]]]:
    merged: list[dict] = []
    timing_parts: list[dict[str, object]] = []
    offset = 0.0
    for part_path, words, part in zip(part_paths, part_words, parts, strict=True):
        pre_roll = float(part.get("pre_roll_seconds") or 0.0)
        first_word_index = len(merged)
        for word in words:
            shifted = dict(word)
            shifted["start"] = round(
                float(word.get("start") or 0.0) + offset + pre_roll,
                3,
            )
            shifted["end"] = round(
                float(word.get("end") or 0.0) + offset + pre_roll,
                3,
            )
            merged.append(shifted)
        duration = probe_audio_duration(part_path)
        if duration is None:
            raise ValueError(
                "ffprobe is required to align word boundaries across block-level TTS audio"
            )
        timing_part: dict[str, object] = {
            "id": str(part.get("id") or ""),
            "pre_roll_seconds": round(pre_roll, 3),
            "pre_roll_start": round(offset, 3),
            "pre_roll_end": round(offset + pre_roll, 3),
            "audio_start": round(offset, 3),
            "tts_audio_start": round(offset + pre_roll, 3),
            "narration_start": (
                round(float(merged[first_word_index]["start"]), 3)
                if len(merged) > first_word_index
                else None
            ),
            "narration_end": (
                round(float(merged[-1]["end"]), 3)
                if len(merged) > first_word_index
                else None
            ),
            "audio_end": round(offset + duration, 3),
            "word_start": first_word_index,
            "word_end": len(merged) - 1,
        }
        for key in ("pane_anchor_start_seconds", "pane_anchor_end_seconds"):
            if part.get(key) is not None:
                timing_part[key] = part[key]
        timing_parts.append(timing_part)
        offset += duration
    return merged, timing_parts


async def synthesize_all(
    sections: list[dict],
    *,
    voice: str,
    rate: str,
    pitch: str,
    outdir: Path,
    collect_timings: bool,
    cache_dir: Path | None = None,
    provider_version: str = TTS_PROVIDER_VERSION,
) -> tuple[list[dict], list[dict]]:
    manifest = []
    timing_sections = []
    section_ids: list[str] = []
    seen_section_ids: set[str] = set()
    for section_index, sec in enumerate(sections, start=1):
        if not isinstance(sec, dict):
            raise ValueError(f"script section {section_index} must be an object")
        sid = str(sec.get("id") or "")
        if not sid.strip():
            raise ValueError("every script section must have an id")
        if sid in seen_section_ids:
            raise ValueError(f"duplicate script section id: {sid!r}")
        seen_section_ids.add(sid)
        section_ids.append(sid)

    for section_index, (sec, sid) in enumerate(zip(sections, section_ids, strict=True), start=1):
        text = str(sec.get("text") or "").strip()
        out_path = outdir / portable_audio_filename(sid, section_index)
        words: list[dict] = []
        timing_parts: list[dict[str, object]] = []
        cache_units: list[dict[str, str]] = []
        provider = TTS_PROVIDER
        post_roll_seconds = _nonnegative_seconds(
            sec.get("post_roll_seconds"),
            label=f"section {sid} post_roll_seconds",
        )
        if text:
            parts = _section_tts_parts(sec, sid, text)
            print(
                f"[edge-tts] {sid} ({len(text)} chars, {len(parts)} cache unit(s), "
                f"voice={voice}, rate={rate}) "
                f"-> {out_path}"
            )
            with tempfile.TemporaryDirectory(
                prefix=f".section-{section_index:04d}.tts-parts.",
                dir=outdir,
            ) as temp:
                parts_dir = Path(temp)
                part_paths: list[Path] = []
                part_words: list[list[dict]] = []
                for index, part in enumerate(parts):
                    pristine_part_path = parts_dir / f"part-{index + 1:04d}.pristine.mp3"
                    unit_words, cache_unit = await _materialize_tts_unit(
                        str(part["text"]),
                        voice=voice,
                        rate=rate,
                        pitch=pitch,
                        out_path=pristine_part_path,
                        collect_timings=collect_timings,
                        cache_dir=cache_dir,
                        provider_version=provider_version,
                        label=f"section {sid} part {index + 1}/{len(parts)}",
                    )
                    pre_roll_seconds = float(part.get("pre_roll_seconds") or 0.0)
                    sequenced_part_path = parts_dir / f"part-{index + 1:04d}.sequenced.mp3"
                    _prefix_audio_with_silence(
                        pristine_part_path,
                        sequenced_part_path,
                        pre_roll_seconds,
                    )
                    cache_units.append(
                        {
                            **cache_unit,
                            "id": str(part["id"]),
                            "output_pre_roll_seconds": round(pre_roll_seconds, 3),
                        }
                    )
                    part_paths.append(sequenced_part_path)
                    part_words.append(unit_words)
                _concatenate_audio_files(part_paths, out_path)
                if collect_timings:
                    words, timing_parts = _merge_part_word_boundaries(
                        part_paths,
                        part_words,
                        parts,
                    )
                _append_silence(out_path, post_roll_seconds)
        else:
            duration = max(1.0, float(sec.get("duration_seconds") or 1.0))
            ffmpeg, _ffprobe = find_ffmpeg_pair(required=False)
            if ffmpeg is None:
                raise ValueError(
                    f"section {sid} is silent and ffmpeg is required to create its audio track"
                )
            print(f"[edge-tts] {sid} (silent, {duration:.3f}s) -> {out_path}")
            subprocess.run(
                [
                    ffmpeg,
                    "-y",
                    "-v",
                    "error",
                    "-f",
                    "lavfi",
                    "-i",
                    "anullsrc=r=44100:cl=stereo",
                    "-t",
                    str(duration),
                    "-c:a",
                    "libmp3lame",
                    str(out_path),
                ],
                check=True,
            )
            provider = "generated-silence"
        manifest.append({
            "id": sid,
            "heading": sec.get("heading", sid),
            "file": out_path.name,
            "bytes": out_path.stat().st_size,
            "provider": provider,
            "provider_version": provider_version if provider == TTS_PROVIDER else None,
            "voice": voice,
            "rate": rate,
            "pitch": pitch,
            "word_boundaries": len(words),
            "script_sha256": script_sha256(text),
            "cache_status": (
                "not_applicable"
                if not cache_units
                else "hit"
                if all(unit["cache_status"] == "hit" for unit in cache_units)
                else "miss"
                if all(unit["cache_status"] == "miss" for unit in cache_units)
                else "partial_hit"
                if any(unit["cache_status"] == "hit" for unit in cache_units)
                else "disabled"
            ),
            "cache_units": cache_units,
            "pre_roll_seconds": round(
                sum(float(part.get("pre_roll_seconds") or 0.0) for part in parts)
                if text
                else 0.0,
                3,
            ),
            "post_roll_seconds": post_roll_seconds,
            "sequencing_applied": bool(
                post_roll_seconds
                or (
                    text
                    and any(float(part.get("pre_roll_seconds") or 0.0) for part in parts)
                )
            ),
        })
        ensure_minimum_audio_duration(
            out_path,
            max(0.0, float(sec.get("duration_seconds") or 0.0)),
        )
        manifest[-1]["bytes"] = out_path.stat().st_size
        if collect_timings:
            timing_sections.append({
                "id": sid,
                "heading": sec.get("heading", sid),
                "file": out_path.name,
                "words": words,
                "tts_parts": timing_parts,
                "post_roll_seconds": post_roll_seconds,
            })
    return manifest, timing_sections


def main() -> int:
    ap = argparse.ArgumentParser(description="Generate pptx2video narration audio with edge-tts.")
    ap.add_argument("script", help="Path to script JSON")
    ap.add_argument(
        "--outdir",
        required=True,
        help="Directory for portable MP3 files and their logical-ID manifest",
    )
    ap.add_argument("--voice", default=None,
                    help=f"Edge voice name (default: script.edge_voice or {DEFAULT_VOICE})")
    ap.add_argument("--rate", default=None, help="Edge rate adjustment, e.g. +0%%, -8%%, +10%%")
    ap.add_argument("--rate-plan", default=None,
                    help="Optional plan_tts_rate.py JSON. Uses recommended_edge_rate and refuses unsafe plans.")
    ap.add_argument("--allow-unsafe-rate-plan", action="store_true",
                    help="Allow a rate plan whose status says the script should be rewritten. Experimental only.")
    ap.add_argument("--pitch", default="+0Hz", help="Edge pitch adjustment, e.g. +0Hz")
    ap.add_argument("--timings-out", default=None,
                    help="Optional JSON path for Edge WordBoundary timings used by visual cue alignment.")
    cache_group = ap.add_mutually_exclusive_group()
    cache_group.add_argument(
        "--cache-dir",
        default=None,
        help=(
            "Persistent pristine TTS cache. Defaults to PPTX2VIDEO_TTS_CACHE_DIR or "
            "the platform user cache directory."
        ),
    )
    cache_group.add_argument(
        "--no-cache",
        action="store_true",
        help="Disable persistent TTS reuse.",
    )
    args = ap.parse_args()

    script_path = Path(args.script).resolve()
    payload = json.loads(script_path.read_text(encoding="utf-8"))
    sections = payload.get("sections") or []
    if not isinstance(sections, list) or not sections:
        sys.exit("[generate_edge_audio] script JSON has no sections array")

    voice = args.voice or payload.get("edge_voice") or DEFAULT_VOICE
    if args.rate_plan:
        rate = load_rate_plan(Path(args.rate_plan).resolve(), allow_unsafe=args.allow_unsafe_rate_plan)
        if args.rate and args.rate != rate:
            print(f"[generate_edge_audio] --rate-plan overrides --rate {args.rate} -> {rate}")
    else:
        rate = args.rate or "+0%"
    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    cache_dir = None
    if not args.no_cache:
        cache_dir = (
            Path(args.cache_dir).expanduser().resolve()
            if args.cache_dir
            else default_tts_cache_dir().resolve()
        )
        cache_dir.mkdir(parents=True, exist_ok=True)

    try:
        manifest, timing_sections = asyncio.run(
            synthesize_all(
                sections,
                voice=voice,
                rate=rate,
                pitch=args.pitch,
                outdir=outdir,
                collect_timings=args.timings_out is not None,
                cache_dir=cache_dir,
            )
        )
    except Exception as exc:
        sys.exit(f"[generate_edge_audio] {exc}")

    (outdir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if args.timings_out:
        timings_path = Path(args.timings_out).resolve()
        timings_path.parent.mkdir(parents=True, exist_ok=True)
        timings_payload = {
            "schema_version": TIMINGS_SCHEMA_VERSION,
            "provider": "edge-tts",
            "provider_version": TTS_PROVIDER_VERSION,
            "voice": voice,
            "rate": rate,
            "pitch": args.pitch,
            "sections": timing_sections,
        }
        timings_path.write_text(json.dumps(timings_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"[edge-tts] wrote word-boundary timings to {timings_path}")
    print(f"\n[edge-tts] wrote {len(manifest)} clips to {outdir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
