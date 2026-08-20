#!/usr/bin/env python3
"""Portable physical audio filenames and manifest-based logical ID lookup."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path


_WINDOWS_RESERVED_STEMS = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}
_WINDOWS_FORBIDDEN_FILENAME_CHARACTERS = frozenset('<>:"/\\|?*')


def portable_audio_filename(section_id: str, section_index: int) -> str:
    """Return a deterministic ASCII filename without exposing ``section_id``.

    The one-based section index keeps files human-orderable. The hash keeps the
    mapping stable and auditable without relying on a sanitized external ID.
    """
    if section_index < 1:
        raise ValueError("section_index must be one-based")
    digest = hashlib.sha256(str(section_id).encode("utf-8")).hexdigest()[:12]
    return f"section-{section_index:04d}-{digest}.mp3"


def is_portable_audio_filename(filename: str) -> bool:
    """Return whether a manifest filename is a portable direct-child MP3 name."""
    if not filename or filename in {".", ".."}:
        return False
    if Path(filename).name != filename or "/" in filename or "\\" in filename:
        return False
    if filename.endswith((" ", ".")):
        return False
    if any(ord(character) < 32 for character in filename):
        return False
    if any(character in _WINDOWS_FORBIDDEN_FILENAME_CHARACTERS for character in filename):
        return False
    if Path(filename).suffix.casefold() != ".mp3":
        return False
    if Path(filename).stem.rstrip(" .").upper() in _WINDOWS_RESERVED_STEMS:
        return False
    return True


def legacy_audio_filename(section_id: str) -> str | None:
    """Return the historical ``<id>.mp3`` name only when it is portable."""
    candidate = f"{section_id}.mp3"
    return candidate if is_portable_audio_filename(candidate) else None


def load_audio_manifest_file_map(audio_dir: Path) -> dict[str, str]:
    """Load an exact logical-ID to portable physical-filename mapping."""
    manifest_path = audio_dir / "manifest.json"
    if not manifest_path.is_file():
        return {}
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"could not read audio manifest {manifest_path}: {exc}") from exc
    entries = payload.get("sections") if isinstance(payload, dict) else payload
    if not isinstance(entries, list):
        raise ValueError(f"audio manifest must be an array or contain a sections array: {manifest_path}")

    mapping: dict[str, str] = {}
    filename_owners: dict[str, str] = {}
    for index, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict):
            raise ValueError(f"audio manifest entry {index} must be an object: {manifest_path}")
        if "id" not in entry or "file" not in entry:
            raise ValueError(f"audio manifest entry {index} needs id and file: {manifest_path}")
        section_id = str(entry["id"])
        filename = str(entry["file"])
        if section_id in mapping:
            raise ValueError(f"duplicate section id {section_id!r} in {manifest_path}")
        if not is_portable_audio_filename(filename):
            raise ValueError(
                f"audio manifest entry {section_id!r} has a non-portable filename: {filename!r}"
            )
        folded_filename = filename.casefold()
        if folded_filename in filename_owners:
            raise ValueError(
                f"audio manifest filename collision between {filename_owners[folded_filename]!r} "
                f"and {section_id!r}: {filename!r}"
            )
        filename_owners[folded_filename] = section_id
        mapping[section_id] = filename
    return mapping


def resolve_audio_path(
    audio_dir: Path,
    section_id: str,
    *,
    file_map: dict[str, str] | None = None,
    allow_legacy: bool = True,
) -> Path:
    """Resolve one logical ID through the manifest, with a safe legacy fallback."""
    mapping = load_audio_manifest_file_map(audio_dir) if file_map is None else file_map
    filename = mapping.get(section_id)
    if filename is None and allow_legacy:
        filename = legacy_audio_filename(section_id)
    if filename is None:
        raise ValueError(
            f"audio section {section_id!r} has no manifest mapping and cannot use a portable legacy filename"
        )
    return audio_dir / filename
