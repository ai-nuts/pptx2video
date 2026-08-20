"""Shared FFmpeg discovery for every pptx2video entry point.

The doctor and runtime stages must select the same binaries.  Keep all PATH
augmentation, environment-variable compatibility, and imageio-ffmpeg fallback
logic here so direct ``python -m pptx2video.*`` calls behave like the CLI.
"""

from __future__ import annotations

import os
import platform
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping


@dataclass(frozen=True)
class FFmpegResolution:
    ffmpeg: str | None
    ffprobe: str | None
    source: str | None


def runtime_environment(
    source: Mapping[str, str] | None = None,
    *,
    platform_name: str | None = None,
) -> dict[str, str]:
    """Return a cross-platform process environment used by doctor and runtime."""
    environment = dict(os.environ if source is None else source)
    for modern, legacy in (
        ("PPTX2VIDEO_FFMPEG", "PAPER2VIDEO_FFMPEG"),
        ("PPTX2VIDEO_FFPROBE", "PAPER2VIDEO_FFPROBE"),
    ):
        if not environment.get(modern) and environment.get(legacy):
            environment[modern] = environment[legacy]

    search = [entry for entry in environment.get("PATH", "").split(os.pathsep) if entry]
    system = platform_name or platform.system()
    if system == "Darwin":
        search = [
            "/Applications/LibreOffice.app/Contents/MacOS",
            "/opt/homebrew/bin",
            "/usr/local/bin",
            *search,
        ]
    elif system == "Windows":
        local_app_data = Path(
            environment.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
        )
        user_profile = Path(environment.get("USERPROFILE", Path.home()))
        search = [
            str(
                Path(environment.get("ProgramFiles", "C:/Program Files"))
                / "LibreOffice"
                / "program"
            ),
            str(local_app_data / "Microsoft" / "WinGet" / "Links"),
            str(user_profile / "scoop" / "shims"),
            "C:/ProgramData/chocolatey/bin",
            *search,
        ]
    environment["PATH"] = os.pathsep.join(dict.fromkeys(search))
    return environment


def _executable_file(value: str | None) -> str | None:
    if not value:
        return None
    path = Path(value).expanduser()
    if not path.is_file():
        return None
    return str(path.resolve())


def _first_executable(*values: str | None) -> str | None:
    """Return the first candidate that names an existing file."""
    for value in values:
        resolved = _executable_file(value)
        if resolved:
            return resolved
    return None


def _imageio_ffmpeg_binary() -> str | None:
    try:
        import imageio_ffmpeg  # type: ignore

        return _executable_file(imageio_ffmpeg.get_ffmpeg_exe())
    except Exception:
        return None


def resolve_ffmpeg(
    environment: Mapping[str, str] | None = None,
) -> FFmpegResolution:
    """Resolve the exact FFmpeg/FFprobe pair used by every runtime stage.

    Selection order preserves the historical renderer behavior: explicit
    pptx2video variables (including paper2video compatibility aliases), the
    generic ``FFMPEG_BINARY`` compatibility override, imageio-ffmpeg's bundled
    static binary, and finally augmented system PATH discovery.
    """
    prepared = runtime_environment(environment)
    explicit_ffmpeg = _first_executable(
        prepared.get("PPTX2VIDEO_FFMPEG"),
        prepared.get("PAPER2VIDEO_FFMPEG"),
        prepared.get("FFMPEG_BINARY"),
    )
    explicit_ffprobe = _first_executable(
        prepared.get("PPTX2VIDEO_FFPROBE"),
        prepared.get("PAPER2VIDEO_FFPROBE"),
    )
    if explicit_ffmpeg:
        return FFmpegResolution(
            ffmpeg=explicit_ffmpeg,
            ffprobe=explicit_ffprobe or explicit_ffmpeg,
            source="environment",
        )

    bundled = _imageio_ffmpeg_binary()
    if bundled:
        return FFmpegResolution(
            ffmpeg=bundled,
            ffprobe=explicit_ffprobe or bundled,
            source="imageio-ffmpeg",
        )

    ffmpeg = shutil.which("ffmpeg", path=prepared.get("PATH"))
    if ffmpeg:
        resolved_ffmpeg = str(Path(ffmpeg).resolve())
        ffprobe = shutil.which("ffprobe", path=prepared.get("PATH"))
        return FFmpegResolution(
            ffmpeg=resolved_ffmpeg,
            ffprobe=(
                explicit_ffprobe
                or (str(Path(ffprobe).resolve()) if ffprobe else resolved_ffmpeg)
            ),
            source="path",
        )
    return FFmpegResolution(ffmpeg=None, ffprobe=None, source=None)


def find_ffmpeg_pair(
    environment: Mapping[str, str] | None = None,
    *,
    required: bool = True,
    component: str = "pptx2video",
) -> tuple[str | None, str | None]:
    resolution = resolve_ffmpeg(environment)
    if resolution.ffmpeg:
        return resolution.ffmpeg, resolution.ffprobe
    if not required:
        return None, None
    raise SystemExit(
        f"[{component}] ffmpeg not found and imageio_ffmpeg is not installed.\n"
        "Pick one:\n"
        "  - System install: install ffmpeg and place it on PATH\n"
        "  - Python fallback: pip install imageio-ffmpeg\n"
        "  - Explicit binary: set PPTX2VIDEO_FFMPEG\n"
    )


def resolved_runtime_environment(
    source: Mapping[str, str] | None = None,
) -> dict[str, str]:
    """Return the runtime environment with the shared selection pinned."""
    environment = runtime_environment(source)
    resolution = resolve_ffmpeg(environment)
    if resolution.ffmpeg:
        environment["PPTX2VIDEO_FFMPEG"] = resolution.ffmpeg
    if resolution.ffprobe:
        environment["PPTX2VIDEO_FFPROBE"] = resolution.ffprobe
    return environment
