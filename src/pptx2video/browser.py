"""Cross-platform Chromium-family browser discovery for SVG rendering."""

from __future__ import annotations

import os
import platform
import shutil
from collections.abc import Mapping
from pathlib import Path


def executable_file(path: Path) -> str | None:
    """Return a normalized path only when it names an executable file."""
    expanded = path.expanduser()
    if expanded.is_file() and os.access(expanded, os.X_OK):
        return str(expanded.resolve())
    return None


def find_system_chromium(environment: Mapping[str, str]) -> str | None:
    """Return an executable system Chromium-family browser without launching it."""
    for variable in (
        "PPTX2VIDEO_CHROME",
        "PAPER2VIDEO_CHROME",
        "CHROME",
        "CHROMIUM",
    ):
        configured = environment.get(variable)
        if configured:
            executable = executable_file(Path(configured))
            if executable:
                return executable

    candidates: list[Path] = []
    system = platform.system()
    if system == "Darwin":
        candidates.extend(
            Path(path)
            for path in (
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            )
        )
    elif system == "Windows":
        program_files = Path(environment.get("ProgramFiles", "C:/Program Files"))
        program_files_x86 = Path(
            environment.get("ProgramFiles(x86)", "C:/Program Files (x86)")
        )
        local_app_data = Path(
            environment.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
        )
        candidates.extend(
            (
                program_files / "Google" / "Chrome" / "Application" / "chrome.exe",
                program_files_x86 / "Google" / "Chrome" / "Application" / "chrome.exe",
                local_app_data / "Google" / "Chrome" / "Application" / "chrome.exe",
                local_app_data / "Chromium" / "Application" / "chrome.exe",
                program_files / "Microsoft" / "Edge" / "Application" / "msedge.exe",
                program_files_x86 / "Microsoft" / "Edge" / "Application" / "msedge.exe",
            )
        )
    for candidate in candidates:
        executable = executable_file(candidate)
        if executable:
            return executable

    search = environment.get("PATH", "").split(os.pathsep)
    if system == "Darwin":
        search = ["/opt/homebrew/bin", "/usr/local/bin", *search]
    elif system == "Windows":
        search = [
            str(local_app_data / "Microsoft" / "WinGet" / "Links"),
            *search,
        ]
    search_path = os.pathsep.join(dict.fromkeys(search))
    for name in (
        "google-chrome",
        "google-chrome-stable",
        "chromium",
        "chromium-browser",
        "msedge",
        "microsoft-edge",
        "microsoft-edge-stable",
    ):
        executable = shutil.which(name, path=search_path)
        if executable:
            return executable
    return None
