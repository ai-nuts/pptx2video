#!/usr/bin/env python3
"""Installable CLI for the standalone deterministic pptx2video runtime."""

from __future__ import annotations

import argparse
import importlib.util
import json
import platform
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

from .browser import executable_file, find_system_chromium
from .ffmpeg import resolve_ffmpeg, resolved_runtime_environment
from .render_parameters import (
    DEFAULT_FPS,
    DEFAULT_PAD_TAIL,
    DEFAULT_START_PAD,
    validate_render_timing,
)


VERSION = "0.5.0"
PYTHON_MODULES = {
    "edge_tts": "edge-tts",
    "lxml": "lxml",
    "PIL": "Pillow",
    "numpy": "numpy",
    "pptx": "python-pptx",
}
OPTIONAL_PYTHON_MODULES = {"openai": "openai"}
REQUIRED_FFMPEG_FILTERS = (
    "ass",
    "aformat",
    "aresample",
    "concat",
    "drawbox",
    "format",
    "geq",
    "overlay",
    "scale",
)
REQUIRED_FFMPEG_ENCODERS = ("aac", "libmp3lame", "libx264")


def _runtime_dir() -> Path:
    here = Path(__file__).resolve().parent
    if (here / "render_edited_pptx.py").is_file():
        return here
    raise SystemExit(f"installed pptx2video runtime not found: {here}")


def _runtime_environment() -> dict[str, str]:
    return resolved_runtime_environment()


def _command_path(name: str, environment: dict[str, str]) -> str | None:
    return shutil.which(name, path=environment.get("PATH"))


def _playwright_chromium_executable() -> str | None:
    """Return Playwright's bundled Chromium path without launching a browser."""
    try:
        from playwright.sync_api import sync_playwright  # type: ignore

        with sync_playwright() as playwright:
            executable_path = Path(playwright.chromium.executable_path)
    except Exception:
        return None
    return executable_file(executable_path)


def _svg_capabilities(environment: dict[str, str]) -> dict[str, object]:
    playwright_available = importlib.util.find_spec("playwright") is not None
    system_browser = find_system_chromium(environment)
    bundled_browser = None
    if playwright_available and system_browser is None:
        bundled_browser = _playwright_chromium_executable()
    browser = system_browser or bundled_browser
    return {
        "requested": True,
        "playwright": playwright_available,
        "browser": browser,
        "browser_source": (
            "system" if system_browser else "playwright" if bundled_browser else None
        ),
        "passed": bool(playwright_available and browser),
    }


def _ffmpeg_listing_names(output: str, *, flag_width: int) -> set[str]:
    names: set[str] = set()
    for line in output.splitlines():
        fields = line.split()
        if len(fields) < 2:
            continue
        flags = fields[0]
        if len(flags) != flag_width or any(char not in ".AFNSTVXCD" for char in flags):
            continue
        if fields[1] == "=":
            continue
        names.add(fields[1])
    return names


def _ffmpeg_capabilities(ffmpeg: str | None) -> dict[str, object]:
    filters = {name: False for name in REQUIRED_FFMPEG_FILTERS}
    encoders = {name: False for name in REQUIRED_FFMPEG_ENCODERS}
    if ffmpeg is None:
        return {
            "checked": False,
            "binary": None,
            "filters": filters,
            "encoders": encoders,
            "passed": False,
        }
    try:
        filter_probe = subprocess.run(
            [ffmpeg, "-hide_banner", "-filters"],
            capture_output=True,
            text=True,
            check=False,
            timeout=20,
        )
        encoder_probe = subprocess.run(
            [ffmpeg, "-hide_banner", "-encoders"],
            capture_output=True,
            text=True,
            check=False,
            timeout=20,
        )
    except (OSError, subprocess.TimeoutExpired):
        return {
            "checked": True,
            "binary": ffmpeg,
            "filters": filters,
            "encoders": encoders,
            "passed": False,
        }
    filter_names = _ffmpeg_listing_names(
        filter_probe.stdout + filter_probe.stderr,
        flag_width=3,
    )
    encoder_names = _ffmpeg_listing_names(
        encoder_probe.stdout + encoder_probe.stderr,
        flag_width=6,
    )
    filters = {name: name in filter_names for name in REQUIRED_FFMPEG_FILTERS}
    encoders = {name: name in encoder_names for name in REQUIRED_FFMPEG_ENCODERS}
    passed = bool(
        filter_probe.returncode == 0
        and encoder_probe.returncode == 0
        and all(filters.values())
        and all(encoders.values())
    )
    return {
        "checked": True,
        "binary": ffmpeg,
        "filters": filters,
        "encoders": encoders,
        "passed": passed,
    }


def _doctor_payload(*, check_svg: bool = False) -> dict[str, object]:
    environment = _runtime_environment()
    ffmpeg_resolution = resolve_ffmpeg(environment)
    commands = {
        "ffmpeg": ffmpeg_resolution.ffmpeg,
        "ffprobe": ffmpeg_resolution.ffprobe,
        "pdftoppm": _command_path("pdftoppm", environment),
    }
    libreoffice = _command_path("libreoffice", environment) or _command_path(
        "soffice", environment
    )
    ffmpeg_capabilities = _ffmpeg_capabilities(commands["ffmpeg"])
    modules = {
        package: importlib.util.find_spec(module) is not None
        for module, package in PYTHON_MODULES.items()
    }
    optional_modules = {
        package: importlib.util.find_spec(module) is not None
        for module, package in OPTIONAL_PYTHON_MODULES.items()
    }
    try:
        runtime = str(_runtime_dir())
    except SystemExit:
        runtime = None
    python_ok = sys.version_info >= (3, 11)
    passed = bool(
        python_ok
        and runtime
        and libreoffice
        and all(commands.values())
        and all(modules.values())
        and ffmpeg_capabilities["passed"]
    )
    payload: dict[str, object] = {
        "passed": passed,
        "platform": platform.platform(),
        "python": sys.version,
        "python_ok": python_ok,
        "runtime": runtime,
        "commands": {**commands, "libreoffice": libreoffice},
        "ffmpeg_capabilities": ffmpeg_capabilities,
        "python_packages": modules,
        "optional_python_packages": optional_modules,
        "edge_tts_network_required": True,
    }
    if check_svg:
        svg_rendering = _svg_capabilities(environment)
        payload["svg_rendering"] = svg_rendering
        payload["passed"] = bool(passed and svg_rendering["passed"])
    return payload


def _doctor(as_json: bool, *, check_svg: bool = False) -> int:
    payload = _doctor_payload(check_svg=check_svg)
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(f"Python >= 3.11: {'OK' if payload['python_ok'] else 'MISSING'}")
        print(f"pptx2video runtime: {payload['runtime'] or 'MISSING'}")
        for name, path in payload["commands"].items():
            print(f"{name}: {path or 'MISSING'}")
        ffmpeg_capabilities = payload["ffmpeg_capabilities"]
        for name, present in ffmpeg_capabilities["filters"].items():
            print(f"FFmpeg filter {name}: {'OK' if present else 'MISSING'}")
        for name, present in ffmpeg_capabilities["encoders"].items():
            print(f"FFmpeg encoder {name}: {'OK' if present else 'MISSING'}")
        for name, present in payload["python_packages"].items():
            print(f"Python package {name}: {'OK' if present else 'MISSING'}")
        for name, present in payload["optional_python_packages"].items():
            print(f"Optional Python package {name}: {'OK' if present else 'MISSING'}")
        svg_rendering = payload.get("svg_rendering")
        if isinstance(svg_rendering, dict):
            print(
                "Python package playwright: "
                f"{'OK' if svg_rendering['playwright'] else 'MISSING'}"
            )
            browser = svg_rendering["browser"] or "MISSING"
            source = svg_rendering["browser_source"]
            source_label = f" ({source})" if source else ""
            print(f"SVG Chromium browser: {browser}{source_label}")
        print("Edge TTS: network access required during rendering")
        print(f"Overall: {'PASS' if payload['passed'] else 'FAIL'}")
    return 0 if payload["passed"] else 2


def _run_runtime(module: str, arguments: list[str]) -> None:
    _runtime_dir()
    module_name = Path(module).stem
    command = [sys.executable, "-m", f"pptx2video.{module_name}", *arguments]
    print("[pptx2video] $ " + shlex.join(command), flush=True)
    subprocess.run(command, check=True, env=_runtime_environment())


def _render(args: argparse.Namespace) -> int:
    pptx = args.pptx.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if not pptx.is_file():
        raise SystemExit(f"PPTX not found: {pptx}")
    if pptx.suffix.lower() != ".pptx":
        raise SystemExit(f"input must be a .pptx file: {pptx}")
    if output.exists():
        raise SystemExit(f"output bundle already exists; choose a fresh path: {output}")
    try:
        validate_render_timing(
            fps=args.fps,
            start_pad=DEFAULT_START_PAD,
            pad_tail=DEFAULT_PAD_TAIL,
        )
    except ValueError as exc:
        raise SystemExit(f"invalid render timing: {exc}") from exc

    command = [
        str(pptx),
        str(output),
        "--resolution",
        args.resolution,
        "--fps",
        str(args.fps),
        "--rate",
        args.rate,
        "--highlight-style",
        args.highlight_style,
    ]
    if args.voice:
        command.extend(["--voice", args.voice])
    if args.tts_cache_dir:
        command.extend(
            ["--tts-cache-dir", str(args.tts_cache_dir.expanduser().resolve())]
        )
    elif args.no_tts_cache:
        command.append("--no-tts-cache")
    if args.ids_from_script:
        command.extend(
            ["--ids-from-script", str(args.ids_from_script.expanduser().resolve())]
        )
    if args.script_json:
        command.extend(["--script-json", str(args.script_json.expanduser().resolve())])
    if args.baseline_pptx:
        command.extend(
            ["--baseline-pptx", str(args.baseline_pptx.expanduser().resolve())]
        )
    command.extend(["--narration-mode", args.narration_mode])
    command.extend(
        ["--narration-order-policy", args.narration_order_policy]
    )
    command.extend(["--animation-order-policy", args.animation_order_policy])
    for identity_order in args.animation_order_sequence:
        command.extend(["--animation-order-sequence", identity_order])
    command.extend(["--semantic-profile", args.semantic_profile])
    command.extend(["--click-group-policy", args.click_group_policy])
    if args.animation_order_report:
        command.extend(
            [
                "--animation-order-report",
                str(args.animation_order_report.expanduser().resolve()),
            ]
        )
    if args.narration_mode == "regenerate":
        if importlib.util.find_spec("openai") is None:
            raise SystemExit(
                "narration regeneration requires the optional openai package"
            )
        command.extend(["--regeneration-model", args.regeneration_model])
    if args.no_subtitles:
        command.append("--no-subtitles")
    if args.keep_temp:
        command.append("--keep-temp")

    try:
        _run_runtime("render_edited_pptx.py", command)
    except subprocess.CalledProcessError as exc:
        return exc.returncode or 1

    report_path = output / "assets" / "meta" / "reports" / "video_qa_report.json"
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"strict QA report is missing or invalid: {report_path}: {exc}") from exc
    counts = report.get("counts") or {}
    if not report.get("passed") or counts.get("error") or counts.get("warning"):
        raise SystemExit(f"strict QA did not pass cleanly: {report_path}")
    print(f"[pptx2video] Strict QA passed with 0 errors and 0 warnings: {report_path}")
    print(f"[pptx2video] Video: {output / 'video.mp4'}")
    return 0


def _bootstrap(args: argparse.Namespace) -> int:
    pptx = args.pptx.expanduser().resolve()
    script = args.script_json.expanduser().resolve()
    output = args.output.expanduser().resolve()
    report = (
        args.report.expanduser().resolve()
        if args.report
        else output.with_suffix(".bootstrap-report.json")
    )
    for path, label in ((pptx, "PPTX"), (script, "script.json")):
        if not path.is_file():
            raise SystemExit(f"{label} not found: {path}")
    if output.exists():
        raise SystemExit(f"bootstrap output already exists: {output}")
    try:
        _run_runtime(
            "bootstrap_editable_pptx.py",
            [str(pptx), "--script-json", str(script), "--out", str(output), "--report-out", str(report)],
        )
    except subprocess.CalledProcessError as exc:
        return exc.returncode or 1
    print(f"[pptx2video] Editable PPTX: {output}")
    print(f"[pptx2video] Bootstrap report: {report}")
    return 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pptx2video", description=__doc__)
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    doctor = subparsers.add_parser("doctor", help="Check Python and native rendering dependencies")
    doctor.add_argument("--json", action="store_true")
    doctor.add_argument(
        "--svg",
        action="store_true",
        help="Also check Playwright and a system or bundled Chromium browser",
    )

    render = subparsers.add_parser("render", help="Render a fresh strictly checked video bundle")
    render.add_argument("pptx", type=Path, help="Edited PowerPoint file at any local path")
    render.add_argument("output", type=Path, help="New output bundle path; it must not exist")
    render.add_argument("--resolution", choices=("720p", "1080p", "1440p", "4k"), default="1080p")
    render.add_argument("--fps", type=int, default=DEFAULT_FPS)
    render.add_argument("--voice", default=None, help="Edge TTS voice")
    render.add_argument("--rate", default="+0%", help="Edge TTS rate")
    tts_cache = render.add_mutually_exclusive_group()
    tts_cache.add_argument(
        "--tts-cache-dir",
        type=Path,
        default=None,
        help=(
            "Persistent content-addressed block TTS cache. The default uses "
            "the platform user cache directory."
        ),
    )
    tts_cache.add_argument(
        "--no-tts-cache",
        action="store_true",
        help="Disable reuse of pristine block-level Edge TTS artifacts.",
    )
    render.add_argument("--ids-from-script", type=Path, default=None, help="Read stable section IDs only")
    render.add_argument(
        "--script-json",
        type=Path,
        default=None,
        help="Use a user-edited script.json as narration authority",
    )
    render.add_argument(
        "--baseline-pptx",
        type=Path,
        default=None,
        help="Previous PPTX for change detection",
    )
    render.add_argument(
        "--narration-mode",
        choices=("keep", "regenerate"),
        default="keep",
        help="Keep PPTX narration or use OpenAI only for changed elements",
    )
    render.add_argument(
        "--narration-order-policy",
        choices=("geometry", "author-notes"),
        default="geometry",
        help=(
            "Preserve canonical Notes block order and insert only Notes-missing "
            "targets by deterministic slide geometry (default); author-notes is "
            "a compatibility alias"
        ),
    )
    render.add_argument(
        "--regeneration-model",
        default="gpt-5.6-sol",
        help="OpenAI model for changed-element narration regeneration",
    )
    render.add_argument(
        "--highlight-style",
        choices=(
            "box",
            "spotlight",
            "cursor",
            "box_cursor",
            "spotlight_cursor",
            "laser",
            "box_laser",
            "spotlight_laser",
        ),
        default="spotlight_laser",
    )
    render.add_argument(
        "--animation-order-policy",
        choices=("auto", "animation-pane", "reading-order"),
        default="auto",
        help=(
            "Ask when narration order conflicts with Animation Pane position. "
            "Non-TTY auto and reading-order exit 3 without rendering."
        ),
    )
    render.add_argument(
        "--animation-order-sequence",
        action="append",
        default=[],
        metavar="[SLIDE=]ORDER",
        help=(
            "Deterministic Identity Map permutation such as CABGDEF. Repeat "
            "as SLIDE=ORDER for multiple conflict pages"
        ),
    )
    render.add_argument(
        "--animation-order-report",
        type=Path,
        default=None,
        help="Preflight decision-report path when rendering stops before bundle creation",
    )
    render.add_argument(
        "--semantic-profile",
        choices=("concise", "detailed"),
        default="concise",
        help=(
            "Choose the persisted semantic description shown beside A/B/C "
            "element identities in animation-order conflict previews"
        ),
    )
    render.add_argument(
        "--click-group-policy",
        choices=("normalize", "preserve"),
        default="normalize",
        help=(
            "Give every non-With-Previous phase a numbered PowerPoint click "
            "badge on every slide (default), or preserve the source mainSeq."
        ),
    )
    render.add_argument("--no-subtitles", action="store_true")
    render.add_argument("--keep-temp", action="store_true")

    bootstrap = subparsers.add_parser(
        "bootstrap",
        help="Add canonical Notes and Alt Text to an ordinary animated PPTX copy",
    )
    bootstrap.add_argument("pptx", type=Path)
    bootstrap.add_argument("--script-json", type=Path, required=True)
    bootstrap.add_argument("--output", type=Path, required=True)
    bootstrap.add_argument("--report", type=Path, default=None)
    return parser


def main() -> int:
    args = _parser().parse_args()
    if args.command == "doctor":
        return _doctor(args.json, check_svg=args.svg)
    if args.command == "render":
        return _render(args)
    if args.command == "bootstrap":
        return _bootstrap(args)
    raise SystemExit(f"unknown command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
