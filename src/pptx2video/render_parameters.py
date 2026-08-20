"""Shared validation for video timing parameters."""

from __future__ import annotations

import math


DEFAULT_FPS = 30
DEFAULT_START_PAD = 0.5
DEFAULT_PAD_TAIL = 0.3


def validate_render_timing(*, fps: int, start_pad: float, pad_tail: float) -> None:
    """Reject timing combinations that cannot produce a valid lead-in video."""
    if fps <= 0:
        raise ValueError("--fps must be positive")
    if not math.isfinite(start_pad) or not math.isfinite(pad_tail):
        raise ValueError("--start-pad and --pad-tail must be finite")
    if start_pad < 0 or pad_tail < 0:
        raise ValueError("--start-pad and --pad-tail must be non-negative")
    if start_pad > 0 and fps * start_pad < 1:
        minimum_start_pad = 1.0 / fps
        raise ValueError(
            "--start-pad must be 0 or span at least one output video frame; "
            f"--fps {fps} with --start-pad {start_pad:g} spans only "
            f"{fps * start_pad:g} frames (minimum {minimum_start_pad:g}s)"
        )
