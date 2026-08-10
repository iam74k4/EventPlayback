"""Platform specific helpers.

Everything that needs to know which OS we are running on lives here, so the
rest of the code base can stay platform agnostic.
"""

from __future__ import annotations

import logging
import os
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

logger = logging.getLogger(__name__)

APP_NAME = "EventPlayback"

IS_WINDOWS = sys.platform.startswith("win")
IS_MACOS = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")


def enable_dpi_awareness() -> None:
    """Opt into per-monitor DPI scaling on Windows.

    Without this the coordinates reported while recording do not match the
    coordinates used during playback on scaled displays. macOS and Linux
    report logical pixels consistently, so this is a no-op there.
    """
    if not IS_WINDOWS:
        return
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # System DPI Aware (fallback)
    except Exception:
        logger.debug("Could not enable DPI awareness", exc_info=True)


@contextmanager
def high_resolution_timer() -> Iterator[bool]:
    """Raise this process's timer resolution to 1 ms for the duration of the block.

    Windows defaults to a ~15.6 ms scheduling tick, and a blocking wait is
    rounded up to it. Replaying a macro recorded at 20 ms resolution would then
    round every gap to 31 ms, so playback must ask for a finer tick and give it
    back afterwards. Timer resolution is per-process on Windows 10 2004 and
    later, so this does not affect the rest of the system.

    Yields:
        Whether fine-grained waits are available. macOS and Linux already
        provide them, so this is a no-op yielding ``True`` there.
    """
    if not IS_WINDOWS:
        yield True
        return

    import ctypes

    TIMERR_NOERROR = 0
    period_ms = 1
    try:
        winmm = ctypes.WinDLL("winmm")
        granted = winmm.timeBeginPeriod(period_ms) == TIMERR_NOERROR
    except Exception:
        logger.warning("Could not raise the timer resolution; playback timing may be coarse")
        yield False
        return

    if not granted:
        logger.warning("The system refused a %d ms timer resolution", period_ms)
    try:
        yield granted
    finally:
        if granted:
            try:
                winmm.timeEndPeriod(period_ms)
            except Exception:
                logger.debug("timeEndPeriod failed", exc_info=True)


def config_dir() -> Path:
    """Return the per-user directory holding EventPlayback's settings."""
    if IS_WINDOWS:
        base = os.getenv("APPDATA") or Path.home() / "AppData" / "Roaming"
    elif IS_MACOS:
        base = Path.home() / "Library" / "Application Support"
    else:
        base = os.getenv("XDG_CONFIG_HOME") or Path.home() / ".config"
    return Path(base) / APP_NAME


def permission_hint() -> str:
    """Return an OS specific hint shown when input hooks cannot be installed."""
    if IS_MACOS:
        return "Grant Accessibility and Input Monitoring permission in System Settings"
    if IS_LINUX:
        return "An X11 session is required; Wayland is not supported"
    return "Try running the application as administrator"
