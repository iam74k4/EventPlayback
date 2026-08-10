"""Platform specific helpers.

Everything that needs to know which OS we are running on lives here, so the
rest of the code base can stay platform agnostic.
"""

from __future__ import annotations

import logging
import os
import sys
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
