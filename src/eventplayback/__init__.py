"""EventPlayback - record and play back mouse and keyboard input.

The package is split into three layers:

``eventplayback.core``
    Recording and playback logic. Imports neither a GUI toolkit nor any OS
    input library, so it can be exercised head-less on any platform.
``eventplayback.services``
    Settings persistence, macro file I/O and hotkey registration.
``eventplayback.ui``
    The customtkinter front-end.
"""

from __future__ import annotations

import os

__all__ = ["__version__"]

# Kept in sync with ``version`` in pyproject.toml. Used when the package is run
# straight from a source checkout, where no distribution metadata exists.
_FALLBACK_VERSION = "1.1.0"


def _resolve_version() -> str:
    override = os.getenv("EVENTPLAYBACK_VERSION", "").strip()
    if override:
        return override
    try:
        from importlib.metadata import version

        return version("eventplayback")
    except Exception:
        return _FALLBACK_VERSION


__version__ = _resolve_version()
