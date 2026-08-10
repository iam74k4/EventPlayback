"""Pluggable input backends.

This package intentionally does **not** import any concrete backend at module
level: importing :mod:`eventplayback.core.backends.pynput_backend` pulls in
pynput, which needs a display server. Tests and head-less tooling can therefore
import the abstractions without requiring one.
"""

from __future__ import annotations

from .base import (
    Clock,
    HotkeyListener,
    InputHandler,
    InputSource,
    InputSynthesizer,
    RealClock,
)

__all__ = [
    "Clock",
    "HotkeyListener",
    "InputHandler",
    "InputSource",
    "InputSynthesizer",
    "RealClock",
]
