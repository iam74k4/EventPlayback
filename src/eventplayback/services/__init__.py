"""Application services: settings, macro files and hotkeys."""

from __future__ import annotations

from .hotkeys import HotkeyManager
from .settings import Settings
from .storage import MacroFileError, load_macro, save_macro

__all__ = [
    "HotkeyManager",
    "MacroFileError",
    "Settings",
    "load_macro",
    "save_macro",
]
