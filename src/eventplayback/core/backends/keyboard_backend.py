"""Hotkeys backed by the ``keyboard`` library, for suppression on Windows.

pynput observes hotkeys but cannot stop them reaching the focused application,
so pressing F9 to stop a recording also sends F9 to whatever is in front. The
``keyboard`` library installs a low-level Windows hook that can swallow the
key. That is its only advantage here, and it costs an optional dependency plus
elevated privileges on Linux, so it is used only when suppression is asked for.

Importing this module requires the ``keyboard`` package; callers are expected
to handle :class:`ImportError`.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

import keyboard

from .base import HotkeyListener

logger = logging.getLogger(__name__)

__all__ = ["KeyboardHotkeyListener"]


class KeyboardHotkeyListener(HotkeyListener):
    """Global hotkeys that can be withheld from other applications.

    Accepts the same ``"ctrl+shift+f9"`` spelling as the settings file, which
    is already the syntax this library uses.
    """

    def __init__(self, *, suppress: bool = True) -> None:
        self._suppress = suppress
        self._handles: list[Any] = []

    def start(self, bindings: dict[str, Callable[[], None]]) -> None:
        self.stop()
        if not bindings:
            return
        try:
            for hotkey, callback in bindings.items():
                self._handles.append(
                    keyboard.add_hotkey(
                        hotkey, callback, suppress=self._suppress, trigger_on_release=False
                    )
                )
        except Exception as exc:
            self.stop()
            raise RuntimeError(f"Could not register hotkeys: {exc}") from exc

    def stop(self) -> None:
        # Remove only the hooks this listener installed; the process may have
        # others, and ``unhook_all`` would take them down too.
        for handle in self._handles:
            try:
                keyboard.remove_hotkey(handle)
            except Exception:
                logger.debug("Could not remove hotkey %r", handle, exc_info=True)
        self._handles = []
