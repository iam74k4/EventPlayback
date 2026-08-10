"""The pynput backed implementation of the input abstractions.

This is the only module in the project that imports pynput. Importing it
requires a usable display server, so it is loaded lazily by the UI rather than
by ``eventplayback.core``.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from pynput import keyboard as pynput_kb
from pynput import mouse as pynput_mouse
from pynput.keyboard import Controller as KeyboardController
from pynput.keyboard import Key, KeyCode
from pynput.mouse import Button
from pynput.mouse import Controller as MouseController

from .base import HotkeyListener, InputHandler, InputSource, InputSynthesizer

logger = logging.getLogger(__name__)

__all__ = [
    "PynputHotkeyListener",
    "PynputSource",
    "PynputSynthesizer",
    "key_to_name",
    "name_to_key",
]

#: Alternative spellings accepted when loading macros. The canonical names are
#: pynput's own ``Key`` member names, which differ per platform (``cmd`` only
#: exists on macOS and Windows, ``f13``-``f20`` only on some systems), so the
#: lookup table is derived from ``Key.__members__`` at call time instead of
#: being hard coded.
_KEY_ALIASES = {
    "escape": "esc",
    "return": "enter",
    "del": "delete",
    "win": "cmd",
    "super": "cmd",
    "meta": "cmd",
    "command": "cmd",
    "pgup": "page_up",
    "pgdn": "page_down",
    "capslock": "caps_lock",
    "numlock": "num_lock",
    "printscreen": "print_screen",
}


def key_to_name(key: Any) -> str | None:
    """Convert a pynput key object into a stable, serialisable name.

    Modifier combinations report control characters in ``char`` (``Ctrl+C``
    arrives as ``"\\x03"``), which would not survive a round trip, so the
    virtual key code is preferred in that case.
    """
    char = getattr(key, "char", None)
    if char:
        if ord(char[0]) >= 32:
            return char
        vk = getattr(key, "vk", None)
        if vk is not None and 32 <= vk < 127:
            return chr(vk).lower()

    name = getattr(key, "name", None)
    if name:
        return str(name)

    vk = getattr(key, "vk", None)
    if vk is not None:
        return f"vk_{vk}"
    return None


def name_to_key(name: str | None) -> Key | KeyCode | str | None:
    """Convert a stored key name back into something pynput can type."""
    if not name:
        return None

    if len(name) == 1:
        return name

    lowered = name.lower()
    lowered = _KEY_ALIASES.get(lowered, lowered)

    special = Key.__members__.get(lowered)
    if special is not None:
        return special

    if lowered.startswith("vk_"):
        try:
            return KeyCode.from_vk(int(lowered[3:]))
        except ValueError:
            pass

    logger.warning("Unknown key name %r will be ignored", name)
    return None


def button_to_name(button: Any) -> str:
    return str(getattr(button, "name", "left"))


def name_to_button(name: str | None) -> Button | None:
    if not name:
        return None
    resolved = Button.__members__.get(name.lower())
    if resolved is None:
        logger.warning("Unknown mouse button %r will be ignored", name)
    return resolved


class PynputSource(InputSource):
    """Captures mouse and keyboard input through pynput listeners."""

    #: How long to wait for a listener thread to wind down.
    JOIN_TIMEOUT = 0.5

    def __init__(self) -> None:
        self._mouse_listener: Any = None
        self._kb_listener: Any = None

    def start(self, handler: InputHandler) -> None:
        def on_move(x: float, y: float) -> None:
            handler.on_move(int(x), int(y))

        def on_click(x: float, y: float, button: Any, pressed: bool) -> None:
            handler.on_click(int(x), int(y), button_to_name(button), bool(pressed))

        def on_scroll(x: float, y: float, dx: float, dy: float) -> None:
            handler.on_scroll(int(x), int(y), int(dx), int(dy))

        def on_press(key: Any) -> None:
            name = key_to_name(key)
            if name:
                handler.on_key(name, True)

        def on_release(key: Any) -> None:
            name = key_to_name(key)
            if name:
                handler.on_key(name, False)

        try:
            self._mouse_listener = pynput_mouse.Listener(
                on_move=on_move, on_click=on_click, on_scroll=on_scroll
            )
            self._kb_listener = pynput_kb.Listener(on_press=on_press, on_release=on_release)
            self._mouse_listener.start()
            self._kb_listener.start()
        except Exception as exc:
            self.stop()
            raise RuntimeError(f"Could not start input capture: {exc}") from exc

    def stop(self) -> None:
        for listener in (self._mouse_listener, self._kb_listener):
            if listener is None:
                continue
            try:
                listener.stop()
                listener.join(timeout=self.JOIN_TIMEOUT)
            except Exception:
                logger.debug("Listener shutdown error", exc_info=True)
        self._mouse_listener = None
        self._kb_listener = None


class PynputSynthesizer(InputSynthesizer):
    """Replays events through pynput controllers."""

    def __init__(self) -> None:
        self._mouse = MouseController()
        self._kb = KeyboardController()

    def move(self, x: int, y: int) -> None:
        self._mouse.position = (x, y)

    def button(self, name: str, pressed: bool) -> None:
        resolved = name_to_button(name)
        if resolved is None:
            return
        if pressed:
            self._mouse.press(resolved)
        else:
            self._mouse.release(resolved)

    def scroll(self, dx: int, dy: int) -> None:
        if dy:
            self._mouse.scroll(0, dy)
        if dx:
            self._mouse.scroll(dx, 0)

    def key(self, name: str, pressed: bool) -> None:
        resolved = name_to_key(name)
        if resolved is None:
            return
        if pressed:
            self._kb.press(resolved)
        else:
            self._kb.release(resolved)


class PynputHotkeyListener(HotkeyListener):
    """Global hotkeys via ``pynput.keyboard.GlobalHotKeys``.

    Using pynput here rather than a second library means a single keyboard hook
    serves both recording and hotkeys, and no elevated privileges are needed on
    Windows.
    """

    def __init__(self) -> None:
        self._listener: Any = None

    @staticmethod
    def _to_pynput_spec(hotkey: str) -> str:
        """Translate ``"ctrl+f9"`` into pynput's ``"<ctrl>+<f9>"`` syntax."""
        parts = []
        for raw in hotkey.split("+"):
            token = raw.strip().lower()
            if not token:
                continue
            token = _KEY_ALIASES.get(token, token)
            parts.append(f"<{token}>" if token in Key.__members__ else token)
        return "+".join(parts)

    def start(self, bindings: dict[str, Callable[[], None]]) -> None:
        self.stop()
        if not bindings:
            return
        spec = {self._to_pynput_spec(key): callback for key, callback in bindings.items()}
        try:
            self._listener = pynput_kb.GlobalHotKeys(spec)
            self._listener.start()
        except Exception as exc:
            self._listener = None
            raise RuntimeError(f"Could not register hotkeys: {exc}") from exc

    def stop(self) -> None:
        if self._listener is None:
            return
        try:
            self._listener.stop()
        except Exception:
            logger.debug("Hotkey listener shutdown error", exc_info=True)
        self._listener = None
