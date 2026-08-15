"""Turning a Tk key event into a hotkey string, and back into a caption.

Rebinding a hotkey by typing it is much easier than spelling it correctly in a
settings file, but Tk names keys its own way (``Escape``, ``Prior``, ``F9``)
while the hotkey listener expects pynput's names (``esc``, ``page_up``,
``f9``). The translation lives here, free of any toolkit import, so it can be
unit tested without a display.
"""

from __future__ import annotations

__all__ = [
    "MODIFIER_MASKS",
    "binding_from_event",
    "describe_binding",
]

#: Tk keysyms that are modifiers. Pressing one is not a complete binding, so
#: capture keeps waiting for the key that goes with it.
_MODIFIER_KEYSYMS = frozenset(
    {
        "shift_l", "shift_r", "control_l", "control_r", "alt_l", "alt_r",
        "meta_l", "meta_r", "super_l", "super_r", "hyper_l", "hyper_r",
        "caps_lock", "num_lock", "scroll_lock", "iso_level3_shift", "mode_switch",
    }
)

#: Tk names punctuation rather than reporting the character, so ``!`` arrives
#: as ``exclam``. The listener wants the character itself.
_PUNCTUATION = {
    "exclam": "!", "at": "@", "numbersign": "#", "dollar": "$", "percent": "%",
    "asciicircum": "^", "ampersand": "&", "asterisk": "*", "parenleft": "(",
    "parenright": ")", "minus": "-", "underscore": "_", "plus": "+", "equal": "=",
    "bracketleft": "[", "bracketright": "]", "braceleft": "{", "braceright": "}",
    "backslash": "\\", "bar": "|", "semicolon": ";", "colon": ":",
    "apostrophe": "'", "quotedbl": '"', "comma": ",", "less": "<", "period": ".",
    "greater": ">", "slash": "/", "question": "?", "grave": "`", "asciitilde": "~",
}

#: Tk keysym -> the name the hotkey listener understands.
_KEYSYM_NAMES = {
    "escape": "esc",
    "return": "enter",
    "kp_enter": "enter",
    "prior": "page_up",
    "next": "page_down",
    "back": "backspace",
    "backspace": "backspace",
    "delete": "delete",
    "insert": "insert",
    "space": "space",
    "tab": "tab",
    "up": "up",
    "down": "down",
    "left": "left",
    "right": "right",
    "home": "home",
    "end": "end",
    "print": "print_screen",
    "pause": "pause",
}

#: Which state bit means which modifier. Tk reports these differently per
#: platform: Alt is Mod1 under X11, a high bit on Windows, and Command takes
#: Mod1's place on macOS.
MODIFIER_MASKS: dict[str, dict[str, int]] = {
    "linux": {"ctrl": 0x0004, "alt": 0x0008, "shift": 0x0001, "cmd": 0x0040},
    "win32": {"ctrl": 0x0004, "alt": 0x20000, "shift": 0x0001},
    "darwin": {"ctrl": 0x0004, "alt": 0x0010, "shift": 0x0001, "cmd": 0x0008},
}

#: Modifiers always appear in this order, so the same combination always
#: produces the same string no matter which key was held first.
_MODIFIER_ORDER = ("ctrl", "alt", "shift", "cmd")

_DISPLAY_NAMES = {
    "ctrl": "Ctrl",
    "alt": "Alt",
    "shift": "Shift",
    "cmd": "Cmd",
    "esc": "Esc",
    "enter": "Enter",
    "page_up": "PgUp",
    "page_down": "PgDn",
    "print_screen": "PrtSc",
    "caps_lock": "CapsLock",
    "backspace": "Backspace",
}


def _key_name(keysym: str) -> str | None:
    """The listener's name for a Tk keysym, or ``None`` if it cannot bind it.

    Rejecting the unknown is deliberate: a binding the listener cannot parse
    would register as a failure at startup, which is a worse outcome than
    declining the key while the user is still watching the settings window.
    """
    lowered = keysym.lower()
    if lowered in _MODIFIER_KEYSYMS:
        return None
    if lowered in _PUNCTUATION:
        return _PUNCTUATION[lowered]
    if lowered in _KEYSYM_NAMES:
        return _KEYSYM_NAMES[lowered]
    if len(lowered) == 1 and lowered.isprintable():
        return lowered
    if lowered.startswith("f") and lowered[1:].isdigit():
        return lowered
    return None


def binding_from_event(keysym: str, state: int, *, platform: str = "linux") -> str | None:
    """Build a binding such as ``"ctrl+shift+f9"`` from a Tk key event.

    Returns ``None`` while the only thing held down is a modifier, which is
    what happens on the way to pressing the real key.
    """
    key = _key_name(keysym or "")
    if key is None:
        return None

    masks = MODIFIER_MASKS.get(platform, MODIFIER_MASKS["linux"])
    parts = [name for name in _MODIFIER_ORDER if state & masks.get(name, 0)]
    # Shift alone already changed the character Tk reported, so recording it
    # as well would turn "!" into an unmatchable "shift+!".
    if parts == ["shift"] and len(key) == 1:
        parts = []
    parts.append(key)
    return "+".join(parts)


def describe_binding(binding: str) -> str:
    """Render a binding for display, e.g. ``ctrl+f9`` as ``Ctrl+F9``."""
    parts = []
    for raw in (binding or "").split("+"):
        token = raw.strip().lower()
        if not token:
            continue
        if token in _DISPLAY_NAMES:
            parts.append(_DISPLAY_NAMES[token])
        elif token.startswith("f") and token[1:].isdigit():
            parts.append(token.upper())
        else:
            parts.append(token.capitalize())
    return "+".join(parts) or "—"
