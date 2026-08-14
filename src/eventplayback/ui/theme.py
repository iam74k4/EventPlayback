"""Design tokens for the window.

Every colour is a ``(light, dark)`` pair, which is exactly what CustomTkinter
accepts for ``fg_color`` and its relatives. Handing widgets the pair rather
than a single hex string is what makes the light appearance mode work: the
toolkit picks the right half, so nothing has to repaint itself when the mode
changes.

Colours are named for their role rather than their hue, so a palette change
stays in this file. Deliberately free of any toolkit import, which is what
lets :mod:`eventplayback.ui.state` remain unit testable without a display.
"""

from __future__ import annotations

__all__ = [
    "BANNER_TEXT",
    "BORDER",
    "BORDER_FOCUS",
    "DANGER",
    "DISABLED",
    "FONT_SIZE_LG",
    "FONT_SIZE_MD",
    "FONT_SIZE_SM",
    "FONT_SIZE_XL",
    "INFO",
    "NEUTRAL",
    "NEUTRAL_HOVER",
    "PLAY",
    "PLAY_HOVER",
    "RADIUS_LG",
    "RADIUS_MD",
    "RADIUS_SM",
    "RECORD",
    "RECORD_HOVER",
    "STATUS_COUNTDOWN",
    "STATUS_IDLE",
    "STATUS_PLAYING",
    "STATUS_RECORDING",
    "STOP",
    "STOP_HOVER",
    "SUCCESS",
    "SURFACE",
    "SURFACE_SUNKEN",
    "TEXT",
    "TEXT_FAINT",
    "TEXT_MUTED",
    "TEXT_ON_ACCENT",
    "WARNING",
    "WINDOW",
    "Color",
    "blend",
    "dim",
    "tint",
]

#: A ``(light, dark)`` colour pair, in the form CustomTkinter expects.
Color = tuple[str, str]

# -- surfaces ------------------------------------------------------------
WINDOW: Color = ("#f2f3f5", "#1b1c1f")
SURFACE: Color = ("#ffffff", "#242629")
SURFACE_SUNKEN: Color = ("#e8eaee", "#17181a")
BORDER: Color = ("#d4d7dd", "#34373d")
BORDER_FOCUS: Color = ("#2f6fed", "#4c8dff")

# -- text ----------------------------------------------------------------
TEXT: Color = ("#1c1e21", "#e9eaec")
TEXT_MUTED: Color = ("#5d646d", "#9ba1a9")
TEXT_FAINT: Color = ("#8a919a", "#6d737a")
#: For text sitting on a filled accent button, where the mode must not flip it.
TEXT_ON_ACCENT: Color = ("#ffffff", "#ffffff")

# -- action colours ------------------------------------------------------
RECORD: Color = ("#d63a3f", "#e2494e")
RECORD_HOVER: Color = ("#bf3035", "#ea5f63")
PLAY: Color = ("#2f6fed", "#3b82f6")
PLAY_HOVER: Color = ("#2560d4", "#5596ff")
#: Stop is deliberately quiet: it is the way out of a state, not a warning.
STOP: Color = ("#5b636e", "#3d434a")
STOP_HOVER: Color = ("#4c545e", "#4a515a")
NEUTRAL: Color = ("#e4e6ea", "#2c2f34")
NEUTRAL_HOVER: Color = ("#d6d9df", "#363a40")
DISABLED: Color = ("#e6e8ec", "#25272b")

# -- semantic ------------------------------------------------------------
INFO: Color = ("#2f6fed", "#4c8dff")
SUCCESS: Color = ("#1a9d54", "#3dd68c")
WARNING: Color = ("#c67c0a", "#f5a524")
DANGER: Color = ("#d63a3f", "#e5484d")

# -- state indicator -----------------------------------------------------
STATUS_IDLE: Color = ("#9aa1aa", "#6d737a")
STATUS_COUNTDOWN: Color = WARNING
STATUS_RECORDING: Color = ("#e5484d", "#e5484d")
STATUS_PLAYING: Color = SUCCESS

# -- metrics -------------------------------------------------------------
RADIUS_SM = 6
RADIUS_MD = 8
RADIUS_LG = 10

FONT_SIZE_SM = 11
FONT_SIZE_MD = 12
FONT_SIZE_LG = 13
FONT_SIZE_XL = 15


# -- colour maths --------------------------------------------------------
def _channels(color: str) -> tuple[int, int, int]:
    value = color.lstrip("#")
    if len(value) == 3:
        value = "".join(char * 2 for char in value)
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def _blend_hex(color: str, towards: str, amount: float) -> str:
    amount = max(0.0, min(1.0, amount))
    source, target = _channels(color), _channels(towards)
    mixed = tuple(
        round(start + (end - start) * amount) for start, end in zip(source, target, strict=True)
    )
    return "#{:02x}{:02x}{:02x}".format(*mixed)


def blend(color: Color, towards: Color, amount: float) -> Color:
    """Mix ``color`` towards ``towards``, per appearance mode.

    ``amount`` is how far to travel: ``0.0`` returns ``color`` unchanged and
    ``1.0`` returns ``towards``.
    """
    return (
        _blend_hex(color[0], towards[0], amount),
        _blend_hex(color[1], towards[1], amount),
    )


def dim(color: Color, amount: float = 0.55) -> Color:
    """A faded version of ``color``, for the calm half of a pulse.

    Fading towards the window rather than dropping to a neutral grey keeps the
    hue recognisable, so the indicator still reads as "recording" at its
    dimmest point.
    """
    return blend(color, WINDOW, amount)


def tint(color: Color, amount: float = 0.85) -> Color:
    """A wash of ``color`` over the window, for banner backgrounds."""
    return blend(color, WINDOW, amount)


#: Readable text colour for each banner severity, darkened just enough to sit
#: on that severity's tinted background in light mode.
BANNER_TEXT: dict[str, Color] = {
    "info": blend(INFO, TEXT, 0.25),
    "success": blend(SUCCESS, TEXT, 0.3),
    "warning": blend(WARNING, TEXT, 0.25),
    "error": blend(DANGER, TEXT, 0.2),
}
