"""How application state maps onto what the window shows.

Deliberately free of any toolkit import: ``build_view`` is a pure function, so
the state/appearance rules can be unit tested without opening a window. The
colours themselves live in :mod:`eventplayback.ui.theme`.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from . import theme
from .theme import Color

__all__ = [
    "MAX_LOOP_COUNT",
    "STATUS_COLORS",
    "AppState",
    "ButtonView",
    "ViewModel",
    "build_view",
    "elide",
    "format_info",
    "format_loop_label",
    "parse_loop_count",
]

#: Highest repetition count accepted from the loop field. ``0`` means infinite.
MAX_LOOP_COUNT = 10_000


class AppState(Enum):
    IDLE = "idle"
    COUNTDOWN = "countdown"
    RECORDING = "recording"
    PLAYING = "playing"


#: The indicator colour for each state. Replaces the old full-window flash:
#: a small dot carries the same information without strobing the whole UI.
STATUS_COLORS: dict[AppState, Color] = {
    AppState.IDLE: theme.STATUS_IDLE,
    AppState.COUNTDOWN: theme.STATUS_COUNTDOWN,
    AppState.RECORDING: theme.STATUS_RECORDING,
    AppState.PLAYING: theme.STATUS_PLAYING,
}


@dataclass(frozen=True)
class ButtonView:
    enabled: bool
    color: Color
    hover: Color

    @property
    def tk_state(self) -> str:
        return "normal" if self.enabled else "disabled"


@dataclass(frozen=True)
class ViewModel:
    status: str
    info: str
    #: Colour of the state indicator, already faded if this is a dim pulse.
    dot: Color
    record: ButtonView
    stop: ButtonView
    play: ButtonView
    loop_enabled: bool
    #: ``0.0``-``1.0`` while there is measurable progress, ``None`` when the
    #: bar should be hidden because nothing quantifiable is happening.
    progress: float | None = None
    progress_color: Color = theme.PLAY


def elide(text: str, limit: int) -> str:
    """Shorten ``text`` to ``limit`` characters, ending with an ellipsis.

    Used wherever a name arrives from outside the application. A file name has
    no length the layout can count on, and wrapping cannot rescue a long run of
    characters with nowhere to break in it.
    """
    if len(text) <= limit:
        return text
    return f"{text[: max(1, limit - 1)]}…"


def parse_loop_count(text: str) -> tuple[int, str | None]:
    """Interpret the loop field.

    Returns the repetition count (``0`` meaning infinite) and, when the input
    had to be corrected, a message explaining what was used instead. The old
    behaviour silently fell back to a single run, which looked like the button
    had simply ignored the request.
    """
    raw = (text or "").strip()
    if not raw:
        return 1, None
    if raw == "∞":
        return 0, None
    try:
        value = int(raw)
    except ValueError:
        return 1, f"'{raw}' is not a number, playing once"
    if value < 0:
        return 0, "Negative loop count treated as infinite"
    if value > MAX_LOOP_COUNT:
        return MAX_LOOP_COUNT, f"Loop count capped at {MAX_LOOP_COUNT}"
    return value, None


def format_info(*, event_count: int, duration: float, recording: bool) -> str:
    plural = "event" if event_count == 1 else "events"
    if recording:
        return f"{event_count} {plural} · recording"
    return f"{event_count} {plural} · {duration:.1f}s"


def format_loop_label(current: int, total: int) -> str:
    """Describe which repetition is running, e.g. ``loop 2/5`` or ``loop 2/∞``."""
    return f"loop {current}/{'∞' if total <= 0 else total}"


def build_view(
    state: AppState,
    *,
    has_events: bool,
    event_count: int,
    duration: float,
    countdown: int = 0,
    countdown_total: int = 0,
    pulse: float = 0.0,
    progress: float | None = None,
    detail: str = "",
) -> ViewModel:
    """Derive every visible property of the window from the current state."""
    recording = state is AppState.RECORDING
    info = format_info(event_count=event_count, duration=duration, recording=recording)
    if detail:
        info = f"{info} · {detail}"

    if state is AppState.IDLE:
        return ViewModel(
            status="Idle",
            info=info,
            dot=theme.STATUS_IDLE,
            record=ButtonView(True, theme.RECORD, theme.RECORD_HOVER),
            stop=ButtonView(False, theme.DISABLED, theme.DISABLED),
            play=ButtonView(
                has_events,
                theme.PLAY if has_events else theme.DISABLED,
                theme.PLAY_HOVER if has_events else theme.DISABLED,
            ),
            loop_enabled=True,
            progress=None,
        )

    if state is AppState.COUNTDOWN:
        status = f"Starting in {countdown}" if countdown > 0 else "Starting"
        # Fills up as the wait runs out, so the bar reads as "almost there".
        if progress is None and countdown_total > 0:
            progress = max(0.0, min(1.0, (countdown_total - countdown) / countdown_total))
    elif recording:
        status = "Recording"
    else:
        status = "Playing"

    color = STATUS_COLORS[state]
    return ViewModel(
        status=status,
        info=info,
        dot=color if pulse <= 0 else theme.dim(color, pulse),
        record=ButtonView(False, theme.DISABLED, theme.DISABLED),
        stop=ButtonView(True, theme.STOP, theme.STOP_HOVER),
        play=ButtonView(False, theme.DISABLED, theme.DISABLED),
        loop_enabled=False,
        progress=progress,
        progress_color=color,
    )
