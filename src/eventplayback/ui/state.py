"""How application state maps onto what the window shows.

Deliberately free of any toolkit import: ``build_view`` is a pure function, so
the state/appearance rules can be unit tested without opening a window.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

__all__ = [
    "ACCENTS",
    "MAX_LOOP_COUNT",
    "AppState",
    "ButtonView",
    "ViewModel",
    "build_view",
    "format_info",
    "parse_loop_count",
]

#: Highest repetition count accepted from the loop field. ``0`` means infinite.
MAX_LOOP_COUNT = 10_000

_IDLE = "#2b2b2b"
_DISABLED = "#7f8c8d"
_RECORD = "#c0392b"
_PLAY = "#2980b9"
_STOP = "#e74c3c"

#: Window background per state, used for the blinking status indication.
ACCENTS = {
    "idle": _IDLE,
    "countdown": "#f39c12",
    "recording": "#e74c3c",
    "playing": "#27ae60",
}


class AppState(Enum):
    IDLE = "idle"
    COUNTDOWN = "countdown"
    RECORDING = "recording"
    PLAYING = "playing"


@dataclass(frozen=True)
class ButtonView:
    enabled: bool
    color: str

    @property
    def tk_state(self) -> str:
        return "normal" if self.enabled else "disabled"


@dataclass(frozen=True)
class ViewModel:
    status: str
    info: str
    accent: str
    record: ButtonView
    stop: ButtonView
    play: ButtonView
    loop_enabled: bool


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
        return f"{event_count} {plural} | Recording..."
    return f"{event_count} {plural} | {duration:.1f}s"


def build_view(
    state: AppState,
    *,
    has_events: bool,
    event_count: int,
    duration: float,
    countdown: int = 0,
    blink_on: bool = True,
) -> ViewModel:
    """Derive every visible property of the window from the current state."""
    recording = state is AppState.RECORDING
    info = format_info(event_count=event_count, duration=duration, recording=recording)

    if state is AppState.IDLE:
        return ViewModel(
            status="Idle",
            info=info,
            accent=ACCENTS["idle"],
            record=ButtonView(True, _RECORD),
            stop=ButtonView(False, _DISABLED),
            play=ButtonView(has_events, _PLAY if has_events else _DISABLED),
            loop_enabled=True,
        )

    if state is AppState.COUNTDOWN:
        status = str(countdown) if countdown > 0 else "Starting"
    elif state is AppState.RECORDING:
        status = "● Recording"
    else:
        status = "▶ Playing"

    return ViewModel(
        status=status,
        info=info,
        accent=ACCENTS[state.value] if blink_on else ACCENTS["idle"],
        record=ButtonView(False, _DISABLED),
        stop=ButtonView(True, _STOP),
        play=ButtonView(False, _DISABLED),
        loop_enabled=False,
    )
