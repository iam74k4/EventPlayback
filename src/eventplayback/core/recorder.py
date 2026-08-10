"""Turns raw input from an :class:`InputSource` into a list of events."""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable, Iterable

from .backends.base import Clock, InputHandler, InputSource, RealClock
from .events import Event, EventType

logger = logging.getLogger(__name__)

__all__ = ["Recorder"]


class Recorder:
    """Record mouse and keyboard input.

    The source and clock are injected so the same class drives the real
    application and the test suite::

        recorder = Recorder(FakeSource(), FakeClock())
        recorder.start()
        ...
        events = recorder.stop()
    """

    #: Mouse moves are throttled to this interval (seconds). Without it a fast
    #: drag produces thousands of near-identical samples.
    DEFAULT_MOVE_INTERVAL = 0.02

    def __init__(
        self,
        source: InputSource,
        clock: Clock | None = None,
        *,
        move_interval: float = DEFAULT_MOVE_INTERVAL,
        excluded_keys: Iterable[str] = (),
    ) -> None:
        self._source = source
        self._clock = clock or RealClock()
        self._move_interval = max(0.0, move_interval)
        self._excluded_keys = {key.lower() for key in excluded_keys}

        self._events: list[Event] = []
        self._start_time = 0.0
        self._last_move = 0.0
        self._recording = False
        self._lock = threading.Lock()
        self.on_event: Callable[[Event], None] | None = None

    # -- lifecycle ------------------------------------------------------
    def start(self) -> None:
        """Start capturing.

        Raises:
            RuntimeError: if the backend cannot install its input hooks.
        """
        with self._lock:
            if self._recording:
                return
            self._events = []
            self._start_time = self._clock.now()
            self._last_move = -self._move_interval
            self._recording = True

        handler = InputHandler(
            on_move=self._on_move,
            on_click=self._on_click,
            on_scroll=self._on_scroll,
            on_key=self._on_key,
        )
        try:
            self._source.start(handler)
        except Exception:
            with self._lock:
                self._recording = False
            raise

    def stop(self) -> list[Event]:
        """Stop capturing and return everything recorded."""
        with self._lock:
            was_recording = self._recording
            self._recording = False
        if was_recording:
            self._source.stop()
        return self.events()

    def is_recording(self) -> bool:
        with self._lock:
            return self._recording

    def events(self) -> list[Event]:
        with self._lock:
            return list(self._events)

    def event_count(self) -> int:
        with self._lock:
            return len(self._events)

    def set_excluded_keys(self, keys: Iterable[str]) -> None:
        """Replace the set of key names that must never be recorded.

        Used to keep the application's own hotkeys out of the macro.
        """
        with self._lock:
            self._excluded_keys = {key.lower() for key in keys}

    # -- capture --------------------------------------------------------
    def _timestamp(self) -> float:
        return self._clock.now() - self._start_time

    def _append(self, event: Event) -> None:
        with self._lock:
            if not self._recording:
                return
            self._events.append(event)
        callback = self.on_event
        if callback is not None:
            try:
                callback(event)
            except Exception:
                logger.debug("on_event callback failed", exc_info=True)

    def _on_move(self, x: int, y: int) -> None:
        with self._lock:
            if not self._recording:
                return
            now = self._clock.now() - self._start_time
            if now - self._last_move < self._move_interval:
                return
            self._last_move = now
        self._append(Event(EventType.MOUSE_MOVE, now, x=x, y=y))

    def _on_click(self, x: int, y: int, button: str, pressed: bool) -> None:
        if not self._recording:
            return
        self._append(
            Event(EventType.MOUSE_CLICK, self._timestamp(), x=x, y=y, button=button, pressed=pressed)
        )

    def _on_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self._recording:
            return
        self._append(
            Event(EventType.MOUSE_SCROLL, self._timestamp(), x=x, y=y, scroll_dx=dx, scroll_dy=dy)
        )

    def _on_key(self, name: str, pressed: bool) -> None:
        if not self._recording or name.lower() in self._excluded_keys:
            return
        event_type = EventType.KEY_PRESS if pressed else EventType.KEY_RELEASE
        self._append(Event(event_type, self._timestamp(), key=name, pressed=pressed))
