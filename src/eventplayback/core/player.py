"""Playback, split into a synchronous engine and a thin threaded wrapper.

:class:`PlaybackEngine` holds all of the timing and dispatch logic and never
touches threads itself, so it can be driven straight from a test with a virtual
clock. :class:`Player` adds the background thread the GUI needs.
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable, Sequence

from .backends.base import Clock, InputSynthesizer, RealClock
from .events import Event, EventType

logger = logging.getLogger(__name__)

__all__ = ["PlaybackEngine", "Player"]


class PlaybackEngine:
    """Replays a sequence of events through an :class:`InputSynthesizer`."""

    def __init__(
        self,
        synthesizer: InputSynthesizer,
        clock: Clock | None = None,
        *,
        on_error: Callable[[str], None] | None = None,
    ) -> None:
        self._synth = synthesizer
        self._clock = clock or RealClock()
        self.on_error = on_error
        self._held_keys: set[str] = set()
        self._held_buttons: set[str] = set()

    def run(
        self,
        events: Sequence[Event],
        loops: int,
        cancel: threading.Event,
        *,
        loop_delay: float = 0.0,
    ) -> bool:
        """Replay ``events``.

        Args:
            events: The events to replay, ordered by timestamp.
            loops: How many times to repeat; ``0`` means repeat until cancelled.
            cancel: Set this to abort; waits are interrupted immediately.
            loop_delay: Extra pause inserted between repetitions.

        Returns:
            ``True`` if every requested repetition finished, ``False`` if the
            run was cancelled.
        """
        if not events:
            return not cancel.is_set()

        completed = True
        try:
            iteration = 0
            while not cancel.is_set():
                iteration += 1
                if not self._play_once(events, cancel):
                    completed = False
                    break
                if loops > 0 and iteration >= loops:
                    break
                if loop_delay > 0 and self._clock.wait(loop_delay, cancel):
                    completed = False
                    break
        finally:
            # Whether we finished or were interrupted mid-chord, nothing may be
            # left held down: a stuck Ctrl or mouse button would otherwise
            # affect every application on the desktop.
            self.release_all()
        return completed and not cancel.is_set()

    def _play_once(self, events: Sequence[Event], cancel: threading.Event) -> bool:
        # Anchor on the first event rather than on zero, so a macro whose first
        # action happens two seconds in does not re-insert that gap every loop.
        origin = events[0].timestamp
        base = self._clock.now()
        for event in events:
            if cancel.is_set():
                return False
            due = base + (event.timestamp - origin)
            remaining = due - self._clock.now()
            if remaining > 0 and self._clock.wait(remaining, cancel):
                return False
            self._dispatch(event)
        return not cancel.is_set()

    def _dispatch(self, event: Event) -> None:
        try:
            if event.type is EventType.MOUSE_MOVE:
                if event.x is not None and event.y is not None:
                    self._synth.move(event.x, event.y)
            elif event.type is EventType.MOUSE_CLICK:
                if event.button and event.pressed is not None:
                    if event.x is not None and event.y is not None:
                        self._synth.move(event.x, event.y)
                    self._synth.button(event.button, event.pressed)
                    if event.pressed:
                        self._held_buttons.add(event.button)
                    else:
                        self._held_buttons.discard(event.button)
            elif event.type is EventType.MOUSE_SCROLL:
                self._synth.scroll(event.scroll_dx or 0, event.scroll_dy or 0)
            elif event.type in (EventType.KEY_PRESS, EventType.KEY_RELEASE):
                if event.key:
                    pressed = event.type is EventType.KEY_PRESS
                    self._synth.key(event.key, pressed)
                    if pressed:
                        self._held_keys.add(event.key)
                    else:
                        self._held_keys.discard(event.key)
        except Exception as exc:
            message = f"Playback error: {exc}"
            logger.error(message, exc_info=True)
            if self.on_error is not None:
                self.on_error(message)

    def release_all(self) -> None:
        """Release every key and button this engine is still holding down."""
        for name in sorted(self._held_keys):
            try:
                self._synth.key(name, False)
            except Exception:
                logger.debug("Could not release key %r", name, exc_info=True)
        for name in sorted(self._held_buttons):
            try:
                self._synth.button(name, False)
            except Exception:
                logger.debug("Could not release button %r", name, exc_info=True)
        self._held_keys.clear()
        self._held_buttons.clear()


class Player:
    """Runs a :class:`PlaybackEngine` on a background thread."""

    #: How long ``stop`` waits for the playback thread to unwind.
    JOIN_TIMEOUT = 2.0

    def __init__(
        self,
        synthesizer: InputSynthesizer,
        clock: Clock | None = None,
    ) -> None:
        self._engine = PlaybackEngine(synthesizer, clock, on_error=self._report_error)
        self._events: list[Event] = []
        self._loop_count = 1
        self._loop_delay = 0.0
        self._playing = False
        self._cancel = threading.Event()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self.on_complete: Callable[[], None] | None = None
        self.on_error: Callable[[str], None] | None = None

    def set_events(self, events: Sequence[Event]) -> None:
        with self._lock:
            self._events = list(events)

    def set_loop(self, count: int, *, delay: float = 0.0) -> None:
        """Set the repetition count (``0`` = until stopped) and inter-loop gap."""
        with self._lock:
            self._loop_count = max(0, count)
            self._loop_delay = max(0.0, delay)

    def start(self) -> bool:
        """Begin playback. Returns ``False`` if already playing or empty."""
        with self._lock:
            if self._playing or not self._events:
                return False
            self._playing = True
            self._cancel = threading.Event()
            events, loops, delay, cancel = (
                list(self._events),
                self._loop_count,
                self._loop_delay,
                self._cancel,
            )
        self._thread = threading.Thread(
            target=self._run, args=(events, loops, delay, cancel), name="playback", daemon=True
        )
        self._thread.start()
        return True

    def stop(self) -> None:
        with self._lock:
            if not self._playing:
                return
            self._cancel.set()
            thread = self._thread
        if thread is not None and thread is not threading.current_thread():
            thread.join(timeout=self.JOIN_TIMEOUT)
            if thread.is_alive():
                logger.warning("Playback thread did not stop within %.1fs", self.JOIN_TIMEOUT)
        with self._lock:
            self._playing = False

    def is_playing(self) -> bool:
        with self._lock:
            return self._playing

    def _run(self, events: list[Event], loops: int, delay: float, cancel: threading.Event) -> None:
        completed = False
        try:
            completed = self._engine.run(events, loops, cancel, loop_delay=delay)
        except Exception as exc:
            logger.error("Playback failed: %s", exc, exc_info=True)
            self._report_error(f"Playback failed: {exc}")
        finally:
            with self._lock:
                self._playing = False
                self._thread = None
        if completed and self.on_complete is not None:
            self.on_complete()

    def _report_error(self, message: str) -> None:
        if self.on_error is not None:
            self.on_error(message)
