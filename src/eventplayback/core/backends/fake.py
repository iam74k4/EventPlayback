"""In-memory backends used by the test suite and for head-less debugging.

These make recording and playback fully deterministic: :class:`FakeClock` never
sleeps, so a five minute macro is verified instantly and without flakiness.
"""

from __future__ import annotations

import threading
from collections.abc import Callable

from .base import Clock, HotkeyListener, InputHandler, InputSource, InputSynthesizer

__all__ = ["FakeClock", "FakeHotkeyListener", "FakeSource", "FakeSynthesizer"]


class FakeClock(Clock):
    """A virtual clock that advances only when waited on."""

    def __init__(self, start: float = 0.0) -> None:
        self.time = start
        self.waits: list[float] = []

    def now(self) -> float:
        return self.time

    def wait(self, seconds: float, cancel: threading.Event) -> bool:
        if seconds > 0:
            self.waits.append(seconds)
            self.time += seconds
        return cancel.is_set()

    def advance(self, seconds: float) -> None:
        self.time += seconds


class FakeSource(InputSource):
    """An input source driven by the test rather than by the OS."""

    def __init__(self) -> None:
        self.handler: InputHandler | None = None
        self.started = False
        self.start_calls = 0
        self.stop_calls = 0
        self.fail_with: Exception | None = None

    def start(self, handler: InputHandler) -> None:
        self.start_calls += 1
        if self.fail_with is not None:
            raise self.fail_with
        self.handler = handler
        self.started = True

    def stop(self) -> None:
        self.stop_calls += 1
        self.started = False

    # -- test drivers ---------------------------------------------------
    def emit_move(self, x: int, y: int) -> None:
        assert self.handler is not None
        self.handler.on_move(x, y)

    def emit_click(self, x: int, y: int, button: str, pressed: bool) -> None:
        assert self.handler is not None
        self.handler.on_click(x, y, button, pressed)

    def emit_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        assert self.handler is not None
        self.handler.on_scroll(x, y, dx, dy)

    def emit_key(self, name: str, pressed: bool) -> None:
        assert self.handler is not None
        self.handler.on_key(name, pressed)


class FakeSynthesizer(InputSynthesizer):
    """Records every synthesized action instead of performing it."""

    def __init__(self) -> None:
        self.actions: list[tuple] = []
        self.fail_on: set[str] = set()

    def _record(self, *action: object) -> None:
        if action[0] in self.fail_on:
            raise RuntimeError(f"synthesizer failure: {action[0]}")
        self.actions.append(tuple(action))

    def move(self, x: int, y: int) -> None:
        self._record("move", x, y)

    def button(self, name: str, pressed: bool) -> None:
        self._record("button", name, pressed)

    def scroll(self, dx: int, dy: int) -> None:
        self._record("scroll", dx, dy)

    def key(self, name: str, pressed: bool) -> None:
        self._record("key", name, pressed)


class FakeHotkeyListener(HotkeyListener):
    def __init__(self) -> None:
        self.bindings: dict[str, Callable[[], None]] = {}
        self.running = False

    def start(self, bindings: dict[str, Callable[[], None]]) -> None:
        self.bindings = dict(bindings)
        self.running = True

    def stop(self) -> None:
        self.running = False

    def trigger(self, hotkey: str) -> None:
        self.bindings[hotkey]()
