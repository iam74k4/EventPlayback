"""Abstractions the recorder and player are written against.

Keeping these narrow is what lets ``core`` stay importable without pynput, a
display server or a running GUI.
"""

from __future__ import annotations

import threading
import time
from abc import ABC, abstractmethod
from collections.abc import Callable
from dataclasses import dataclass

__all__ = [
    "Clock",
    "HotkeyListener",
    "InputHandler",
    "InputSource",
    "InputSynthesizer",
    "RealClock",
]


@dataclass(frozen=True)
class InputHandler:
    """Callbacks an :class:`InputSource` invokes for normalised input.

    Coordinates are integers, buttons and keys are backend-neutral strings, and
    ``pressed`` is ``True`` for a press and ``False`` for a release.
    """

    on_move: Callable[[int, int], None]
    on_click: Callable[[int, int, str, bool], None]
    on_scroll: Callable[[int, int, int, int], None]
    on_key: Callable[[str, bool], None]


class InputSource(ABC):
    """Captures input events from the operating system."""

    @abstractmethod
    def start(self, handler: InputHandler) -> None:
        """Begin delivering events to ``handler``.

        Raises:
            RuntimeError: if the platform refuses to install the input hooks.
        """

    @abstractmethod
    def stop(self) -> None:
        """Stop delivering events. Must be safe to call when not started."""


class InputSynthesizer(ABC):
    """Injects input events back into the operating system."""

    @abstractmethod
    def move(self, x: int, y: int) -> None: ...

    @abstractmethod
    def button(self, name: str, pressed: bool) -> None: ...

    @abstractmethod
    def scroll(self, dx: int, dy: int) -> None: ...

    @abstractmethod
    def key(self, name: str, pressed: bool) -> None: ...


class HotkeyListener(ABC):
    """Watches for global hotkeys while other applications have focus."""

    @abstractmethod
    def start(self, bindings: dict[str, Callable[[], None]]) -> None:
        """Register ``{"f9": callback}`` style bindings and start listening."""

    @abstractmethod
    def stop(self) -> None: ...


class Clock(ABC):
    """Time source. Swapped for a virtual clock to make playback tests exact."""

    @abstractmethod
    def now(self) -> float:
        """Monotonic time in seconds. Only differences are meaningful."""

    @abstractmethod
    def wait(self, seconds: float, cancel: threading.Event) -> bool:
        """Wait up to ``seconds``, returning early when ``cancel`` is set.

        Returns:
            ``True`` if the wait was cancelled, ``False`` if it ran to term.
        """


class RealClock(Clock):
    """Wall-clock implementation used by the application.

    ``threading.Event.wait`` is subject to the OS timer granularity (~15 ms on
    Windows by default), which is enough to smear a macro that was recorded at
    20 ms resolution. The final few milliseconds are therefore spun out, so
    waits stay both accurate and promptly cancellable.
    """

    #: How long before the deadline to switch from sleeping to spinning.
    SPIN_THRESHOLD = 0.002

    def now(self) -> float:
        return time.perf_counter()

    def wait(self, seconds: float, cancel: threading.Event) -> bool:
        if seconds <= 0:
            return cancel.is_set()
        deadline = time.perf_counter() + seconds
        coarse = seconds - self.SPIN_THRESHOLD
        if coarse > 0 and cancel.wait(coarse):
            return True
        while time.perf_counter() < deadline:
            if cancel.is_set():
                return True
            time.sleep(0)
        return cancel.is_set()
