"""Abstractions the recorder and player are written against.

Keeping these narrow is what lets ``core`` stay importable without pynput, a
display server or a running GUI.
"""

from __future__ import annotations

import threading
import time
from abc import ABC, abstractmethod
from collections.abc import Callable, Iterator
from contextlib import AbstractContextManager, contextmanager, nullcontext
from dataclasses import dataclass

from ...platform_support import high_resolution_timer

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

    def precision_scope(self) -> AbstractContextManager[None]:
        """Hold whatever the platform needs for accurate waits.

        Entered around a playback run. The default does nothing, which is what
        a virtual clock wants.
        """
        return nullcontext()


class RealClock(Clock):
    """Wall-clock implementation used by the application.

    A blocking wait only resolves to the OS scheduling tick, which is too
    coarse for a macro recorded at 20 ms resolution. Two things keep playback
    accurate: :meth:`precision_scope` asks the OS for a 1 ms tick, and the last
    couple of milliseconds before each deadline are spun out rather than slept
    through, so waits stay both accurate and promptly cancellable.
    """

    #: How long before the deadline to switch from sleeping to spinning.
    DEFAULT_SPIN = 0.002
    #: Used when the OS will not grant a fine timer tick. Spinning through a
    #: whole tick costs CPU, but only while a macro is actually being replayed,
    #: and it is the difference between a faithful replay and a smeared one.
    FALLBACK_SPIN = 0.016

    def __init__(self) -> None:
        self._spin = self.DEFAULT_SPIN

    @property
    def spin_threshold(self) -> float:
        return self._spin

    @contextmanager
    def precision_scope(self) -> Iterator[None]:
        with high_resolution_timer() as fine_grained:
            previous = self._spin
            self._spin = self.DEFAULT_SPIN if fine_grained else self.FALLBACK_SPIN
            try:
                yield
            finally:
                self._spin = previous

    def now(self) -> float:
        return time.perf_counter()

    def wait(self, seconds: float, cancel: threading.Event) -> bool:
        if seconds <= 0:
            return cancel.is_set()
        deadline = time.perf_counter() + seconds
        coarse = seconds - self._spin
        if coarse > 0 and cancel.wait(coarse):
            return True
        while time.perf_counter() < deadline:
            if cancel.is_set():
                return True
            time.sleep(0)
        return cancel.is_set()
