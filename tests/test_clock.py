from __future__ import annotations

import math
import threading
import time

import pytest

from eventplayback.core.backends.base import RealClock
from eventplayback.platform_support import high_resolution_timer


def test_high_resolution_timer_is_a_working_context_manager():
    with high_resolution_timer() as fine_grained:
        assert isinstance(fine_grained, bool)


def test_high_resolution_timer_can_be_nested():
    with high_resolution_timer(), high_resolution_timer():
        pass


def test_precision_scope_tightens_the_spin_threshold_and_restores_it():
    clock = RealClock()
    outside = clock.spin_threshold
    with clock.precision_scope():
        assert clock.spin_threshold in (RealClock.DEFAULT_SPIN, RealClock.FALLBACK_SPIN)
    assert clock.spin_threshold == outside


def test_precision_scope_restores_even_when_the_body_raises():
    clock = RealClock()
    outside = clock.spin_threshold
    with pytest.raises(RuntimeError):
        with clock.precision_scope():
            raise RuntimeError("playback exploded")
    assert clock.spin_threshold == outside


def test_a_coarse_tick_is_compensated_by_spinning_longer():
    """Without a fine timer the clock must spin through a whole tick."""
    assert RealClock.FALLBACK_SPIN > RealClock.DEFAULT_SPIN
    assert RealClock.FALLBACK_SPIN >= 0.0156  # one default Windows tick


def test_wait_returns_false_when_it_runs_to_term():
    assert RealClock().wait(0.005, threading.Event()) is False


def test_wait_returns_true_immediately_when_already_cancelled():
    cancel = threading.Event()
    cancel.set()
    started = time.perf_counter()
    assert RealClock().wait(5.0, cancel) is True
    assert time.perf_counter() - started < 0.5


def test_wait_is_interrupted_promptly_by_another_thread():
    cancel = threading.Event()
    threading.Timer(0.05, cancel.set).start()
    started = time.perf_counter()
    assert RealClock().wait(10.0, cancel) is True
    assert time.perf_counter() - started < 2.0


def test_non_positive_waits_return_at_once():
    clock, cancel = RealClock(), threading.Event()
    assert clock.wait(0.0, cancel) is False
    assert clock.wait(-1.0, cancel) is False


def test_wait_does_not_return_early():
    clock, cancel = RealClock(), threading.Event()
    started = time.perf_counter()
    clock.wait(0.02, cancel)
    assert time.perf_counter() - started >= 0.019


def test_now_is_monotonic():
    clock = RealClock()
    first = clock.now()
    clock.wait(0.001, threading.Event())
    assert clock.now() > first


class CoarseTickEvent(threading.Event):
    """An Event whose wait rounds up to a coarse scheduling tick.

    Stands in for the ~15.6 ms default tick on Windows, so the fallback path
    can be verified without a Windows machine.
    """

    TICK = 0.0156

    def wait(self, timeout=None):  # noqa: D102
        if timeout and timeout > 0:
            ticks = math.ceil(timeout / self.TICK)
            time.sleep(ticks * self.TICK)
        return self.is_set()


def _overshoot(spin: float, target: float) -> float:
    clock = RealClock()
    clock._spin = spin
    cancel = CoarseTickEvent()
    started = time.perf_counter()
    clock.wait(target, cancel)
    return time.perf_counter() - started - target


def test_a_coarse_tick_smears_a_20ms_gap_when_the_spin_window_is_small():
    """This is the failure the Windows fallback exists to prevent."""
    assert _overshoot(RealClock.DEFAULT_SPIN, 0.020) > 0.005


def test_the_fallback_spin_window_absorbs_the_coarse_tick():
    assert _overshoot(RealClock.FALLBACK_SPIN, 0.020) < 0.002
