from __future__ import annotations

import threading

import pytest

from eventplayback.core.backends.fake import FakeClock, FakeSynthesizer
from eventplayback.core.events import Event, EventType
from eventplayback.core.player import PlaybackEngine, Player


@pytest.fixture
def engine_setup():
    synth = FakeSynthesizer()
    clock = FakeClock()
    return synth, clock, PlaybackEngine(synth, clock)


def key(ts: float, name: str, pressed: bool) -> Event:
    return Event(
        EventType.KEY_PRESS if pressed else EventType.KEY_RELEASE, ts, key=name, pressed=pressed
    )


def test_events_are_replayed_in_order(engine_setup):
    synth, _clock, engine = engine_setup
    events = [
        Event(EventType.MOUSE_MOVE, 0.0, x=1, y=2),
        Event(EventType.MOUSE_SCROLL, 0.1, x=1, y=2, scroll_dx=0, scroll_dy=-3),
        key(0.2, "a", True),
        key(0.3, "a", False),
    ]
    assert engine.run(events, 1, threading.Event())
    assert synth.actions == [
        ("move", 1, 2),
        ("scroll", 0, -3),
        ("key", "a", True),
        ("key", "a", False),
    ]


def test_click_moves_the_cursor_to_the_recorded_position(engine_setup):
    synth, _clock, engine = engine_setup
    events = [Event(EventType.MOUSE_CLICK, 0.0, x=7, y=9, button="left", pressed=True)]
    engine.run(events, 1, threading.Event())
    assert synth.actions[0] == ("move", 7, 9)
    assert synth.actions[1] == ("button", "left", True)


def test_waits_match_the_gaps_between_events(engine_setup):
    _synth, clock, engine = engine_setup
    events = [key(0.0, "a", True), key(0.5, "b", True), key(2.0, "c", True)]
    engine.run(events, 1, threading.Event())
    assert clock.waits == pytest.approx([0.5, 1.5])


def test_leading_silence_is_not_replayed_on_every_loop(engine_setup):
    """A macro whose first event is at t=2s must not pause 2s before each loop."""
    _synth, clock, engine = engine_setup
    events = [key(2.0, "a", True), key(2.5, "b", True)]
    engine.run(events, 3, threading.Event())
    assert clock.waits == pytest.approx([0.5, 0.5, 0.5])


def test_loop_delay_is_inserted_between_repetitions(engine_setup):
    _synth, clock, engine = engine_setup
    events = [key(0.0, "a", True), key(0.25, "a", False)]
    engine.run(events, 3, threading.Event(), loop_delay=1.0)
    assert clock.waits == pytest.approx([0.25, 1.0, 0.25, 1.0, 0.25])


def test_loop_count_repeats_exactly(engine_setup):
    synth, _clock, engine = engine_setup
    engine.run([key(0.0, "a", True), key(0.0, "a", False)], 3, threading.Event())
    assert len(synth.actions) == 6


def test_zero_loop_count_runs_until_cancelled(engine_setup):
    synth, _clock, engine = engine_setup
    cancel = threading.Event()

    class StopAfter(FakeSynthesizer):
        def key(self, name, pressed):
            super().key(name, pressed)
            if len(self.actions) >= 5:
                cancel.set()

    synth = StopAfter()
    engine = PlaybackEngine(synth, FakeClock())
    assert engine.run([key(0.0, "a", True)], 0, cancel) is False
    presses = [a for a in synth.actions if a == ("key", "a", True)]
    assert len(presses) == 5
    assert synth.actions[-1] == ("key", "a", False)  # released on the way out


def test_cancelling_releases_keys_that_were_still_held(engine_setup):
    synth, _clock, engine = engine_setup
    cancel = threading.Event()

    events = [key(0.0, "ctrl", True), key(0.1, "c", True), key(5.0, "c", False)]

    class CancellingClock(FakeClock):
        def wait(self, seconds, cancel_event):
            if seconds > 1.0:  # about to wait for the release
                cancel_event.set()
            return super().wait(seconds, cancel_event)

    engine = PlaybackEngine(synth, CancellingClock())
    assert engine.run(events, 1, cancel) is False
    # ctrl and c were pressed but never released by the macro itself.
    assert ("key", "ctrl", False) in synth.actions
    assert ("key", "c", False) in synth.actions


def test_buttons_still_down_are_released_too(engine_setup):
    synth, _clock, engine = engine_setup
    events = [Event(EventType.MOUSE_CLICK, 0.0, x=0, y=0, button="left", pressed=True)]
    engine.run(events, 1, threading.Event())
    assert synth.actions[-1] == ("button", "left", False)


def test_keys_released_by_the_macro_are_not_released_twice(engine_setup):
    synth, _clock, engine = engine_setup
    engine.run([key(0.0, "a", True), key(0.1, "a", False)], 1, threading.Event())
    assert synth.actions.count(("key", "a", False)) == 1


def test_a_failing_event_is_reported_and_playback_continues():
    synth = FakeSynthesizer()
    synth.fail_on = {"scroll"}
    errors: list[str] = []
    engine = PlaybackEngine(synth, FakeClock(), on_error=errors.append)

    events = [
        Event(EventType.MOUSE_SCROLL, 0.0, x=0, y=0, scroll_dy=1),
        key(0.1, "a", True),
    ]
    assert engine.run(events, 1, threading.Event())
    assert len(errors) == 1
    assert ("key", "a", True) in synth.actions


def test_empty_event_list_completes_immediately(engine_setup):
    _synth, clock, engine = engine_setup
    assert engine.run([], 5, threading.Event())
    assert clock.waits == []


def test_already_cancelled_run_does_nothing(engine_setup):
    synth, _clock, engine = engine_setup
    cancel = threading.Event()
    cancel.set()
    assert engine.run([key(0.0, "a", True)], 1, cancel) is False
    assert synth.actions == []


# -- threaded wrapper ---------------------------------------------------


def test_player_reports_completion():
    synth = FakeSynthesizer()
    player = Player(synth, FakeClock())
    done = threading.Event()
    player.on_complete = done.set
    player.set_events([key(0.0, "a", True), key(0.1, "a", False)])
    player.set_loop(2)

    assert player.start()
    assert done.wait(timeout=5)
    assert not player.is_playing()
    assert len(synth.actions) == 4


def test_player_refuses_to_start_without_events():
    player = Player(FakeSynthesizer(), FakeClock())
    assert player.start() is False


def test_player_will_not_start_twice():
    player = Player(FakeSynthesizer(), FakeClock())
    player.set_events([key(0.0, "a", True)])
    player.set_loop(0)  # runs forever
    assert player.start()
    try:
        assert player.start() is False
    finally:
        player.stop()


def test_stopping_an_infinite_run_terminates_the_thread():
    player = Player(FakeSynthesizer(), FakeClock())
    player.set_events([key(0.0, "a", True), key(0.05, "a", False)])
    player.set_loop(0)
    player.start()
    player.stop()
    assert not player.is_playing()


def test_stop_before_start_is_safe():
    player = Player(FakeSynthesizer(), FakeClock())
    player.stop()
    assert not player.is_playing()


def test_progress_reports_the_position_within_a_repetition(engine_setup):
    synth, clock, engine = engine_setup
    seen: list[tuple[int, float]] = []
    engine.on_progress = lambda loop, elapsed: seen.append((loop, elapsed))
    engine.run([key(0.0, "a", True), key(0.5, "a", False)], 2, threading.Event())
    assert seen == [(1, 0.0), (1, 0.5), (2, 0.0), (2, 0.5)]


def test_a_failing_progress_callback_does_not_stop_playback(engine_setup):
    synth, _clock, engine = engine_setup

    def explode(_loop: int, _elapsed: float) -> None:
        raise RuntimeError("display gone")

    engine.on_progress = explode
    assert engine.run([key(0.0, "a", True), key(0.1, "a", False)], 1, threading.Event())
    assert len(synth.actions) == 2


def test_player_exposes_a_progress_snapshot():
    player = Player(FakeSynthesizer(), FakeClock())
    assert player.progress().loop == 0

    player.set_events([key(0.0, "a", True), key(2.0, "a", False)])
    player.set_loop(3)
    done = threading.Event()
    player.on_complete = done.set
    assert player.start()
    assert done.wait(timeout=5)

    final = player.progress()
    assert final.duration == 2.0
    assert final.loops == 3
    assert final.loop == 3
    assert final.fraction == 1.0


def test_progress_fraction_is_zero_for_an_instant_macro():
    player = Player(FakeSynthesizer(), FakeClock())
    player.set_events([key(1.5, "a", True)])
    assert player.progress().fraction == 0.0
    player.set_events([])
    assert player.progress().fraction == 0.0
