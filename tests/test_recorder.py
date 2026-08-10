from __future__ import annotations

import pytest

from eventplayback.core.backends.fake import FakeClock, FakeSource
from eventplayback.core.events import EventType
from eventplayback.core.recorder import Recorder


@pytest.fixture
def setup():
    source = FakeSource()
    clock = FakeClock()
    return source, clock, Recorder(source, clock, move_interval=0.02)


def test_nothing_is_captured_before_start(setup):
    source, _clock, recorder = setup
    assert recorder.event_count() == 0
    assert not recorder.is_recording()
    assert source.handler is None


def test_records_clicks_scrolls_and_keys(setup):
    source, clock, recorder = setup
    recorder.start()
    clock.advance(0.5)
    source.emit_click(3, 4, "left", True)
    clock.advance(0.1)
    source.emit_scroll(3, 4, 0, -2)
    clock.advance(0.1)
    source.emit_key("a", True)
    events = recorder.stop()

    assert [e.type for e in events] == [
        EventType.MOUSE_CLICK,
        EventType.MOUSE_SCROLL,
        EventType.KEY_PRESS,
    ]
    assert events[0].timestamp == pytest.approx(0.5)
    assert events[1].scroll_dy == -2
    assert events[2].key == "a"


def test_first_move_is_kept_and_later_ones_are_throttled(setup):
    source, clock, recorder = setup
    recorder.start()
    source.emit_move(1, 1)
    clock.advance(0.005)
    source.emit_move(2, 2)  # dropped: inside the throttle window
    clock.advance(0.05)
    source.emit_move(3, 3)
    events = recorder.stop()

    assert [(e.x, e.y) for e in events] == [(1, 1), (3, 3)]


def test_reserved_hotkeys_are_never_recorded():
    source, clock = FakeSource(), FakeClock()
    recorder = Recorder(source, clock, excluded_keys={"f9", "f10", "esc"})
    recorder.start()
    source.emit_key("F9", True)  # matching is case insensitive
    source.emit_key("esc", True)
    source.emit_key("b", True)
    assert [e.key for e in recorder.stop()] == ["b"]


def test_excluded_keys_can_be_changed_between_runs(setup):
    source, _clock, recorder = setup
    recorder.set_excluded_keys({"f7"})
    recorder.start()
    source.emit_key("f7", True)
    source.emit_key("f9", True)
    assert [e.key for e in recorder.stop()] == ["f9"]


def test_timestamps_restart_from_zero_on_each_run(setup):
    source, clock, recorder = setup
    recorder.start()
    clock.advance(1.0)
    source.emit_key("a", True)
    recorder.stop()

    clock.advance(30.0)
    recorder.start()
    clock.advance(0.25)
    source.emit_key("b", True)
    events = recorder.stop()

    assert len(events) == 1
    assert events[0].timestamp == pytest.approx(0.25)


def test_events_arriving_after_stop_are_ignored(setup):
    source, _clock, recorder = setup
    recorder.start()
    source.emit_key("a", True)
    recorder.stop()
    source.emit_key("b", True)  # a listener thread still winding down
    assert [e.key for e in recorder.events()] == ["a"]


def test_start_is_idempotent(setup):
    source, _clock, recorder = setup
    recorder.start()
    recorder.start()
    assert source.start_calls == 1


def test_stop_without_start_is_safe(setup):
    source, _clock, recorder = setup
    assert recorder.stop() == []
    assert source.stop_calls == 0


def test_failure_to_hook_input_leaves_the_recorder_idle(setup):
    source, _clock, recorder = setup
    source.fail_with = RuntimeError("permission denied")
    with pytest.raises(RuntimeError, match="permission denied"):
        recorder.start()
    assert not recorder.is_recording()


def test_on_event_callback_fires_and_cannot_break_recording(setup):
    source, _clock, recorder = setup
    seen = []

    def callback(event):
        seen.append(event)
        raise ValueError("UI blew up")

    recorder.on_event = callback
    recorder.start()
    source.emit_key("a", True)
    source.emit_key("a", False)
    assert len(seen) == 2
    assert recorder.event_count() == 2
