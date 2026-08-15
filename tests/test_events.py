from __future__ import annotations

import pytest

from eventplayback.core.events import Event, EventType


def test_round_trip_keeps_every_populated_field():
    event = Event(EventType.MOUSE_CLICK, 1.5, x=10, y=20, button="left", pressed=True)
    assert Event.from_dict(event.to_dict()) == event


def test_to_dict_omits_unset_fields():
    data = Event(EventType.MOUSE_MOVE, 0.25, x=1, y=2).to_dict()
    assert data == {"type": "mouse_move", "timestamp": 0.25, "x": 1, "y": 2}


@pytest.mark.parametrize(
    "payload, message",
    [
        ({"timestamp": 0.0}, "type"),
        ({"type": "mouse_move"}, "timestamp"),
        ({"type": "nope", "timestamp": 0.0}, "Invalid event type"),
        ({"type": "mouse_move", "timestamp": "soon"}, "timestamp must be a number"),
        ({"type": "mouse_move", "timestamp": 0.0, "x": "left"}, "x must be a number"),
        ({"type": "key_press", "timestamp": 0.0, "key": 5}, "key must be a string"),
        ({"type": "key_press", "timestamp": 0.0, "pressed": "yes"}, "pressed must be a boolean"),
    ],
)
def test_invalid_payloads_are_rejected(payload, message):
    with pytest.raises(ValueError, match=message):
        Event.from_dict(payload)


def test_non_object_event_is_rejected():
    with pytest.raises(ValueError, match="must be an object"):
        Event.from_dict([1, 2, 3])


def test_float_coordinates_are_narrowed_to_int():
    event = Event.from_dict({"type": "mouse_move", "timestamp": 0.0, "x": 10.9, "y": -3.2})
    assert (event.x, event.y) == (10, -3)


def test_shifted_moves_timestamp_without_touching_the_original():
    event = Event(EventType.KEY_PRESS, 2.0, key="a", pressed=True)
    moved = event.shifted(-1.5)
    assert moved.timestamp == 0.5
    assert event.timestamp == 2.0
    assert moved.key == "a"
