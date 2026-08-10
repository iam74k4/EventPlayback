from __future__ import annotations

import pytest

from eventplayback.core.events import Event, EventType
from eventplayback.core.macro import FORMAT_VERSION, Macro


def make_events(*timestamps: float) -> list[Event]:
    return [Event(EventType.MOUSE_MOVE, ts, x=0, y=0) for ts in timestamps]


def test_duration_of_empty_macro_is_zero():
    assert Macro().duration == 0.0


def test_duration_uses_the_latest_timestamp_not_the_last_element():
    macro = Macro(events=make_events(0.0, 5.0, 2.0))
    assert macro.duration == 5.0


def test_round_trip():
    macro = Macro(name="demo", events=make_events(0.0, 1.0), created_at="2024-01-01T00:00:00")
    restored = Macro.from_dict(macro.to_dict())
    assert restored.name == "demo"
    assert restored.created_at == "2024-01-01T00:00:00"
    assert [e.timestamp for e in restored.events] == [0.0, 1.0]


def test_saved_payload_carries_a_format_version():
    assert Macro().to_dict()["version"] == FORMAT_VERSION


def test_files_without_a_version_are_accepted():
    macro = Macro.from_dict({"name": "legacy", "events": [{"type": "key_press", "timestamp": 0.0}]})
    assert len(macro.events) == 1


def test_newer_format_versions_are_refused():
    with pytest.raises(ValueError, match="newer than supported"):
        Macro.from_dict({"version": FORMAT_VERSION + 1, "events": []})


def test_events_are_sorted_by_timestamp_on_load():
    macro = Macro.from_dict(
        {
            "events": [
                {"type": "mouse_move", "timestamp": 2.0},
                {"type": "mouse_move", "timestamp": 0.5},
            ]
        }
    )
    assert [e.timestamp for e in macro.events] == [0.5, 2.0]


def test_sort_is_stable_so_press_release_pairs_keep_their_order():
    macro = Macro.from_dict(
        {
            "events": [
                {"type": "key_press", "timestamp": 1.0, "key": "a"},
                {"type": "key_release", "timestamp": 1.0, "key": "a"},
            ]
        }
    )
    assert [e.type.value for e in macro.events] == ["key_press", "key_release"]


def test_bad_event_reports_its_index():
    with pytest.raises(ValueError, match=r"event \[1\]"):
        Macro.from_dict({"events": [{"type": "key_press", "timestamp": 0.0}, {"type": "bogus"}]})


@pytest.mark.parametrize("payload", ["nope", {"name": "x"}, {"events": "not-a-list"}])
def test_malformed_documents_are_rejected(payload):
    with pytest.raises(ValueError):
        Macro.from_dict(payload)
