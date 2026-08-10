from __future__ import annotations

import pytest

from eventplayback.ui.state import (
    ACCENTS,
    MAX_LOOP_COUNT,
    AppState,
    build_view,
    format_info,
    parse_loop_count,
)


def view(state, **kwargs):
    defaults = {"has_events": True, "event_count": 3, "duration": 1.25}
    return build_view(state, **{**defaults, **kwargs})


def test_idle_offers_record_and_play_when_a_macro_exists():
    v = view(AppState.IDLE)
    assert v.status == "Idle"
    assert v.record.enabled and v.play.enabled
    assert not v.stop.enabled
    assert v.loop_enabled


def test_play_is_disabled_without_a_macro():
    v = view(AppState.IDLE, has_events=False, event_count=0, duration=0.0)
    assert not v.play.enabled
    assert v.info == "0 events | 0.0s"


@pytest.mark.parametrize(
    "state", [AppState.COUNTDOWN, AppState.RECORDING, AppState.PLAYING]
)
def test_only_stop_is_available_while_busy(state):
    v = view(state)
    assert v.stop.enabled
    assert not v.record.enabled
    assert not v.play.enabled
    assert not v.loop_enabled


def test_countdown_shows_the_remaining_seconds():
    assert view(AppState.COUNTDOWN, countdown=3).status == "3"
    assert view(AppState.COUNTDOWN, countdown=0).status == "Starting"


def test_blink_alternates_between_the_state_accent_and_idle():
    lit = view(AppState.RECORDING, blink_on=True)
    dark = view(AppState.RECORDING, blink_on=False)
    assert lit.accent == ACCENTS["recording"]
    assert dark.accent == ACCENTS["idle"]


def test_idle_never_blinks():
    assert view(AppState.IDLE, blink_on=False).accent == ACCENTS["idle"]


def test_recording_shows_live_progress_instead_of_a_duration():
    assert view(AppState.RECORDING, event_count=12).info == "12 events | Recording..."


def test_info_uses_a_singular_noun_for_one_event():
    assert format_info(event_count=1, duration=0.5, recording=False) == "1 event | 0.5s"


@pytest.mark.parametrize(
    "text, expected",
    [("1", 1), ("", 1), ("  5 ", 5), ("0", 0), ("-3", 0), ("999999", MAX_LOOP_COUNT)],
)
def test_loop_count_parsing(text, expected):
    assert parse_loop_count(text)[0] == expected


@pytest.mark.parametrize("text", ["abc", "1.5", "3x", "--2"])
def test_non_integers_fall_back_to_one_run_with_an_explanation(text):
    count, warning = parse_loop_count(text)
    assert count == 1
    assert warning is not None


def test_valid_input_produces_no_warning():
    assert parse_loop_count("4") == (4, None)


def test_clamped_input_explains_itself():
    assert parse_loop_count(str(MAX_LOOP_COUNT + 1))[1] is not None
