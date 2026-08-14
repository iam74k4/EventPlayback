from __future__ import annotations

import pytest

from eventplayback.ui import theme
from eventplayback.ui.state import (
    MAX_LOOP_COUNT,
    STATUS_COLORS,
    AppState,
    build_view,
    elide,
    format_info,
    format_loop_label,
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
    assert v.info == "0 events · 0.0s"


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
    assert view(AppState.COUNTDOWN, countdown=3).status == "Starting in 3"
    assert view(AppState.COUNTDOWN, countdown=0).status == "Starting"


def test_countdown_progress_fills_as_the_wait_runs_out():
    assert view(AppState.COUNTDOWN, countdown=3, countdown_total=3).progress == 0.0
    assert view(AppState.COUNTDOWN, countdown=1, countdown_total=4).progress == 0.75


def test_pulse_fades_the_indicator_without_losing_its_hue():
    lit = view(AppState.RECORDING, pulse=0.0)
    faded = view(AppState.RECORDING, pulse=0.5)
    assert lit.dot == STATUS_COLORS[AppState.RECORDING]
    assert faded.dot != lit.dot


def test_idle_has_no_progress_bar_and_no_pulse():
    v = view(AppState.IDLE, pulse=0.5)
    assert v.progress is None
    assert v.dot == theme.STATUS_IDLE


def test_recording_shows_live_progress_instead_of_a_duration():
    assert view(AppState.RECORDING, event_count=12).info == "12 events · recording"


def test_info_uses_a_singular_noun_for_one_event():
    assert format_info(event_count=1, duration=0.5, recording=False) == "1 event · 0.5s"


def test_detail_is_appended_to_the_info_line():
    assert view(AppState.PLAYING, detail="loop 2/5").info.endswith("· loop 2/5")


@pytest.mark.parametrize(
    "current, total, expected",
    [(1, 3, "loop 1/3"), (2, 0, "loop 2/∞"), (4, -1, "loop 4/∞")],
)
def test_loop_label(current, total, expected):
    assert format_loop_label(current, total) == expected


@pytest.mark.parametrize(
    "text, expected",
    [("1", 1), ("", 1), ("  5 ", 5), ("0", 0), ("∞", 0), ("-3", 0), ("999999", MAX_LOOP_COUNT)],
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


@pytest.mark.parametrize(
    "text, limit, expected",
    [
        ("short", 10, "short"),
        ("exactly-10", 10, "exactly-10"),
        ("eleven-char", 10, "eleven-ch…"),
        ("", 5, ""),
        ("abc", 1, "a…"),
    ],
)
def test_elide_keeps_the_limit(text, limit, expected):
    assert elide(text, limit) == expected
    assert len(elide(text, limit)) <= max(limit, 2)
