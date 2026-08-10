from __future__ import annotations

import json

from eventplayback.services.settings import DEFAULT_HOTKEYS, Settings


def test_defaults_round_trip_through_a_file(tmp_path):
    path = tmp_path / "settings.json"
    assert Settings(loop_count=7, always_on_top=False).save(path)
    loaded = Settings.load(path)
    assert loaded.loop_count == 7
    assert loaded.always_on_top is False
    assert loaded.hotkeys == DEFAULT_HOTKEYS


def test_missing_file_yields_defaults(tmp_path):
    assert Settings.load(tmp_path / "absent.json") == Settings()


def test_corrupt_file_yields_defaults(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text("{{{", encoding="utf-8")
    assert Settings.load(path) == Settings()


def test_unknown_fields_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text(json.dumps({"loop_count": 3, "colour_scheme": "neon"}), encoding="utf-8")
    assert Settings.load(path).loop_count == 3


def test_out_of_range_values_are_clamped():
    assert Settings(loop_count=-5).normalized().loop_count == 0
    assert Settings(loop_count=10**9).normalized().loop_count == 100_000
    assert Settings(countdown_seconds=999).normalized().countdown_seconds == 60
    assert Settings(loop_count="three").normalized().loop_count == 1


def test_unknown_appearance_mode_falls_back_to_dark():
    assert Settings(appearance_mode="neon").normalized().appearance_mode == "dark"
    assert Settings(appearance_mode="light").normalized().appearance_mode == "light"


def test_stale_last_directory_is_dropped(tmp_path):
    assert Settings(last_directory=str(tmp_path)).normalized().last_directory == str(tmp_path)
    assert Settings(last_directory=str(tmp_path / "gone")).normalized().last_directory == ""


def test_partial_hotkeys_are_filled_in_from_the_defaults():
    normalized = Settings(hotkeys={"record": "F7"}).normalized()
    assert normalized.hotkeys["record"] == "f7"
    assert normalized.hotkeys["play"] == DEFAULT_HOTKEYS["play"]


def test_unknown_hotkey_actions_are_discarded():
    assert "explode" not in Settings(hotkeys={"explode": "f1"}).normalized().hotkeys


def test_reserved_keys_cover_every_component_of_a_combination():
    settings = Settings(hotkeys={"record": "ctrl+shift+f9"}).normalized()
    assert {"ctrl", "shift", "f9"} <= settings.reserved_keys


def test_save_creates_missing_directories(tmp_path):
    path = tmp_path / "nested" / "settings.json"
    assert Settings().save(path)
    assert path.is_file()
