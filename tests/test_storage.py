from __future__ import annotations

import json

import pytest

from eventplayback.core.events import Event, EventType
from eventplayback.core.macro import Macro
from eventplayback.services.storage import MacroFileError, load_macro, save_macro


@pytest.fixture
def macro() -> Macro:
    return Macro(
        name="demo",
        events=[
            Event(EventType.MOUSE_MOVE, 0.0, x=1, y=2),
            Event(EventType.KEY_PRESS, 0.5, key="a", pressed=True),
        ],
    )


def test_save_then_load(tmp_path, macro):
    path = save_macro(macro, tmp_path / "demo.json")
    loaded = load_macro(path)
    assert loaded.name == "demo"
    assert [e.timestamp for e in loaded.events] == [0.0, 0.5]


def test_json_suffix_is_enforced(tmp_path, macro):
    path = save_macro(macro, tmp_path / "demo.txt")
    assert path.name == "demo.json"


def test_missing_directories_are_created(tmp_path, macro):
    path = save_macro(macro, tmp_path / "nested" / "deep" / "demo.json")
    assert path.is_file()


def test_save_leaves_no_temporary_files_behind(tmp_path, macro):
    save_macro(macro, tmp_path / "demo.json")
    assert [p.name for p in tmp_path.iterdir()] == ["demo.json"]


def test_unicode_names_survive_the_round_trip(tmp_path):
    macro = Macro(name="マクロ", events=[Event(EventType.KEY_PRESS, 0.0, key="a")])
    path = save_macro(macro, tmp_path / "m.json")
    assert json.loads(path.read_text(encoding="utf-8"))["name"] == "マクロ"
    assert load_macro(path).name == "マクロ"


def test_loading_a_missing_file_reports_clearly(tmp_path):
    with pytest.raises(MacroFileError, match="File not found"):
        load_macro(tmp_path / "nope.json")


def test_loading_a_directory_is_refused(tmp_path):
    with pytest.raises(MacroFileError, match="File not found"):
        load_macro(tmp_path)


def test_broken_json_reports_a_format_error(tmp_path):
    path = tmp_path / "broken.json"
    path.write_text("{not json", encoding="utf-8")
    with pytest.raises(MacroFileError, match="JSON format error"):
        load_macro(path)


def test_valid_json_with_the_wrong_shape_reports_a_data_error(tmp_path):
    path = tmp_path / "wrong.json"
    path.write_text('{"events": [{"type": "nope", "timestamp": 0}]}', encoding="utf-8")
    with pytest.raises(MacroFileError, match="Data format error"):
        load_macro(path)


def test_oversized_files_are_refused(tmp_path, monkeypatch):
    import eventplayback.services.storage as storage

    monkeypatch.setattr(storage, "MAX_FILE_BYTES", 10)
    path = tmp_path / "big.json"
    path.write_text(json.dumps({"events": []}) + " " * 100, encoding="utf-8")
    with pytest.raises(MacroFileError, match="too large"):
        load_macro(path)


def test_untitled_macros_are_named_after_the_file(tmp_path):
    path = tmp_path / "my-macro.json"
    path.write_text('{"events": []}', encoding="utf-8")
    assert load_macro(path).name == "my-macro"


def test_existing_file_is_replaced_not_appended(tmp_path, macro):
    target = tmp_path / "demo.json"
    save_macro(macro, target)
    save_macro(Macro(name="second", events=[Event(EventType.KEY_PRESS, 0.0, key="b")]), target)
    reloaded = load_macro(target)
    assert reloaded.name == "second"
    assert len(reloaded.events) == 1
