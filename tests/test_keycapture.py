from __future__ import annotations

import pytest

from eventplayback.ui.keycapture import binding_from_event, describe_binding

CTRL = 0x0004
SHIFT = 0x0001
ALT_X11 = 0x0008
ALT_WIN = 0x20000


def test_a_bare_function_key():
    assert binding_from_event("F9", 0) == "f9"


def test_tk_names_are_translated_to_listener_names():
    assert binding_from_event("Escape", 0) == "esc"
    assert binding_from_event("Return", 0) == "enter"
    assert binding_from_event("Prior", 0) == "page_up"


def test_modifiers_are_prefixed_in_a_fixed_order():
    assert binding_from_event("F9", CTRL | SHIFT) == "ctrl+shift+f9"
    assert binding_from_event("F9", SHIFT | CTRL) == "ctrl+shift+f9"


def test_alt_is_read_from_the_right_bit_per_platform():
    assert binding_from_event("F9", ALT_X11, platform="linux") == "alt+f9"
    assert binding_from_event("F9", ALT_WIN, platform="win32") == "alt+f9"
    # The X11 bit means Command on macOS, not Alt.
    assert binding_from_event("F9", 0x0008, platform="darwin") == "cmd+f9"


def test_an_unknown_platform_falls_back_rather_than_failing():
    assert binding_from_event("F9", CTRL, platform="plan9") == "ctrl+f9"


@pytest.mark.parametrize("keysym", ["Shift_L", "Control_R", "Alt_L", "Caps_Lock", "Super_L"])
def test_holding_only_a_modifier_is_not_a_binding(keysym):
    assert binding_from_event(keysym, CTRL) is None


def test_punctuation_keysyms_become_the_character_itself():
    assert binding_from_event("exclam", SHIFT) == "!"
    assert binding_from_event("comma", 0) == ","


def test_shift_is_dropped_for_a_character_it_already_changed():
    # Tk reports the shifted character, so "shift+!" could never match.
    assert binding_from_event("exclam", SHIFT) == "!"


def test_keys_the_listener_could_not_parse_are_refused():
    assert binding_from_event("KP_Enter", 0) == "enter"
    assert binding_from_event("XF86AudioPlay", 0) is None


def test_shift_is_kept_for_keys_it_does_not_rewrite():
    assert binding_from_event("F9", SHIFT) == "shift+f9"


def test_empty_input_is_rejected():
    assert binding_from_event("", 0) is None


@pytest.mark.parametrize(
    "binding, expected",
    [
        ("f9", "F9"),
        ("ctrl+shift+f9", "Ctrl+Shift+F9"),
        ("esc", "Esc"),
        ("page_up", "PgUp"),
        ("a", "A"),
        ("", "—"),
    ],
)
def test_bindings_are_described_for_display(binding, expected):
    assert describe_binding(binding) == expected
