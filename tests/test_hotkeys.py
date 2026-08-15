from __future__ import annotations

import pytest

from eventplayback.core.backends.fake import FakeHotkeyListener
from eventplayback.services.hotkeys import HotkeyManager, select_hotkey_backend


# -- backend selection --------------------------------------------------
@pytest.mark.parametrize(
    "suppress, is_windows, available, expected",
    [
        (False, True, True, "pynput"),    # no suppression asked for
        (True, True, True, "keyboard"),   # the one case keyboard is needed
        (True, True, False, "pynput"),    # optional dependency not installed
        (True, False, True, "pynput"),    # keyboard needs root off Windows
        (False, False, False, "pynput"),
    ],
)
def test_backend_selection(suppress, is_windows, available, expected):
    assert (
        select_hotkey_backend(
            suppress=suppress, is_windows=is_windows, keyboard_available=available
        )
        == expected
    )


# -- manager ------------------------------------------------------------
@pytest.fixture
def manager():
    listener = FakeHotkeyListener()
    return listener, HotkeyManager(listener)


def test_bindings_reach_the_listener(manager):
    listener, mgr = manager
    fired = []
    mgr.register("record", lambda: fired.append("record"))
    mgr.register("play", lambda: fired.append("play"))

    assert mgr.apply({"record": "f9", "play": "f10"})
    assert mgr.is_active
    assert set(listener.bindings) == {"f9", "f10"}

    listener.trigger("f9")
    assert fired == ["record"]


def test_actions_without_a_callback_are_skipped(manager):
    listener, mgr = manager
    mgr.register("record", lambda: None)
    mgr.apply({"record": "f9", "explode": "f12"})
    assert set(listener.bindings) == {"f9"}


def test_rebinding_replaces_the_previous_keys(manager):
    listener, mgr = manager
    mgr.register("record", lambda: None)
    mgr.apply({"record": "f9"})
    mgr.apply({"record": "ctrl+r"})
    assert set(listener.bindings) == {"ctrl+r"}
    assert mgr.bindings == {"record": "ctrl+r"}


def test_registration_failure_is_reported_not_raised():
    class Failing(FakeHotkeyListener):
        def start(self, bindings):
            raise RuntimeError("permission denied")

    errors: list[str] = []
    mgr = HotkeyManager(Failing(), on_error=errors.append)
    mgr.register("record", lambda: None)

    assert mgr.apply({"record": "f9"}) is False
    assert not mgr.is_active
    assert errors == ["permission denied"]


def test_stop_deactivates_the_listener(manager):
    listener, mgr = manager
    mgr.register("record", lambda: None)
    mgr.apply({"record": "f9"})
    mgr.stop()
    assert not listener.running
    assert not mgr.is_active


def test_stop_survives_a_failing_listener():
    class Failing(FakeHotkeyListener):
        def stop(self):
            raise RuntimeError("already gone")

    mgr = HotkeyManager(Failing())
    mgr.stop()  # must not raise
    assert not mgr.is_active
