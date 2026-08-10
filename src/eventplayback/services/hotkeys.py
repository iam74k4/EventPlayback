"""Global hotkey registration, kept independent of the listener backend."""

from __future__ import annotations

import logging
from collections.abc import Callable

from ..core.backends.base import HotkeyListener
from ..platform_support import IS_WINDOWS

logger = logging.getLogger(__name__)

__all__ = ["HotkeyManager", "create_listener", "select_hotkey_backend"]


def select_hotkey_backend(*, suppress: bool, is_windows: bool, keyboard_available: bool) -> str:
    """Decide which backend serves the hotkeys.

    ``pynput`` is the default because it needs no extra dependency and no
    elevated privileges. ``keyboard`` is chosen only when the user asked for
    hotkeys to be withheld from other applications, which is the one thing
    pynput cannot do, and only on Windows: elsewhere the library needs root.

    Returns:
        Either ``"keyboard"`` or ``"pynput"``.
    """
    if suppress and is_windows and keyboard_available:
        return "keyboard"
    return "pynput"


def _keyboard_available() -> bool:
    from importlib.util import find_spec

    try:
        return find_spec("keyboard") is not None
    except (ImportError, ValueError):
        return False


def create_listener(*, suppress: bool = False) -> tuple[HotkeyListener, bool]:
    """Build the hotkey listener that best matches what was asked for.

    Returns:
        The listener, and whether it will actually suppress hotkeys. The two
        can disagree: suppression is silently unavailable off Windows or
        without the optional dependency, and the caller should say so.
    """
    choice = select_hotkey_backend(
        suppress=suppress, is_windows=IS_WINDOWS, keyboard_available=_keyboard_available()
    )
    if choice == "keyboard":
        try:
            from ..core.backends.keyboard_backend import KeyboardHotkeyListener

            return KeyboardHotkeyListener(suppress=True), True
        except ImportError:
            logger.info("The keyboard library is unavailable; hotkeys will not be suppressed")

    from ..core.backends.pynput_backend import PynputHotkeyListener

    return PynputHotkeyListener(), False


class HotkeyManager:
    """Binds action names to hotkeys and keeps the listener in sync.

    Registration failures are reported rather than raised: on macOS the first
    run is refused until Accessibility permission is granted, and the
    application must stay usable through its buttons in the meantime.
    """

    def __init__(
        self,
        listener: HotkeyListener,
        *,
        on_error: Callable[[str], None] | None = None,
    ) -> None:
        self._listener = listener
        self._actions: dict[str, Callable[[], None]] = {}
        self._bindings: dict[str, str] = {}
        self._active = False
        self.on_error = on_error

    def register(self, action: str, callback: Callable[[], None]) -> None:
        """Attach the callback invoked when ``action``'s hotkey fires."""
        self._actions[action] = callback

    def apply(self, bindings: dict[str, str]) -> bool:
        """(Re)start the listener for ``{"record": "f9"}`` style bindings.

        Returns:
            ``True`` if the hotkeys are now active.
        """
        self._bindings = dict(bindings)
        resolved: dict[str, Callable[[], None]] = {}
        for action, hotkey in self._bindings.items():
            callback = self._actions.get(action)
            if callback is None:
                logger.debug("No callback registered for action %r", action)
                continue
            resolved[hotkey] = callback

        try:
            self._listener.start(resolved)
            self._active = True
        except Exception as exc:
            self._active = False
            logger.warning("Hotkeys unavailable: %s", exc)
            if self.on_error is not None:
                self.on_error(str(exc))
        return self._active

    @property
    def is_active(self) -> bool:
        return self._active

    @property
    def bindings(self) -> dict[str, str]:
        return dict(self._bindings)

    def stop(self) -> None:
        try:
            self._listener.stop()
        except Exception:
            logger.debug("Hotkey listener shutdown error", exc_info=True)
        self._active = False
