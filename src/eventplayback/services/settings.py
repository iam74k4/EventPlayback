"""User settings, persisted as JSON in the per-user config directory."""

from __future__ import annotations

import json
import logging
import os
import tempfile
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

from ..platform_support import config_dir

logger = logging.getLogger(__name__)

__all__ = ["DEFAULT_HOTKEYS", "Settings"]

#: Action name -> hotkey. Kept in one place so the recorder, the hotkey
#: listener and the UI all agree on which keys are reserved.
DEFAULT_HOTKEYS: dict[str, str] = {
    "record": "f9",
    "play": "f10",
    "stop": "esc",
}

SETTINGS_FILENAME = "settings.json"

#: How many recently opened macros are remembered.
MAX_RECENT_FILES = 8


@dataclass
class Settings:
    """Everything that survives a restart."""

    loop_count: int = 1
    #: Pause inserted between repetitions, in seconds.
    loop_delay: float = 0.0
    countdown_seconds: int = 3
    always_on_top: bool = True
    appearance_mode: str = "dark"
    last_directory: str = ""
    #: Most recently opened macros, newest first.
    recent_files: list[str] = field(default_factory=list)
    hotkeys: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_HOTKEYS))
    #: Withhold hotkeys from the focused application. Windows only, and needs
    #: the optional ``keyboard`` dependency; ignored elsewhere.
    suppress_hotkeys: bool = False

    # -- validation -----------------------------------------------------
    def normalized(self) -> Settings:
        """Return a copy with every field clamped to a usable value."""
        hotkeys = dict(DEFAULT_HOTKEYS)
        for action, key in (self.hotkeys or {}).items():
            if action in DEFAULT_HOTKEYS and isinstance(key, str) and key.strip():
                hotkeys[action] = key.strip().lower()

        appearance = self.appearance_mode if self.appearance_mode in ("dark", "light", "system") else "dark"
        directory = self.last_directory if isinstance(self.last_directory, str) else ""
        if directory and not Path(directory).is_dir():
            directory = ""

        return Settings(
            loop_count=_clamp_int(self.loop_count, 0, 100_000, 1),
            loop_delay=_clamp_float(self.loop_delay, 0.0, 3600.0, 0.0),
            countdown_seconds=_clamp_int(self.countdown_seconds, 0, 60, 3),
            always_on_top=bool(self.always_on_top),
            appearance_mode=appearance,
            last_directory=directory,
            recent_files=_clean_recent(self.recent_files),
            hotkeys=hotkeys,
            suppress_hotkeys=bool(self.suppress_hotkeys),
        )

    def remember_file(self, path: str | os.PathLike[str]) -> None:
        """Move ``path`` to the front of the recent list."""
        entry = str(Path(path))
        remaining = [item for item in self.recent_files if item != entry]
        self.recent_files = [entry, *remaining][:MAX_RECENT_FILES]

    @property
    def reserved_keys(self) -> set[str]:
        """Key names that must be kept out of recordings.

        A binding such as ``ctrl+f9`` reserves each of its components, since
        recording only sees individual key names.
        """
        reserved: set[str] = set()
        for binding in self.hotkeys.values():
            for part in binding.split("+"):
                token = part.strip().lower()
                if token:
                    reserved.add(token)
        return reserved

    # -- persistence ----------------------------------------------------
    @staticmethod
    def default_path() -> Path:
        return config_dir() / SETTINGS_FILENAME

    @classmethod
    def from_dict(cls, data: Any) -> Settings:
        if not isinstance(data, dict):
            return cls()
        known = {f.name for f in fields(cls)}
        kwargs = {key: value for key, value in data.items() if key in known}
        try:
            return cls(**kwargs).normalized()
        except TypeError:
            logger.warning("Settings file has unexpected field types; using defaults")
            return cls()

    @classmethod
    def load(cls, path: str | os.PathLike[str] | None = None) -> Settings:
        """Load settings, falling back to defaults for anything unreadable."""
        target = Path(path) if path is not None else cls.default_path()
        try:
            if not target.is_file():
                return cls()
            return cls.from_dict(json.loads(target.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
            logger.warning("Could not read settings from %s: %s", target, exc)
            return cls()

    def save(self, path: str | os.PathLike[str] | None = None) -> bool:
        """Persist settings. Returns ``False`` if they could not be written."""
        target = Path(path) if path is not None else self.default_path()
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            payload = json.dumps(asdict(self.normalized()), ensure_ascii=False, indent=2)
            handle, temp_name = tempfile.mkstemp(dir=str(target.parent), prefix=".settings-", suffix=".tmp")
            try:
                with os.fdopen(handle, "w", encoding="utf-8") as stream:
                    stream.write(payload)
                os.replace(temp_name, target)
            except BaseException:
                try:
                    os.unlink(temp_name)
                except OSError:
                    pass
                raise
            return True
        except OSError as exc:
            logger.warning("Could not save settings to %s: %s", target, exc)
            return False


def _clamp_int(value: Any, low: int, high: int, fallback: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return fallback
    return max(low, min(high, number))


def _clamp_float(value: Any, low: float, high: float, fallback: float) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback
    if number != number:  # NaN, which no comparison would catch
        return fallback
    return max(low, min(high, number))


def _clean_recent(value: Any) -> list[str]:
    """Keep the readable entries, newest first, without duplicates.

    Files that have since been deleted are dropped, the same way an unusable
    ``last_directory`` is: offering a menu entry that can only fail is worse
    than a shorter menu.
    """
    if not isinstance(value, list):
        return []
    seen: set[str] = set()
    kept: list[str] = []
    for item in value:
        if not isinstance(item, str) or not item.strip():
            continue
        entry = str(Path(item))
        if entry in seen:
            continue
        seen.add(entry)
        try:
            if not Path(entry).is_file():
                continue
        except OSError:  # an unreachable network path, for instance
            continue
        kept.append(entry)
        if len(kept) >= MAX_RECENT_FILES:
            break
    return kept
