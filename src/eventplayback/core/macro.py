"""A macro is a named, ordered sequence of events plus its metadata."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from .events import Event

__all__ = ["FORMAT_VERSION", "Macro"]

#: Bumped whenever the on-disk layout changes incompatibly. Files without the
#: field are treated as version 1, which is what EventPlayback 1.0 wrote.
FORMAT_VERSION = 1


@dataclass
class Macro:
    name: str = "New Macro"
    events: list[Event] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    @property
    def duration(self) -> float:
        """Length of the macro in seconds."""
        if not self.events:
            return 0.0
        return max(event.timestamp for event in self.events)

    def __len__(self) -> int:
        return len(self.events)

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": FORMAT_VERSION,
            "name": self.name,
            "created_at": self.created_at,
            "events": [event.to_dict() for event in self.events],
        }

    @classmethod
    def from_dict(cls, data: Any) -> Macro:
        if not isinstance(data, dict):
            raise ValueError("Data must be in dictionary format")
        if "events" not in data:
            raise ValueError("Required field (events) is missing")
        if not isinstance(data["events"], list):
            raise ValueError("events must be in list format")

        version = data.get("version", FORMAT_VERSION)
        if isinstance(version, int) and version > FORMAT_VERSION:
            raise ValueError(
                f"File format version {version} is newer than supported version {FORMAT_VERSION}"
            )

        events: list[Event] = []
        for index, raw in enumerate(data["events"]):
            try:
                events.append(Event.from_dict(raw))
            except Exception as exc:
                raise ValueError(f"Failed to load event [{index}]: {exc}") from exc

        # A stable sort keeps press/release pairs that share a timestamp in the
        # order they were written while still tolerating out-of-order files.
        events.sort(key=lambda event: event.timestamp)

        name = data.get("name", "Untitled")
        created_at = data.get("created_at", "")
        return cls(
            name=name if isinstance(name, str) else "Untitled",
            events=events,
            created_at=created_at if isinstance(created_at, str) else "",
        )
