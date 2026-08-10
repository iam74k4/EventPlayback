"""The event model shared by the recorder, the player and the file format."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any

__all__ = ["CANONICAL_BUTTONS", "Event", "EventType"]


class EventType(Enum):
    MOUSE_MOVE = "mouse_move"
    MOUSE_CLICK = "mouse_click"
    MOUSE_SCROLL = "mouse_scroll"
    KEY_PRESS = "key_press"
    KEY_RELEASE = "key_release"


#: Buttons every supported platform reports. Extra buttons (``button8`` on X11,
#: for example) are accepted as-is so nothing is silently dropped.
CANONICAL_BUTTONS = frozenset({"left", "right", "middle"})

_OPTIONAL_FIELDS = ("x", "y", "button", "key", "pressed", "scroll_dx", "scroll_dy")


def _as_int(value: Any, field: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{field} must be a number, got {type(value).__name__}")
    return int(value)


@dataclass
class Event:
    """A single recorded input event, timestamped relative to recording start."""

    type: EventType
    timestamp: float
    x: int | None = None
    y: int | None = None
    button: str | None = None
    key: str | None = None
    pressed: bool | None = None
    scroll_dx: int | None = None
    scroll_dy: int | None = None

    @property
    def is_mouse(self) -> bool:
        return self.type in (EventType.MOUSE_MOVE, EventType.MOUSE_CLICK, EventType.MOUSE_SCROLL)

    @property
    def is_key(self) -> bool:
        return self.type in (EventType.KEY_PRESS, EventType.KEY_RELEASE)

    def shifted(self, delta: float) -> Event:
        """Return a copy whose timestamp is moved by ``delta`` seconds."""
        from dataclasses import replace

        return replace(self, timestamp=self.timestamp + delta)

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"type": self.type.value, "timestamp": self.timestamp}
        for name in _OPTIONAL_FIELDS:
            value = getattr(self, name)
            if value is not None:
                data[name] = value
        return data

    @classmethod
    def from_dict(cls, data: Any) -> Event:
        if not isinstance(data, dict):
            raise ValueError(f"Event must be an object, got {type(data).__name__}")
        for required in ("type", "timestamp"):
            if required not in data:
                raise ValueError(f"Required field ({required}) is missing")
        try:
            event_type = EventType(data["type"])
        except ValueError:
            raise ValueError(f"Invalid event type: {data['type']!r}") from None

        timestamp = data["timestamp"]
        if isinstance(timestamp, bool) or not isinstance(timestamp, (int, float)):
            raise ValueError(f"timestamp must be a number, got {type(timestamp).__name__}")

        button = data.get("button")
        if button is not None and not isinstance(button, str):
            raise ValueError(f"button must be a string, got {type(button).__name__}")
        key = data.get("key")
        if key is not None and not isinstance(key, str):
            raise ValueError(f"key must be a string, got {type(key).__name__}")
        pressed = data.get("pressed")
        if pressed is not None and not isinstance(pressed, bool):
            raise ValueError(f"pressed must be a boolean, got {type(pressed).__name__}")

        return cls(
            type=event_type,
            timestamp=float(timestamp),
            x=_as_int(data.get("x"), "x"),
            y=_as_int(data.get("y"), "y"),
            button=button,
            key=key,
            pressed=pressed,
            scroll_dx=_as_int(data.get("scroll_dx"), "scroll_dx"),
            scroll_dy=_as_int(data.get("scroll_dy"), "scroll_dy"),
        )
