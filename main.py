"""EventPlayback - Lightweight Version"""

import customtkinter as ctk
from tkinter import filedialog
import json
import logging
import os
import time
import threading
import ctypes
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Optional, Callable, List, Dict, Union, Any

# Version information (can be set via environment variable during build)
__version__ = os.getenv("EVENTPLAYBACK_VERSION", "dev")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

import keyboard
from pynput import mouse, keyboard as pynput_kb
from pynput.mouse import Controller as MouseController, Button as MouseButton
from pynput.keyboard import Controller as KeyboardController, Key

# DPI Awareness setting (Windows) - Prevents coordinate misalignment
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # System DPI Aware (fallback)
    except Exception:
        pass

# Theme setting
ctk.set_appearance_mode("dark")


# === Data Models ===

class EventType(Enum):
    MOUSE_MOVE = "mouse_move"
    MOUSE_CLICK = "mouse_click"
    MOUSE_SCROLL = "mouse_scroll"
    KEY_PRESS = "key_press"
    KEY_RELEASE = "key_release"


@dataclass
class Event:
    type: EventType
    timestamp: float
    x: Optional[int] = None
    y: Optional[int] = None
    button: Optional[str] = None
    key: Optional[str] = None
    pressed: Optional[bool] = None
    scroll_dx: Optional[int] = None
    scroll_dy: Optional[int] = None

    def to_dict(self) -> Dict[str, Any]:
        d = {"type": self.type.value, "timestamp": self.timestamp}
        for k in ["x", "y", "button", "key", "pressed", "scroll_dx", "scroll_dy"]:
            v = getattr(self, k)
            if v is not None:
                d[k] = v
        return d

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Event":
        if "type" not in d or "timestamp" not in d:
            raise ValueError("Required fields (type, timestamp) are missing")
        try:
            event_type = EventType(d["type"])
        except ValueError:
            raise ValueError(f"Invalid event type: {d['type']}")
        if not isinstance(d["timestamp"], (int, float)):
            raise ValueError(f"timestamp must be a number: {type(d['timestamp'])}")
        return cls(
            type=event_type,
            timestamp=float(d["timestamp"]),
            x=d.get("x"), y=d.get("y"),
            button=d.get("button"), key=d.get("key"),
            pressed=d.get("pressed"),
            scroll_dx=d.get("scroll_dx"), scroll_dy=d.get("scroll_dy"),
        )


@dataclass
class Macro:
    name: str = "New Macro"
    events: List[Event] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "created_at": self.created_at,
            "events": [e.to_dict() for e in self.events],
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Macro":
        if not isinstance(d, dict):
            raise ValueError("Data must be in dictionary format")
        if "events" not in d:
            raise ValueError("Required field (events) is missing")
        if not isinstance(d["events"], list):
            raise ValueError("events must be in list format")
        events = []
        for i, e in enumerate(d["events"]):
            try:
                events.append(Event.from_dict(e))
            except Exception as ex:
                raise ValueError(f"Failed to load event [{i}]: {str(ex)}")
        return cls(
            name=d.get("name", "Untitled"),
            events=events,
            created_at=d.get("created_at", ""),
        )

    @property
    def duration(self) -> float:
        return self.events[-1].timestamp if self.events else 0.0


# === Recording ===

class Recorder:
    """
    Record mouse and keyboard input events.
    
    Usage:
        recorder = Recorder()
        recorder.start()
        # ... perform actions ...
        events = recorder.stop()
    """
    INTERVAL = 0.02  # Mouse movement throttling interval (seconds) - Record at 20ms intervals for performance optimization
    EXCLUDED_HOTKEYS = ("f9", "f10", "escape")  # Hotkey exclusion list (not recorded)

    def __init__(self) -> None:
        self._events: List[Event] = []
        self._start_time: float = 0.0
        self._recording: bool = False
        self._last_move: float = 0.0
        self._mouse_listener: Optional[Any] = None
        self._kb_listener: Optional[Any] = None
        self._lock: threading.Lock = threading.Lock()
        self.on_event: Optional[Callable[[Event], None]] = None

    def start(self) -> None:
        if self._recording:
            return
        self._events.clear()
        self._start_time = time.perf_counter()
        self._last_move = 0.0
        self._recording = True

        self._mouse_listener = mouse.Listener(
            on_move=self._on_move,
            on_click=self._on_click,
            on_scroll=self._on_scroll,
        )
        self._kb_listener = pynput_kb.Listener(
            on_press=self._on_press,
            on_release=self._on_release,
        )
        self._mouse_listener.start()
        self._kb_listener.start()

    def stop(self) -> List[Event]:
        if not self._recording:
            return self._events.copy()
        self._recording = False
        if self._mouse_listener:
            self._mouse_listener.stop()
            try:
                self._mouse_listener.join(timeout=0.5)
            except Exception as e:
                logger.error(f"Mouse listener stop error: {e}", exc_info=True)
        if self._kb_listener:
            self._kb_listener.stop()
            try:
                self._kb_listener.join(timeout=0.5)
            except Exception as e:
                logger.error(f"Keyboard listener stop error: {e}", exc_info=True)
        return self._events.copy()

    def is_recording(self) -> bool:
        return self._recording

    def event_count(self) -> int:
        with self._lock:
            return len(self._events)

    def _ts(self) -> float:
        return time.perf_counter() - self._start_time

    def _add(self, event: Event) -> None:
        with self._lock:
            self._events.append(event)
        if self.on_event:
            self.on_event(event)

    def _on_move(self, x: float, y: float) -> None:
        if not self._recording:
            return
        t = self._ts()
        if t - self._last_move < self.INTERVAL:
            return
        self._last_move = t
        self._add(Event(EventType.MOUSE_MOVE, t, x=int(x), y=int(y)))

    def _on_click(self, x: float, y: float, button: Any, pressed: bool) -> None:
        if not self._recording:
            return
        btn = "left" if button == mouse.Button.left else "right" if button == mouse.Button.right else "middle"
        self._add(Event(EventType.MOUSE_CLICK, self._ts(), x=int(x), y=int(y), button=btn, pressed=pressed))

    def _on_scroll(self, x: float, y: float, dx: int, dy: int) -> None:
        if not self._recording:
            return
        self._add(Event(EventType.MOUSE_SCROLL, self._ts(), x=int(x), y=int(y), scroll_dx=dx, scroll_dy=dy))

    def _on_press(self, key: Any) -> None:
        if not self._recording:
            return
        name = self._key_name(key)
        if name and name not in self.EXCLUDED_HOTKEYS:
            self._add(Event(EventType.KEY_PRESS, self._ts(), key=name, pressed=True))

    def _on_release(self, key: Any) -> None:
        if not self._recording:
            return
        name = self._key_name(key)
        if name and name not in self.EXCLUDED_HOTKEYS:
            self._add(Event(EventType.KEY_RELEASE, self._ts(), key=name, pressed=False))

    def _key_name(self, key) -> Optional[str]:
        """
        Convert pynput Key object to string name.
        
        Args:
            key: pynput Key object
            
        Returns:
            Key name string (e.g., "a", "f1", "space"), None if conversion fails
        """
        # Normal character keys (a-z, 0-9, etc.)
        try:
            if key.char is not None:
                return key.char
        except AttributeError:
            pass
        # Special keys (F1-F12, arrow keys, etc.)
        try:
            return key.name
        except AttributeError:
            return None


# === Playback ===

class Player:
    """
    Play back recorded mouse and keyboard events.
    
    Usage:
        player = Player()
        player.set_events(events)
        player.set_loop(1)  # Play once
        player.start()
    """
    # Valid special characters for single character keys
    VALID_SPECIAL_CHARS = "!@#$%^&*()_+-=[]{}|;:'\",.<>?/~`"
    
    def __init__(self) -> None:
        self._events: List[Event] = []
        self._playing: bool = False
        self._stop_flag: bool = False
        self._loop_count: int = 1
        self._thread: Optional[threading.Thread] = None
        self._lock: threading.Lock = threading.Lock()
        self._mouse: MouseController = MouseController()
        self._kb: KeyboardController = KeyboardController()
        self.on_complete: Optional[Callable[[], None]] = None
        self.on_error: Optional[Callable[[str], None]] = None

    def set_events(self, events: List[Event]) -> None:
        self._events = events.copy()

    def set_loop(self, count: int) -> None:
        self._loop_count = max(0, count)

    def start(self) -> None:
        with self._lock:
            if self._playing or not self._events:
                return
            self._playing = True
            self._stop_flag = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        with self._lock:
            if not self._playing:
                return
            self._stop_flag = True
        if self._thread:
            self._thread.join(timeout=1.0)
            if self._thread.is_alive():
                # Warn if thread is still alive after timeout
                logger.warning("Playback thread is taking too long to stop")
        with self._lock:
            self._playing = False

    def is_playing(self) -> bool:
        with self._lock:
            return self._playing

    def _run(self) -> None:
        loop = 0
        while not self._stop_flag:
            loop += 1
            self._play_once()
            if self._loop_count > 0 and loop >= self._loop_count:
                break
        with self._lock:
            self._playing = False
        if self.on_complete and not self._stop_flag:
            self.on_complete()

    def _play_once(self) -> None:
        if not self._events:
            return
        start = time.perf_counter()
        for event in self._events:
            if self._stop_flag:
                break
            wait = start + event.timestamp - time.perf_counter()
            if wait > 0:
                time.sleep(wait)
            self._play_event(event)

    def _play_event(self, e: Event) -> None:
        """Play a single event."""
        try:
            if e.type == EventType.MOUSE_MOVE:
                self._play_mouse_move(e)
            elif e.type == EventType.MOUSE_CLICK:
                self._play_mouse_click(e)
            elif e.type == EventType.MOUSE_SCROLL:
                self._play_mouse_scroll(e)
            elif e.type == EventType.KEY_PRESS:
                self._play_key_press(e)
            elif e.type == EventType.KEY_RELEASE:
                self._play_key_release(e)
        except Exception as ex:
            error_msg = f"Playback error: {ex}"
            logger.error(error_msg, exc_info=True)
            if self.on_error:
                self.on_error(error_msg)

    def _play_mouse_move(self, e: Event) -> None:
        """Play mouse move event."""
        if e.x is not None and e.y is not None:
            self._mouse.position = (e.x, e.y)

    def _play_mouse_click(self, e: Event) -> None:
        """Play mouse click event."""
        if e.button and e.pressed is not None and e.x is not None and e.y is not None:
            btn = MouseButton.left if e.button == "left" else MouseButton.right if e.button == "right" else MouseButton.middle
            if e.pressed:
                self._mouse.press(btn)
            else:
                self._mouse.release(btn)

    def _play_mouse_scroll(self, e: Event) -> None:
        """Play mouse scroll event."""
        if e.x is not None and e.y is not None:
            if e.scroll_dy:
                self._mouse.scroll(0, e.scroll_dy)
            if e.scroll_dx:
                self._mouse.scroll(e.scroll_dx, 0)

    def _play_key_press(self, e: Event) -> None:
        """Play key press event."""
        if e.key is not None:
            key = self._to_key(e.key)
            if key is not None:
                self._kb.press(key)

    def _play_key_release(self, e: Event) -> None:
        """Play key release event."""
        if e.key is not None:
            key = self._to_key(e.key)
            if key is not None:
                self._kb.release(key)

    def _to_key(self, name: str) -> Union[Key, str, None]:
        """
        Convert key name string to pynput Key object.
        
        Args:
            name: Key name string (e.g., "a", "space", "f1")
            
        Returns:
            pynput Key object or string for normal keys, None for invalid keys
        """
        if name is None:
            return None
        name_lower = name.lower()
        special = {
            "space": Key.space, "enter": Key.enter, "tab": Key.tab,
            "backspace": Key.backspace, "delete": Key.delete,
            "escape": Key.esc, "shift": Key.shift, "shift_l": Key.shift_l,
            "shift_r": Key.shift_r, "ctrl": Key.ctrl, "ctrl_l": Key.ctrl_l,
            "ctrl_r": Key.ctrl_r, "alt": Key.alt, "alt_l": Key.alt_l,
            "alt_r": Key.alt_r, "alt_gr": Key.alt_gr,
            "caps_lock": Key.caps_lock, "up": Key.up, "down": Key.down,
            "left": Key.left, "right": Key.right, "home": Key.home,
            "end": Key.end, "page_up": Key.page_up, "page_down": Key.page_down,
            "insert": Key.insert, "f1": Key.f1, "f2": Key.f2, "f3": Key.f3,
            "f4": Key.f4, "f5": Key.f5, "f6": Key.f6, "f7": Key.f7,
            "f8": Key.f8, "f9": Key.f9, "f10": Key.f10, "f11": Key.f11, "f12": Key.f12,
        }
        result = special.get(name_lower)
        if result is not None:
            return result
        # Normal character keys (single alphanumeric or special character)
        if len(name_lower) == 1 and (name_lower.isalnum() or name_lower in self.VALID_SPECIAL_CHARS):
            return name_lower
        # Invalid key name - log warning and return None
        logger.warning(f"Unknown key name '{name}' will be ignored")
        return None


# === GUI ===

class App(ctk.CTk):
    STATE_IDLE = 0
    STATE_COUNTDOWN = 1
    STATE_RECORDING = 2
    STATE_PLAYING = 3

    COUNTDOWN_SECONDS = 3  # Countdown time (seconds) - Wait time before starting recording/playback
    BLINK_INTERVAL_MS = 500  # UI blink interval (milliseconds) - Blink speed for status display
    TOAST_DURATION_MS = 2000  # Toast message display time (milliseconds) - Display duration for notification messages

    COLORS = {
        "idle": "#2b2b2b",
        "countdown": "#f39c12",
        "recording": "#e74c3c",
        "playing": "#27ae60",
    }

    def __init__(self) -> None:
        super().__init__()
        title = f"EventPlayback" if __version__ == "dev" else f"EventPlayback v{__version__}"
        self.title(title)
        self.geometry("420x100")
        self.resizable(False, False)
        self.configure(fg_color=self.COLORS["idle"])
        self.attributes("-topmost", True)

        self.recorder = Recorder()
        self.player = Player()
        self.macro = Macro()

        self._state: int = self.STATE_IDLE
        self._countdown: int = 0
        self._blink_on: bool = True
        self._blink_id: Optional[str] = None
        self._countdown_id: Optional[str] = None
        self._pending: Optional[str] = None

        self._setup_ui()
        self._setup_hotkeys()

        self.recorder.on_event = lambda e: self.after(0, self._update_info)
        self.player.on_complete = lambda: self.after(0, self._on_complete)
        self.player.on_error = lambda msg: self.after(0, lambda: self._toast(msg))

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_ui(self) -> None:
        # Main frame
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=8, pady=8)

        # Button frame
        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 8))

        self.rec_btn = ctk.CTkButton(btn_frame, text="● Record", width=80, height=36,
                                      font=ctk.CTkFont(size=13, weight="bold"),
                                      fg_color="#c0392b", hover_color="#e74c3c",
                                      command=self._on_rec)
        self.rec_btn.pack(side="left", padx=(0, 4))

        self.stop_btn = ctk.CTkButton(btn_frame, text="■ Stop", width=80, height=36,
                                       font=ctk.CTkFont(size=13, weight="bold"),
                                       fg_color="#7f8c8d", hover_color="#95a5a6",
                                       command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(0, 4))

        self.play_btn = ctk.CTkButton(btn_frame, text="▶ Play", width=80, height=36,
                                       font=ctk.CTkFont(size=13, weight="bold"),
                                       fg_color="#2980b9", hover_color="#3498db",
                                       command=self._on_play)
        self.play_btn.pack(side="left", padx=(0, 8))

        # Loop
        loop_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        loop_frame.pack(side="left", padx=(4, 0))
        ctk.CTkLabel(loop_frame, text="×", font=ctk.CTkFont(size=12), text_color="#888").pack(side="left")
        self.loop_var = ctk.StringVar(value="1")
        self.loop_entry = ctk.CTkEntry(loop_frame, textvariable=self.loop_var, width=36, height=28,
                                        font=ctk.CTkFont(size=12), justify="center",
                                        fg_color="#3a3a3a", border_color="#555", text_color="white")
        self.loop_entry.pack(side="left", padx=4)

        # File
        file_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        file_frame.pack(side="right")
        ctk.CTkButton(file_frame, text="📂", width=32, height=28, fg_color="#7f8c8d",
                      hover_color="#95a5a6", command=self._open).pack(side="left", padx=2)
        ctk.CTkButton(file_frame, text="💾", width=32, height=28, fg_color="#7f8c8d",
                      hover_color="#95a5a6", command=self._save).pack(side="left", padx=2)

        # Status
        status_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        status_frame.pack(fill="x")

        self.status_label = ctk.CTkLabel(status_frame, text="Idle",
                                          font=ctk.CTkFont(size=16, weight="bold"),
                                          text_color="white")
        self.status_label.pack(side="left")

        self.info_label = ctk.CTkLabel(status_frame, text="0 events | 0.0s",
                                        font=ctk.CTkFont(size=12), text_color="#aaa")
        self.info_label.pack(side="right")

    def _setup_hotkeys(self) -> None:
        # Hotkeys are called from outside the main thread, so update GUI via after()
        keyboard.add_hotkey("f9", lambda: self.after(0, self._on_rec))
        keyboard.add_hotkey("f10", lambda: self.after(0, self._on_play))
        keyboard.add_hotkey("escape", lambda: self.after(0, self._on_stop))

    def _on_rec(self) -> None:
        if self._state == self.STATE_IDLE:
            self._start_countdown("record")
        elif self._state == self.STATE_RECORDING:
            self._stop_rec()

    def _on_play(self) -> None:
        if self._state == self.STATE_IDLE:
            if self.macro.events:
                self._start_countdown("play")
            else:
                self._toast("No recording available")
        elif self._state == self.STATE_PLAYING:
            self._stop_play()

    def _on_stop(self) -> None:
        if self._state == self.STATE_COUNTDOWN:
            self._cancel_countdown()
        elif self._state == self.STATE_RECORDING:
            self._stop_rec()
        elif self._state == self.STATE_PLAYING:
            self._stop_play()

    def _start_countdown(self, action: str) -> None:
        self._state = self.STATE_COUNTDOWN
        self._pending = action
        self._countdown = self.COUNTDOWN_SECONDS
        self._update_state()
        self._do_countdown()
        self._start_blink()

    def _do_countdown(self) -> None:
        if self._countdown > 0:
            self.status_label.configure(text=str(self._countdown))
            self._countdown -= 1
            self._countdown_id = self.after(1000, self._do_countdown)
        else:
            self._stop_blink()
            if self._pending == "record":
                self._start_rec()
            else:
                self._start_play()

    def _cancel_countdown(self) -> None:
        if self._countdown_id:
            self.after_cancel(self._countdown_id)
        self._stop_blink()
        self._state = self.STATE_IDLE
        self._update_state()

    def _start_rec(self) -> None:
        self._state = self.STATE_RECORDING
        self.recorder.start()
        self._start_blink()
        self._update_state()

    def _stop_rec(self) -> None:
        events = self.recorder.stop()
        self.macro.events = events
        self._stop_blink()
        self._state = self.STATE_IDLE
        self._update_state()
        if events:
            self._toast(f"Recorded {len(events)} events")

    def _start_play(self) -> None:
        self._state = self.STATE_PLAYING
        try:
            loop = int(self.loop_var.get())
            if loop < 0:
                loop = 0  # 0 means infinite
            elif loop > 1000:  # Reasonable upper limit
                loop = 1000
        except (ValueError, TypeError):
            loop = 1
        self.player.set_events(self.macro.events)
        self.player.set_loop(loop)
        self.player.start()
        self._start_blink()
        self._update_state()

    def _stop_play(self) -> None:
        self.player.stop()
        self._stop_blink()
        self._state = self.STATE_IDLE
        self._update_state()

    def _on_complete(self) -> None:
        self._stop_blink()
        self._state = self.STATE_IDLE
        self._update_state()
        self._toast("Playback complete")

    def _start_blink(self) -> None:
        self._blink_on = True
        self._do_blink()

    def _stop_blink(self) -> None:
        if self._blink_id:
            self.after_cancel(self._blink_id)
            self._blink_id = None
        self.configure(fg_color=self.COLORS["idle"])

    def _do_blink(self) -> None:
        if self._state == self.STATE_IDLE:
            return
        color = {
            self.STATE_COUNTDOWN: "countdown",
            self.STATE_RECORDING: "recording",
            self.STATE_PLAYING: "playing",
        }.get(self._state, "idle")
        bg = self.COLORS[color] if self._blink_on else self.COLORS["idle"]
        self.configure(fg_color=bg)
        self._blink_on = not self._blink_on
        self._blink_id = self.after(self.BLINK_INTERVAL_MS, self._do_blink)

    def _update_state(self) -> None:
        if self._state == self.STATE_IDLE:
            self.status_label.configure(text="Idle")
            self.rec_btn.configure(state="normal", fg_color="#c0392b")
            self.stop_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.play_btn.configure(state="normal" if self.macro.events else "disabled",
                                     fg_color="#2980b9" if self.macro.events else "#7f8c8d")
            self.loop_entry.configure(state="normal")
        elif self._state == self.STATE_COUNTDOWN:
            self.rec_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.stop_btn.configure(state="normal", fg_color="#e74c3c")
            self.play_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.loop_entry.configure(state="disabled")
        elif self._state == self.STATE_RECORDING:
            self.status_label.configure(text="● Recording")
            self.rec_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.stop_btn.configure(state="normal", fg_color="#e74c3c")
            self.play_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.loop_entry.configure(state="disabled")
        elif self._state == self.STATE_PLAYING:
            self.status_label.configure(text="▶ Playing")
            self.rec_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.stop_btn.configure(state="normal", fg_color="#e74c3c")
            self.play_btn.configure(state="disabled", fg_color="#7f8c8d")
            self.loop_entry.configure(state="disabled")
        self._update_info()

    def _update_info(self) -> None:
        if self._state == self.STATE_RECORDING:
            count = self.recorder.event_count()
            self.info_label.configure(text=f"{count} events | Recording...")
        else:
            count = len(self.macro.events)
            dur = self.macro.duration
            self.info_label.configure(text=f"{count} events | {dur:.1f}s")

    def _toast(self, msg: str) -> None:
        self.status_label.configure(text=msg, text_color="#f1c40f")
        self.after(self.TOAST_DURATION_MS, lambda: self.status_label.configure(text="Idle", text_color="white") if self._state == self.STATE_IDLE else None)

    def _save(self) -> None:
        if not self.macro.events:
            self._toast("No data available")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if path:
            try:
                # Path validation
                if not os.path.isabs(path):
                    path = os.path.abspath(path)
                # Normalize path to prevent path traversal attacks
                normalized_path = os.path.normpath(path)
                # Check for path traversal attempts (more robust check)
                path_parts = normalized_path.split(os.sep)
                if any(part == ".." or part.startswith("~") for part in path_parts):
                    self._toast("Invalid file path")
                    logger.warning(f"Path traversal attempt detected: {path}")
                    return
                # Ensure directory exists
                dir_path = os.path.dirname(normalized_path)
                if dir_path and not os.path.exists(dir_path):
                    os.makedirs(dir_path, exist_ok=True)
                # Save file
                with open(normalized_path, "w", encoding="utf-8") as f:
                    json.dump(self.macro.to_dict(), f, ensure_ascii=False, indent=2)
                self._toast("Saved")
            except PermissionError:
                self._toast("File is in use")
            except OSError as e:
                self._toast(f"Save error: {str(e)}")
            except Exception as e:
                self._toast(f"Save failed: {str(e)}")

    def _open(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            try:
                # Path validation
                if not os.path.isabs(path):
                    path = os.path.abspath(path)
                # Normalize path to prevent path traversal attacks
                normalized_path = os.path.normpath(path)
                # Check for path traversal attempts (more robust check)
                path_parts = normalized_path.split(os.sep)
                if any(part == ".." or part.startswith("~") for part in path_parts):
                    self._toast("Invalid file path")
                    logger.warning(f"Path traversal attempt detected: {path}")
                    return
                # Verify file exists and is a file (not a directory)
                if not os.path.isfile(normalized_path):
                    self._toast("File not found")
                    return
                with open(normalized_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.macro = Macro.from_dict(data)
                self._update_info()
                self._update_state()
                self._toast(f"Loaded {len(self.macro.events)} events")
            except FileNotFoundError:
                self._toast("File not found")
            except json.JSONDecodeError as e:
                self._toast(f"JSON format error: {str(e)}")
            except KeyError as e:
                self._toast(f"Missing required field: {str(e)}")
            except ValueError as e:
                self._toast(f"Data format error: {str(e)}")
            except Exception as e:
                self._toast(f"Load failed: {str(e)}")

    def _on_close(self) -> None:
        # Clean up hotkeys
        try:
            keyboard.unhook_all()
            # clear_all_hotkeys may not exist in all versions of keyboard library
            if hasattr(keyboard, 'clear_all_hotkeys'):
                keyboard.clear_all_hotkeys()
        except Exception as e:
            logger.error(f"Hotkey cleanup error: {e}", exc_info=True)
        # Stop recording/playing
        if self.recorder.is_recording():
            self.recorder.stop()
        if self.player.is_playing():
            self.player.stop()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
