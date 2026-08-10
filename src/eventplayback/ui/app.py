"""The application window.

The window owns no recording or playback logic of its own: it translates user
intent into calls on :mod:`eventplayback.core` and renders whatever
:func:`eventplayback.ui.state.build_view` says the current state looks like.
"""

from __future__ import annotations

import logging
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from .. import __version__
from ..core.backends.base import RealClock
from ..core.backends.pynput_backend import PynputSource, PynputSynthesizer
from ..core.macro import Macro
from ..core.player import Player
from ..core.recorder import Recorder
from ..platform_support import permission_hint
from ..services.hotkeys import HotkeyManager, create_listener
from ..services.settings import Settings
from ..services.storage import MacroFileError, load_macro, save_macro
from .state import AppState, ViewModel, build_view, parse_loop_count

logger = logging.getLogger(__name__)

__all__ = ["App"]


class App(ctk.CTk):
    BLINK_INTERVAL_MS = 500
    TOAST_DURATION_MS = 2500

    def __init__(self, settings: Settings | None = None) -> None:
        super().__init__()
        self.settings = (settings or Settings.load()).normalized()
        ctk.set_appearance_mode(self.settings.appearance_mode)

        self.title("EventPlayback" if __version__ == "dev" else f"EventPlayback v{__version__}")
        self.geometry("420x100")
        self.resizable(False, False)
        self.attributes("-topmost", self.settings.always_on_top)

        clock = RealClock()
        self.recorder = Recorder(
            PynputSource(), clock, excluded_keys=self.settings.reserved_keys
        )
        self.player = Player(PynputSynthesizer(), clock)
        self.macro = Macro()

        self._state = AppState.IDLE
        self._countdown = 0
        self._blink_on = True
        self._blink_id: str | None = None
        self._countdown_id: str | None = None
        self._pending: str | None = None
        self._toast_token = 0

        self._build_widgets()

        self.recorder.on_event = lambda _event: self.after(0, self._refresh_info)
        self.player.on_complete = lambda: self.after(0, self._on_playback_complete)
        self.player.on_error = lambda message: self.after(0, self._toast, message)

        listener, self.hotkeys_suppressed = create_listener(suppress=self.settings.suppress_hotkeys)
        self.hotkeys = HotkeyManager(
            listener,
            on_error=lambda message: self.after(0, self._toast, f"Hotkeys unavailable: {message}"),
        )
        self.hotkeys.register("record", lambda: self.after(0, self.toggle_record))
        self.hotkeys.register("play", lambda: self.after(0, self.toggle_play))
        self.hotkeys.register("stop", lambda: self.after(0, self.request_stop))
        if not self.hotkeys.apply(self.settings.hotkeys):
            self.after(100, self._toast, permission_hint())
        elif self.settings.suppress_hotkeys and not self.hotkeys_suppressed:
            # Asked for, but unavailable on this platform or without the extra.
            self.after(100, self._toast, "Hotkey suppression needs Windows and 'keyboard'")

        self.protocol("WM_DELETE_WINDOW", self.close)
        self._render()

    # -- construction ---------------------------------------------------
    def _build_widgets(self) -> None:
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=8, pady=8)

        buttons = ctk.CTkFrame(main, fg_color="transparent")
        buttons.pack(fill="x", pady=(0, 8))

        bold = ctk.CTkFont(size=13, weight="bold")
        self.rec_btn = ctk.CTkButton(
            buttons, text="● Record", width=80, height=36, font=bold,
            fg_color="#c0392b", hover_color="#e74c3c", command=self.toggle_record,
        )
        self.rec_btn.pack(side="left", padx=(0, 4))

        self.stop_btn = ctk.CTkButton(
            buttons, text="■ Stop", width=80, height=36, font=bold,
            fg_color="#7f8c8d", hover_color="#95a5a6", command=self.request_stop, state="disabled",
        )
        self.stop_btn.pack(side="left", padx=(0, 4))

        self.play_btn = ctk.CTkButton(
            buttons, text="▶ Play", width=80, height=36, font=bold,
            fg_color="#2980b9", hover_color="#3498db", command=self.toggle_play,
        )
        self.play_btn.pack(side="left", padx=(0, 8))

        loop_frame = ctk.CTkFrame(buttons, fg_color="transparent")
        loop_frame.pack(side="left", padx=(4, 0))
        ctk.CTkLabel(loop_frame, text="×", font=ctk.CTkFont(size=12), text_color="#888").pack(side="left")
        self.loop_var = ctk.StringVar(value=str(self.settings.loop_count))
        self.loop_entry = ctk.CTkEntry(
            loop_frame, textvariable=self.loop_var, width=36, height=28,
            font=ctk.CTkFont(size=12), justify="center",
            fg_color="#3a3a3a", border_color="#555", text_color="white",
        )
        self.loop_entry.pack(side="left", padx=4)

        files = ctk.CTkFrame(buttons, fg_color="transparent")
        files.pack(side="right")
        ctk.CTkButton(files, text="📂", width=32, height=28, fg_color="#7f8c8d",
                      hover_color="#95a5a6", command=self.open_macro).pack(side="left", padx=2)
        ctk.CTkButton(files, text="💾", width=32, height=28, fg_color="#7f8c8d",
                      hover_color="#95a5a6", command=self.save_macro).pack(side="left", padx=2)

        status = ctk.CTkFrame(main, fg_color="transparent")
        status.pack(fill="x")
        self.status_label = ctk.CTkLabel(
            status, text="Idle", font=ctk.CTkFont(size=16, weight="bold"), text_color="white"
        )
        self.status_label.pack(side="left")
        self.info_label = ctk.CTkLabel(
            status, text="0 events | 0.0s", font=ctk.CTkFont(size=12), text_color="#aaa"
        )
        self.info_label.pack(side="right")

    # -- rendering ------------------------------------------------------
    def _view(self) -> ViewModel:
        return build_view(
            self._state,
            has_events=bool(self.macro.events),
            event_count=(
                self.recorder.event_count()
                if self._state is AppState.RECORDING
                else len(self.macro.events)
            ),
            duration=self.macro.duration,
            countdown=self._countdown,
            blink_on=self._blink_on,
        )

    def _render(self) -> None:
        view = self._view()
        self.status_label.configure(text=view.status, text_color="white")
        self.info_label.configure(text=view.info)
        self.configure(fg_color=view.accent)
        for button, spec in (
            (self.rec_btn, view.record),
            (self.stop_btn, view.stop),
            (self.play_btn, view.play),
        ):
            button.configure(state=spec.tk_state, fg_color=spec.color)
        self.loop_entry.configure(state="normal" if view.loop_enabled else "disabled")

    def _refresh_info(self) -> None:
        self.info_label.configure(text=self._view().info)

    def _toast(self, message: str) -> None:
        """Show a transient message without losing the status text underneath.

        Each toast carries a token so an earlier toast's timer cannot clear a
        later one.
        """
        self._toast_token += 1
        token = self._toast_token
        self.status_label.configure(text=message, text_color="#f1c40f")

        def restore() -> None:
            if token == self._toast_token:
                self.status_label.configure(text=self._view().status, text_color="white")

        self.after(self.TOAST_DURATION_MS, restore)

    # -- blinking -------------------------------------------------------
    def _start_blink(self) -> None:
        self._stop_blink(reset=False)
        self._blink_on = True
        self._blink()

    def _stop_blink(self, *, reset: bool = True) -> None:
        if self._blink_id is not None:
            self.after_cancel(self._blink_id)
            self._blink_id = None
        if reset:
            self._blink_on = True
            self.configure(fg_color=self._view().accent)

    def _blink(self) -> None:
        if self._state is AppState.IDLE:
            self._stop_blink()
            return
        self.configure(fg_color=self._view().accent)
        self._blink_on = not self._blink_on
        self._blink_id = self.after(self.BLINK_INTERVAL_MS, self._blink)

    # -- intents --------------------------------------------------------
    def toggle_record(self) -> None:
        if self._state is AppState.IDLE:
            self._start_countdown("record")
        elif self._state is AppState.RECORDING:
            self._finish_recording()

    def toggle_play(self) -> None:
        if self._state is AppState.IDLE:
            if self.macro.events:
                self._start_countdown("play")
            else:
                self._toast("No recording available")
        elif self._state is AppState.PLAYING:
            self._finish_playback()

    def request_stop(self) -> None:
        if self._state is AppState.COUNTDOWN:
            self._cancel_countdown()
        elif self._state is AppState.RECORDING:
            self._finish_recording()
        elif self._state is AppState.PLAYING:
            self._finish_playback()

    # -- countdown ------------------------------------------------------
    def _start_countdown(self, action: str) -> None:
        self._pending = action
        self._countdown = self.settings.countdown_seconds
        self._state = AppState.COUNTDOWN
        self._start_blink()
        self._render()
        self._tick_countdown()

    def _tick_countdown(self) -> None:
        if self._countdown > 0:
            self._render()
            self._countdown -= 1
            self._countdown_id = self.after(1000, self._tick_countdown)
            return
        self._countdown_id = None
        self._stop_blink()
        if self._pending == "record":
            self._begin_recording()
        else:
            self._begin_playback()

    def _cancel_countdown(self) -> None:
        if self._countdown_id is not None:
            self.after_cancel(self._countdown_id)
            self._countdown_id = None
        self._pending = None
        self._state = AppState.IDLE
        self._stop_blink()
        self._render()

    # -- recording ------------------------------------------------------
    def _begin_recording(self) -> None:
        self.recorder.set_excluded_keys(self.settings.reserved_keys)
        try:
            self.recorder.start()
        except RuntimeError as exc:
            self._state = AppState.IDLE
            self._render()
            logger.error("Could not start recording: %s", exc)
            self._toast(f"{exc}. {permission_hint()}")
            return
        self._state = AppState.RECORDING
        self._start_blink()
        self._render()

    def _finish_recording(self) -> None:
        events = self.recorder.stop()
        self.macro = Macro(name=self.macro.name, events=events)
        self._state = AppState.IDLE
        self._stop_blink()
        self._render()
        if events:
            self._toast(f"Recorded {len(events)} events")

    # -- playback -------------------------------------------------------
    def _begin_playback(self) -> None:
        loops, warning = parse_loop_count(self.loop_var.get())
        self.settings.loop_count = loops
        self.player.set_events(self.macro.events)
        self.player.set_loop(loops)
        if not self.player.start():
            self._state = AppState.IDLE
            self._render()
            self._toast("Nothing to play")
            return
        self._state = AppState.PLAYING
        self._start_blink()
        self._render()
        if warning:
            self._toast(warning)

    def _finish_playback(self) -> None:
        self.player.stop()
        self._state = AppState.IDLE
        self._stop_blink()
        self._render()

    def _on_playback_complete(self) -> None:
        self._state = AppState.IDLE
        self._stop_blink()
        self._render()
        self._toast("Playback complete")

    # -- files ----------------------------------------------------------
    def _initial_dir(self) -> str:
        return self.settings.last_directory or str(Path.home())

    def save_macro(self) -> None:
        if not self.macro.events:
            self._toast("No data available")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=self._initial_dir(),
            initialfile=f"{self.macro.name}.json",
        )
        if not path:
            return
        try:
            saved = save_macro(self.macro, path)
        except MacroFileError as exc:
            self._toast(str(exc))
            return
        self.settings.last_directory = str(saved.parent)
        self.macro.name = saved.stem
        self._toast("Saved")

    def open_macro(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")], initialdir=self._initial_dir()
        )
        if not path:
            return
        try:
            macro = load_macro(path)
        except MacroFileError as exc:
            self._toast(str(exc))
            return
        self.macro = macro
        self.settings.last_directory = str(Path(path).parent)
        self._render()
        self._toast(f"Loaded {len(macro.events)} events")

    # -- shutdown -------------------------------------------------------
    def close(self) -> None:
        self.hotkeys.stop()
        if self.recorder.is_recording():
            self.recorder.stop()
        if self.player.is_playing():
            self.player.stop()
        self.settings.save()
        self.destroy()
