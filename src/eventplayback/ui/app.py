"""The application window.

The window owns no recording or playback logic of its own: it translates user
intent into calls on :mod:`eventplayback.core` and renders whatever
:func:`eventplayback.ui.state.build_view` says the current state looks like.
"""

from __future__ import annotations

import logging
import math
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

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
from . import theme
from .banner import Banner
from .keycapture import describe_binding
from .settings_dialog import SettingsDialog
from .state import AppState, ViewModel, build_view, elide, format_loop_label, parse_loop_count
from .tooltip import Tooltip

logger = logging.getLogger(__name__)

__all__ = ["App"]


class App(ctk.CTk):
    #: One step of the indicator pulse. Short enough to read as a fade rather
    #: than a blink, which is what the old full-window flash looked like.
    PULSE_INTERVAL_MS = 110
    #: Fade amounts through one cycle: bright, out, and back again.
    PULSE_STEPS = (0.0, 0.12, 0.30, 0.48, 0.55, 0.48, 0.30, 0.12)
    #: The countdown bar is redrawn far more often than once a second so that
    #: it slides rather than jumps between whole seconds.
    COUNTDOWN_INTERVAL_MS = 50
    #: The recorder fires per input event, which is far too often to touch a
    #: widget; the running count is polled at this interval instead.
    RECORDING_POLL_MS = 100
    #: How often the playback position is read off the player.
    PLAYBACK_POLL_MS = 60

    #: Longest macro name shown before it is elided. Keeping the label a fixed
    #: width is what stops a long name from stretching the window.
    MAX_NAME_CHARS = 22

    def __init__(self, settings: Settings | None = None) -> None:
        super().__init__()
        self.settings = (settings or Settings.load()).normalized()
        ctk.set_appearance_mode(self.settings.appearance_mode)

        self.title("EventPlayback" if __version__ == "dev" else f"EventPlayback v{__version__}")
        # No explicit geometry anywhere: Tk then keeps the window at whatever
        # the visible rows ask for, growing for the banner and shrinking again
        # afterwards. Computing it by hand went wrong on a scaled display,
        # where CustomTkinter re-applies its scaling to any geometry string and
        # the window doubled in width every time a message appeared.
        self.resizable(False, False)
        self.attributes("-topmost", self.settings.always_on_top)
        self.configure(fg_color=theme.WINDOW)

        clock = RealClock()
        self.recorder = Recorder(
            PynputSource(), clock, excluded_keys=self.settings.reserved_keys
        )
        self.player = Player(PynputSynthesizer(), clock)
        self.macro = Macro()

        self._state = AppState.IDLE
        self._countdown = 0
        self._countdown_total = 0
        self._countdown_deadline = 0.0
        self._countdown_progress: float | None = None
        self._pulse = 0.0
        self._pulse_index = 0
        self._pulse_id: str | None = None
        self._countdown_id: str | None = None
        self._recording_poll_id: str | None = None
        self._playback_poll_id: str | None = None
        self._pending: str | None = None
        self._settings_dialog: SettingsDialog | None = None
        self._recent_menu = tk.Menu(self, tearoff=0)
        #: Whether the macro in memory differs from what is on disk.
        self._dirty = False

        self._build_widgets()
        self._bind_shortcuts()
        self._refresh_hotkey_tooltips()

        self.player.on_complete = lambda: self.after(0, self._on_playback_complete)
        self.player.on_error = lambda message: self.after(0, self._notify, message, "error")

        listener, self.hotkeys_suppressed = create_listener(suppress=self.settings.suppress_hotkeys)
        self.hotkeys = HotkeyManager(
            listener,
            on_error=lambda message: self.after(
                0, self._notify, f"Hotkeys unavailable: {message}", "error"
            ),
        )
        self.hotkeys.register("record", lambda: self.after(0, self.toggle_record))
        self.hotkeys.register("play", lambda: self.after(0, self.toggle_play))
        self.hotkeys.register("stop", lambda: self.after(0, self.request_stop))
        if not self.hotkeys.apply(self.settings.hotkeys):
            self.after(100, self._notify, permission_hint(), "error")
        elif self.settings.suppress_hotkeys and not self.hotkeys_suppressed:
            # Asked for, but unavailable on this platform or without the extra.
            self.after(
                100, self._notify, "Hotkey suppression needs Windows and 'keyboard'", "warning"
            )

        self.protocol("WM_DELETE_WINDOW", self.close)
        self._render()

    # -- construction ---------------------------------------------------
    def _build_widgets(self) -> None:
        bold = ctk.CTkFont(size=theme.FONT_SIZE_LG, weight="bold")
        small = ctk.CTkFont(size=theme.FONT_SIZE_MD)

        self._root_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._root_frame.pack(fill="both", expand=True, padx=12, pady=10)

        # -- header: what the application is doing right now
        header = ctk.CTkFrame(self._root_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 8))

        self.dot = ctk.CTkLabel(
            header, text="●", width=12, font=ctk.CTkFont(size=theme.FONT_SIZE_LG)
        )
        self.dot.pack(side="left")
        self.status_label = ctk.CTkLabel(
            header,
            text="Idle",
            font=ctk.CTkFont(size=theme.FONT_SIZE_XL, weight="bold"),
            text_color=theme.TEXT,
        )
        self.status_label.pack(side="left", padx=(7, 0))

        self.settings_btn = ctk.CTkButton(
            header,
            text="⚙",
            width=28,
            height=24,
            corner_radius=theme.RADIUS_SM,
            fg_color="transparent",
            hover_color=theme.NEUTRAL_HOVER,
            text_color=theme.TEXT_MUTED,
            font=ctk.CTkFont(size=theme.FONT_SIZE_LG),
            command=self.open_settings,
        )
        self.settings_btn.pack(side="right")
        Tooltip(self.settings_btn, "Settings (Ctrl+,)")

        # Which macro is loaded, and whether it has unsaved changes. Without
        # this the window gave no clue what the buttons were about to replay.
        self.macro_label = ctk.CTkLabel(
            header, text="", width=160, font=small, text_color=theme.TEXT_MUTED, anchor="e"
        )
        self.macro_label.pack(side="right", padx=(8, 8))
        self._macro_tip = Tooltip(self.macro_label, "")

        # -- toolbar: the transport controls
        toolbar = ctk.CTkFrame(self._root_frame, fg_color="transparent")
        toolbar.pack(fill="x")

        self.rec_btn = ctk.CTkButton(
            toolbar, text="● Record", width=92, height=34, font=bold,
            corner_radius=theme.RADIUS_MD, command=self.toggle_record,
        )
        self.rec_btn.pack(side="left")
        self._rec_tip = Tooltip(self.rec_btn, "Start recording")

        self.stop_btn = ctk.CTkButton(
            toolbar, text="■ Stop", width=78, height=34, font=bold,
            corner_radius=theme.RADIUS_MD, command=self.request_stop,
        )
        self.stop_btn.pack(side="left", padx=6)
        self._stop_tip = Tooltip(self.stop_btn, "Stop")

        self.play_btn = ctk.CTkButton(
            toolbar, text="▶ Play", width=82, height=34, font=bold,
            corner_radius=theme.RADIUS_MD, command=self.toggle_play,
        )
        self.play_btn.pack(side="left")
        self._play_tip = Tooltip(self.play_btn, "Start playback")

        loop_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        loop_frame.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(loop_frame, text="×", font=small, text_color=theme.TEXT_FAINT).pack(side="left")
        self.loop_var = ctk.StringVar(value=self._format_loop_count(self.settings.loop_count))
        self.loop_entry = ctk.CTkEntry(
            loop_frame,
            textvariable=self.loop_var,
            width=44,
            height=30,
            corner_radius=theme.RADIUS_SM,
            font=small,
            justify="center",
            fg_color=theme.SURFACE,
            border_color=theme.BORDER,
            text_color=theme.TEXT,
            validate="key",
            validatecommand=(self.register(self._validate_loop), "%P"),
        )
        self.loop_entry.pack(side="left", padx=(4, 0))
        Tooltip(self.loop_entry, "Repetitions — 0 repeats until stopped")

        files = ctk.CTkFrame(toolbar, fg_color="transparent")
        files.pack(side="right")
        self.open_btn = self._quiet_button(files, "Open", self.open_macro)
        self.open_btn.pack(side="left")
        self.recent_btn = ctk.CTkButton(
            files,
            text="▾",
            width=18,
            height=30,
            corner_radius=theme.RADIUS_SM,
            fg_color=theme.NEUTRAL,
            hover_color=theme.NEUTRAL_HOVER,
            text_color=theme.TEXT_MUTED,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            command=self._show_recent_menu,
        )
        self.recent_btn.pack(side="left", padx=(2, 0))
        self.save_btn = self._quiet_button(files, "Save", self.save_macro)
        self.save_btn.pack(side="left", padx=(6, 0))
        Tooltip(self.open_btn, "Open a macro (Ctrl+O)")
        Tooltip(self.recent_btn, "Recently opened")
        Tooltip(self.save_btn, "Save the macro (Ctrl+S)")

        # -- progress: only on screen while there is something to measure
        self.progress = ctk.CTkProgressBar(
            self._root_frame, height=4, corner_radius=2, fg_color=theme.SURFACE_SUNKEN
        )
        self.progress.set(0)
        self._progress_visible = False

        # -- footer: the running totals
        self.info_label = ctk.CTkLabel(
            self._root_frame,
            text="",
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT_MUTED,
            anchor="w",
        )
        self.info_label.pack(fill="x", pady=(8, 0))

        self.banner = Banner(self._root_frame)

    def _quiet_button(self, master: ctk.CTkFrame, text: str, command) -> ctk.CTkButton:
        """A secondary button.

        Worded rather than pictured: the emoji the toolbar used to carry
        rendered as a colour glyph on Windows, an outline on Linux and nothing
        at all where the font was missing.
        """
        return ctk.CTkButton(
            master,
            text=text,
            width=52,
            height=30,
            corner_radius=theme.RADIUS_SM,
            fg_color=theme.NEUTRAL,
            hover_color=theme.NEUTRAL_HOVER,
            text_color=theme.TEXT,
            font=ctk.CTkFont(size=theme.FONT_SIZE_MD),
            command=command,
        )

    def _bind_shortcuts(self) -> None:
        self.bind("<Control-o>", lambda _event: self.open_macro())
        self.bind("<Control-s>", lambda _event: self.save_macro())
        self.bind("<Control-comma>", lambda _event: self.open_settings())
        # Enter in the loop field is the natural way to say "go".
        self.loop_entry.bind("<Return>", lambda _event: self.toggle_play())
        self.loop_entry.bind("<FocusOut>", lambda _event: self._normalize_loop_field())

    # -- loop field -----------------------------------------------------
    @staticmethod
    def _format_loop_count(count: int) -> str:
        return "∞" if count <= 0 else str(count)

    @staticmethod
    def _validate_loop(proposed: str) -> bool:
        """Only let plausible values into the field.

        Rejecting the keystroke is friendlier than accepting ``12a`` and
        explaining afterwards that it was played once instead.
        """
        return proposed in ("", "∞") or (proposed.isdigit() and len(proposed) <= 5)

    def _normalize_loop_field(self) -> None:
        text = self.loop_var.get().strip()
        if not text:
            self.loop_var.set("1")
            return
        if text.isdigit() and int(text) == 0:
            # Show what zero actually means, now that it is committed.
            self.loop_var.set("∞")

    # -- recent files ---------------------------------------------------
    def _show_recent_menu(self) -> None:
        # Reused rather than rebuilt: a fresh menu on every click would stay
        # alive as a child of the window until the application closed.
        menu = self._recent_menu
        menu.delete(0, "end")
        recent = [path for path in self.settings.recent_files if Path(path).is_file()]
        if recent:
            for path in recent:
                menu.add_command(
                    label=Path(path).name,
                    command=lambda target=path: self._load_path(target),
                )
        else:
            menu.add_command(label="No recent macros", state="disabled")
        try:
            menu.tk_popup(
                self.recent_btn.winfo_rootx(),
                self.recent_btn.winfo_rooty() + self.recent_btn.winfo_height(),
            )
        finally:
            menu.grab_release()

    # -- rendering ------------------------------------------------------
    def _view(self) -> ViewModel:
        progress: float | None = None
        detail = ""
        if self._state is AppState.COUNTDOWN:
            progress = self._countdown_progress
        elif self._state is AppState.PLAYING:
            snapshot = self.player.progress()
            progress = snapshot.fraction
            # "loop 1/1" would be noise on the common single run.
            if snapshot.loops != 1:
                detail = format_loop_label(snapshot.loop, snapshot.loops)

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
            countdown_total=self._countdown_total,
            pulse=self._pulse,
            progress=progress,
            detail=detail,
        )

    def _render(self) -> None:
        view = self._view()
        self.dot.configure(text_color=view.dot)
        self.status_label.configure(text=view.status)
        self.info_label.configure(text=view.info)
        for button, spec in (
            (self.rec_btn, view.record),
            (self.stop_btn, view.stop),
            (self.play_btn, view.play),
        ):
            button.configure(
                state=spec.tk_state,
                fg_color=spec.color,
                hover_color=spec.hover,
                text_color=theme.TEXT_ON_ACCENT if spec.enabled else theme.TEXT_FAINT,
            )
        self.loop_entry.configure(state="normal" if view.loop_enabled else "disabled")
        self._render_macro_name()
        self._render_progress(view)

    def _render_macro_name(self) -> None:
        name = self.macro.name if self.macro.events else ""
        shown = elide(name, self.MAX_NAME_CHARS)
        marker = " •" if self._dirty else ""
        self.macro_label.configure(text=f"{shown}{marker}")
        hint = "Unsaved changes" if self._dirty else ""
        if shown != name:
            hint = f"{name} — {hint}" if hint else name
        self._macro_tip.set_text(hint)

    def _render_progress(self, view: ViewModel) -> None:
        if view.progress is None:
            if self._progress_visible:
                self.progress.pack_forget()
                self._progress_visible = False
            return
        self.progress.configure(progress_color=view.progress_color)
        self.progress.set(max(0.0, min(1.0, view.progress)))
        if not self._progress_visible:
            self.progress.pack(fill="x", pady=(8, 0), before=self.info_label)
            self._progress_visible = True

    def _notify(self, message: str, severity: str = "info") -> None:
        """Show a message in the banner row, below the controls.

        Unlike the status text it replaces, this cannot hide what the
        application is doing, and errors stay until they are dismissed.
        """
        self.banner.show(message, severity)

    # -- pulse ----------------------------------------------------------
    def _start_pulse(self) -> None:
        self._stop_pulse(reset=False)
        self._pulse_index = 0
        self._pulse_tick()

    def _stop_pulse(self, *, reset: bool = True) -> None:
        if self._pulse_id is not None:
            self.after_cancel(self._pulse_id)
            self._pulse_id = None
        if reset:
            self._pulse = 0.0
            self.dot.configure(text_color=self._view().dot)

    def _pulse_tick(self) -> None:
        if self._state is AppState.IDLE:
            self._stop_pulse()
            return
        self._pulse = self.PULSE_STEPS[self._pulse_index % len(self.PULSE_STEPS)]
        self._pulse_index += 1
        self.dot.configure(text_color=self._view().dot)
        self._pulse_id = self.after(self.PULSE_INTERVAL_MS, self._pulse_tick)

    # -- intents --------------------------------------------------------
    def toggle_record(self) -> None:
        if self._state is AppState.IDLE:
            if not self._confirm_discard("Recording again will replace it."):
                return
            self._start_countdown("record")
        elif self._state is AppState.RECORDING:
            self._finish_recording()

    def toggle_play(self) -> None:
        if self._state is AppState.IDLE:
            if self.macro.events:
                self._start_countdown("play")
            else:
                self._notify("Nothing recorded yet", "warning")
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
        total = self.settings.countdown_seconds
        if total <= 0:
            self._begin_pending()
            return
        self._countdown_total = total
        self._countdown = total
        self._countdown_deadline = time.monotonic() + total
        self._countdown_progress = 0.0
        self._state = AppState.COUNTDOWN
        self._start_pulse()
        self._render()
        self._countdown_id = self.after(self.COUNTDOWN_INTERVAL_MS, self._tick_countdown)

    def _tick_countdown(self) -> None:
        remaining = self._countdown_deadline - time.monotonic()
        if remaining <= 0:
            self._countdown_id = None
            self._countdown = 0
            self._countdown_progress = None
            self._stop_pulse()
            self._begin_pending()
            return
        self._countdown = math.ceil(remaining)
        self._countdown_progress = 1.0 - (remaining / self._countdown_total)
        self._render()
        self._countdown_id = self.after(self.COUNTDOWN_INTERVAL_MS, self._tick_countdown)

    def _cancel_countdown(self) -> None:
        if self._countdown_id is not None:
            self.after_cancel(self._countdown_id)
            self._countdown_id = None
        self._pending = None
        self._countdown_progress = None
        self._state = AppState.IDLE
        self._stop_pulse()
        self._render()

    def _begin_pending(self) -> None:
        if self._pending == "record":
            self._begin_recording()
        else:
            self._begin_playback()

    # -- recording ------------------------------------------------------
    def _begin_recording(self) -> None:
        self.recorder.set_excluded_keys(self.settings.reserved_keys)
        try:
            self.recorder.start()
        except RuntimeError as exc:
            self._state = AppState.IDLE
            self._render()
            logger.error("Could not start recording: %s", exc)
            self._notify(f"{exc}. {permission_hint()}", "error")
            return
        self._state = AppState.RECORDING
        self._start_pulse()
        self._render()
        self._poll_recording()

    def _poll_recording(self) -> None:
        """Refresh the event count while recording.

        Polling rather than reacting to every captured event: a fast mouse
        produces hundreds of events a second, and marshalling each one onto the
        UI thread would flood it for no visible benefit.
        """
        if self._state is not AppState.RECORDING:
            self._recording_poll_id = None
            return
        self.info_label.configure(text=self._view().info)
        self._recording_poll_id = self.after(self.RECORDING_POLL_MS, self._poll_recording)

    def _finish_recording(self) -> None:
        if self._recording_poll_id is not None:
            self.after_cancel(self._recording_poll_id)
            self._recording_poll_id = None
        events = self.recorder.stop()
        self.macro = Macro(name=self.macro.name, events=events)
        self._state = AppState.IDLE
        self._stop_pulse()
        if events:
            self._dirty = True
        self._render()
        if events:
            self._notify(f"Recorded {len(events)} events", "success")

    # -- playback -------------------------------------------------------
    def _begin_playback(self) -> None:
        loops, warning = parse_loop_count(self.loop_var.get())
        self.settings.loop_count = loops
        self.player.set_events(self.macro.events)
        self.player.set_loop(loops, delay=self.settings.loop_delay)
        if not self.player.start():
            self._state = AppState.IDLE
            self._render()
            self._notify("Nothing to play", "warning")
            return
        self._state = AppState.PLAYING
        self._start_pulse()
        self._render()
        self._poll_playback()
        if warning:
            self._notify(warning, "warning")

    def _poll_playback(self) -> None:
        """Read the playback position and move the bar.

        Only the two widgets that change are touched: a full re-render sixteen
        times a second would reconfigure every button for nothing.
        """
        if self._state is not AppState.PLAYING:
            self._playback_poll_id = None
            return
        view = self._view()
        self.info_label.configure(text=view.info)
        self._render_progress(view)
        self._playback_poll_id = self.after(self.PLAYBACK_POLL_MS, self._poll_playback)

    def _stop_playback_poll(self) -> None:
        if self._playback_poll_id is not None:
            self.after_cancel(self._playback_poll_id)
            self._playback_poll_id = None

    def _finish_playback(self) -> None:
        self._stop_playback_poll()
        self.player.stop()
        self._state = AppState.IDLE
        self._stop_pulse()
        self._render()

    def _on_playback_complete(self) -> None:
        self._stop_playback_poll()
        self._state = AppState.IDLE
        self._stop_pulse()
        self._render()
        self._notify("Playback complete", "success")

    # -- settings -------------------------------------------------------
    def open_settings(self) -> None:
        if self._settings_dialog is not None and self._settings_dialog.winfo_exists():
            self._settings_dialog.focus()
            return
        # Hand the keyboard over completely: with the global hotkeys still
        # armed, pressing F9 to rebind it would also start a recording.
        self.hotkeys.stop()
        self._settings_dialog = SettingsDialog(
            self,
            self.settings,
            on_change=self._apply_setting,
            on_close=self._on_settings_closed,
        )

    def _apply_setting(self, field: str) -> None:
        """Make a settings change visible without waiting for a restart."""
        if field == "appearance_mode":
            ctk.set_appearance_mode(self.settings.appearance_mode)
        elif field == "always_on_top":
            self.attributes("-topmost", self.settings.always_on_top)
        elif field == "hotkeys":
            self.recorder.set_excluded_keys(self.settings.reserved_keys)
            self._refresh_hotkey_tooltips()
        elif field == "suppress_hotkeys":
            self._rebuild_hotkey_listener()

    def _on_settings_closed(self) -> None:
        self._settings_dialog = None
        self.settings = self.settings.normalized()
        if not self.hotkeys.apply(self.settings.hotkeys):
            self._notify(permission_hint(), "error")
        self.settings.save()
        self._render()

    def _rebuild_hotkey_listener(self) -> None:
        """Swap the listener when suppression is switched on or off."""
        self.hotkeys.stop()
        listener, self.hotkeys_suppressed = create_listener(suppress=self.settings.suppress_hotkeys)
        self.hotkeys = HotkeyManager(
            listener,
            on_error=lambda message: self.after(
                0, self._notify, f"Hotkeys unavailable: {message}", "error"
            ),
        )
        self.hotkeys.register("record", lambda: self.after(0, self.toggle_record))
        self.hotkeys.register("play", lambda: self.after(0, self.toggle_play))
        self.hotkeys.register("stop", lambda: self.after(0, self.request_stop))
        if self.settings.suppress_hotkeys and not self.hotkeys_suppressed:
            self._notify("Hotkey suppression needs Windows and 'keyboard'", "warning")

    def _refresh_hotkey_tooltips(self) -> None:
        bindings = self.settings.hotkeys
        self._rec_tip.set_text(f"Start recording ({describe_binding(bindings['record'])})")
        self._play_tip.set_text(f"Start playback ({describe_binding(bindings['play'])})")
        self._stop_tip.set_text(f"Stop ({describe_binding(bindings['stop'])})")

    # -- files ----------------------------------------------------------
    def _initial_dir(self) -> str:
        return self.settings.last_directory or str(Path.home())

    def save_macro(self) -> bool:
        """Save the macro. Returns whether it ended up on disk."""
        if not self.macro.events:
            self._notify("Nothing to save yet", "warning")
            return False
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=self._initial_dir(),
            initialfile=f"{self.macro.name}.json",
            parent=self,
        )
        if not path:
            return False
        try:
            saved = save_macro(self.macro, path)
        except MacroFileError as exc:
            self._notify(str(exc), "error")
            return False
        self.settings.last_directory = str(saved.parent)
        self.settings.remember_file(saved)
        self.macro.name = saved.stem
        self._dirty = False
        self._render()
        self._notify(f"Saved to {elide(saved.name, 40)}", "success")
        return True

    def open_macro(self) -> None:
        if not self._confirm_discard("Opening a macro will replace it."):
            return
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")], initialdir=self._initial_dir(), parent=self
        )
        if path:
            self._load_path(path, confirmed=True)

    def _load_path(self, path: str, *, confirmed: bool = False) -> None:
        """Load ``path``, asking about unsaved work unless that already happened.

        Choosing to discard does not clear the flag -- the recording is still
        unsaved until something replaces it -- so without ``confirmed`` the
        question would be asked a second time on the way through here.
        """
        if not confirmed and not self._confirm_discard("Opening a macro will replace it."):
            return
        try:
            macro = load_macro(path)
        except MacroFileError as exc:
            self._notify(str(exc), "error")
            return
        self.macro = macro
        self._dirty = False
        self.settings.last_directory = str(Path(path).parent)
        self.settings.remember_file(path)
        self._render()
        self._notify(f"Loaded {len(macro.events)} events", "success")

    def _confirm_discard(self, consequence: str) -> bool:
        """Ask before an unsaved recording would be lost.

        Returns ``True`` when it is safe to carry on, either because there was
        nothing to lose or because the user dealt with it.
        """
        if not self._dirty or not self.macro.events:
            return True
        answer = messagebox.askyesnocancel(
            "Unsaved recording",
            f"The current recording has not been saved. {consequence}\n\nSave it first?",
            parent=self,
        )
        if answer is None:
            return False
        if answer:
            return self.save_macro()
        return True

    # -- shutdown -------------------------------------------------------
    def close(self) -> None:
        if self._state is AppState.RECORDING:
            self._finish_recording()
        if not self._confirm_discard("Closing will discard it."):
            return
        self.hotkeys.stop()
        if self.recorder.is_recording():
            self.recorder.stop()
        if self.player.is_playing():
            self.player.stop()
        self.settings.loop_count = parse_loop_count(self.loop_var.get())[0]
        self.settings.save()
        self.destroy()
