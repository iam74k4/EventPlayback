"""The settings window.

Everything in :class:`~eventplayback.services.settings.Settings` used to be
reachable only by editing JSON by hand and restarting, which meant the
countdown, the theme and the hotkeys were effectively fixed. This dialog edits
the live settings object and reports each change, so the window can apply it
straight away.
"""

from __future__ import annotations

import sys
from collections.abc import Callable

import customtkinter as ctk

from ..platform_support import IS_WINDOWS
from ..services.settings import DEFAULT_HOTKEYS, Settings
from . import theme
from .keycapture import binding_from_event, describe_binding

__all__ = ["SettingsDialog"]

_ACTION_LABELS = {"record": "Record", "play": "Play", "stop": "Stop"}


class SettingsDialog(ctk.CTkToplevel):
    """Edits ``settings`` in place, calling ``on_change`` with the field name."""

    WIDTH = 400
    HEIGHT = 330

    def __init__(
        self,
        master: ctk.CTk,
        settings: Settings,
        *,
        on_change: Callable[[str], None],
        on_close: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master)
        self._settings = settings
        self._on_change = on_change
        self._on_close = on_close
        self._capturing: str | None = None
        self._capture_funcid: str | None = None
        self._hotkey_buttons: dict[str, ctk.CTkButton] = {}

        self.title("Settings")
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.resizable(False, False)
        self.configure(fg_color=theme.WINDOW)
        self.protocol("WM_DELETE_WINDOW", self.close)

        self._build()
        self._refresh_hotkeys()

        # Sit above the main window and take the keyboard, so a captured key
        # cannot leak into whatever is behind.
        self.transient(master)
        self.after(10, self._grab)

    def _grab(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except Exception:  # a window manager that refuses grabs is not fatal
            pass

    # -- construction ---------------------------------------------------
    def _build(self) -> None:
        tabs = ctk.CTkTabview(
            self,
            fg_color=theme.SURFACE,
            segmented_button_selected_color=theme.PLAY,
            segmented_button_selected_hover_color=theme.PLAY_HOVER,
            corner_radius=theme.RADIUS_MD,
        )
        tabs.pack(fill="both", expand=True, padx=12, pady=(10, 6))
        tabs.add("General")
        tabs.add("Hotkeys")

        self._build_general(tabs.tab("General"))
        self._build_hotkeys(tabs.tab("Hotkeys"))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkButton(
            footer,
            text="Reset to defaults",
            width=140,
            height=28,
            corner_radius=theme.RADIUS_SM,
            fg_color=theme.NEUTRAL,
            hover_color=theme.NEUTRAL_HOVER,
            text_color=theme.TEXT,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            command=self._reset,
        ).pack(side="left")
        ctk.CTkButton(
            footer,
            text="Close",
            width=80,
            height=28,
            corner_radius=theme.RADIUS_SM,
            fg_color=theme.PLAY,
            hover_color=theme.PLAY_HOVER,
            text_color=theme.TEXT_ON_ACCENT,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            command=self.close,
        ).pack(side="right")

    def _row_label(self, parent: ctk.CTkFrame, text: str, row: int) -> None:
        ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(size=theme.FONT_SIZE_MD),
            text_color=theme.TEXT,
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=8)

    def _build_general(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(1, weight=1)

        self._row_label(parent, "Appearance", 0)
        self._appearance = ctk.CTkSegmentedButton(
            parent,
            values=["Dark", "Light", "System"],
            width=200,
            height=28,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            command=self._set_appearance,
        )
        self._appearance.set(self._settings.appearance_mode.capitalize())
        self._appearance.grid(row=0, column=1, sticky="e", pady=8)

        self._row_label(parent, "Always on top", 1)
        self._on_top = ctk.CTkSwitch(
            parent,
            text="",
            width=44,
            progress_color=theme.PLAY,
            command=self._set_on_top,
        )
        if self._settings.always_on_top:
            self._on_top.select()
        self._on_top.grid(row=1, column=1, sticky="e", pady=8)

        self._row_label(parent, "Countdown", 2)
        countdown = ctk.CTkFrame(parent, fg_color="transparent")
        countdown.grid(row=2, column=1, sticky="e", pady=8)
        self._countdown_value = ctk.CTkLabel(
            countdown,
            text="",
            width=44,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT_MUTED,
        )
        self._countdown_value.pack(side="right")
        self._countdown = ctk.CTkSlider(
            countdown,
            from_=0,
            to=10,
            number_of_steps=10,
            width=150,
            button_color=theme.PLAY,
            button_hover_color=theme.PLAY_HOVER,
            progress_color=theme.PLAY,
            command=self._set_countdown,
        )
        self._countdown.set(self._settings.countdown_seconds)
        self._countdown.pack(side="right", padx=(0, 8))
        self._show_countdown(self._settings.countdown_seconds)

        self._row_label(parent, "Delay between loops", 3)
        delay = ctk.CTkFrame(parent, fg_color="transparent")
        delay.grid(row=3, column=1, sticky="e", pady=8)
        self._delay_value = ctk.CTkLabel(
            delay,
            text="",
            width=44,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT_MUTED,
        )
        self._delay_value.pack(side="right")
        self._delay = ctk.CTkSlider(
            delay,
            from_=0,
            to=10,
            number_of_steps=100,
            width=150,
            button_color=theme.PLAY,
            button_hover_color=theme.PLAY_HOVER,
            progress_color=theme.PLAY,
            command=self._set_delay,
        )
        self._delay.set(min(self._settings.loop_delay, 10.0))
        self._delay.pack(side="right", padx=(0, 8))
        self._show_delay(self._settings.loop_delay)

    def _build_hotkeys(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(1, weight=1)

        for row, (action, label) in enumerate(_ACTION_LABELS.items()):
            self._row_label(parent, label, row)
            button = ctk.CTkButton(
                parent,
                text="",
                width=150,
                height=28,
                corner_radius=theme.RADIUS_SM,
                fg_color=theme.NEUTRAL,
                hover_color=theme.NEUTRAL_HOVER,
                text_color=theme.TEXT,
                font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
                command=lambda name=action: self._toggle_capture(name),
            )
            button.grid(row=row, column=1, sticky="e", pady=8)
            self._hotkey_buttons[action] = button

        self._hotkey_hint = ctk.CTkLabel(
            parent,
            text="Click a binding, then press the keys you want.",
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT_MUTED,
            wraplength=330,
            justify="left",
            anchor="w",
        )
        self._hotkey_hint.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 2))

        self._suppress = ctk.CTkSwitch(
            parent,
            text="Withhold hotkeys from other applications",
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT if IS_WINDOWS else theme.TEXT_FAINT,
            progress_color=theme.PLAY,
            command=self._set_suppress,
        )
        if self._settings.suppress_hotkeys:
            self._suppress.select()
        if not IS_WINDOWS:
            self._suppress.configure(state="disabled")
        self._suppress.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        if not IS_WINDOWS:
            ctk.CTkLabel(
                parent,
                text="Withholding needs Windows and the 'keyboard' extra.",
                font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
                text_color=theme.TEXT_FAINT,
                anchor="w",
            ).grid(row=5, column=0, columnspan=2, sticky="w")

    # -- general handlers -----------------------------------------------
    def _set_appearance(self, choice: str) -> None:
        self._settings.appearance_mode = choice.lower()
        self._on_change("appearance_mode")

    def _set_on_top(self) -> None:
        self._settings.always_on_top = bool(self._on_top.get())
        self._on_change("always_on_top")

    def _show_countdown(self, seconds: int) -> None:
        self._countdown_value.configure(text="Off" if seconds <= 0 else f"{seconds} s")

    def _set_countdown(self, value: float) -> None:
        seconds = int(round(value))
        self._settings.countdown_seconds = seconds
        self._show_countdown(seconds)
        self._on_change("countdown_seconds")

    def _show_delay(self, seconds: float) -> None:
        self._delay_value.configure(text="None" if seconds <= 0 else f"{seconds:.1f} s")

    def _set_delay(self, value: float) -> None:
        seconds = round(value, 1)
        self._settings.loop_delay = seconds
        self._show_delay(seconds)
        self._on_change("loop_delay")

    # -- hotkey handlers ------------------------------------------------
    def _refresh_hotkeys(self) -> None:
        for action, button in self._hotkey_buttons.items():
            if action == self._capturing:
                continue
            button.configure(
                text=describe_binding(self._settings.hotkeys.get(action, "")),
                fg_color=theme.NEUTRAL,
                text_color=theme.TEXT,
            )

    def _toggle_capture(self, action: str) -> None:
        if self._capturing == action:
            self._end_capture()
            return
        self._end_capture()
        self._capturing = action
        self._hotkey_buttons[action].configure(
            text="Press keys…", fg_color=theme.INFO, text_color=theme.TEXT_ON_ACCENT
        )
        self._capture_funcid = self.bind("<KeyPress>", self._on_capture_key, add="+")
        self.focus_force()

    def _end_capture(self) -> None:
        if self._capture_funcid is not None:
            self.unbind("<KeyPress>", self._capture_funcid)
            self._capture_funcid = None
        self._capturing = None
        self._refresh_hotkeys()

    def _on_capture_key(self, event: object) -> str:
        action = self._capturing
        if action is None:
            return ""
        keysym = str(getattr(event, "keysym", ""))
        try:
            state = int(getattr(event, "state", 0))
        except (TypeError, ValueError):
            state = 0

        binding = binding_from_event(keysym, state, platform=sys.platform)
        if binding is None:
            return "break"  # still only modifiers held; keep listening

        clash = next(
            (
                other
                for other, value in self._settings.hotkeys.items()
                if value == binding and other != action
            ),
            None,
        )
        if clash is not None:
            self._hotkey_hint.configure(
                text=f"{describe_binding(binding)} is already bound to {_ACTION_LABELS[clash]}.",
                text_color=theme.WARNING,
            )
            self._end_capture()
            return "break"

        self._settings.hotkeys[action] = binding
        self._hotkey_hint.configure(
            text="Bound keys are kept out of recordings.", text_color=theme.TEXT_MUTED
        )
        self._end_capture()
        self._on_change("hotkeys")
        return "break"

    def _set_suppress(self) -> None:
        self._settings.suppress_hotkeys = bool(self._suppress.get())
        self._on_change("suppress_hotkeys")

    # -- lifecycle ------------------------------------------------------
    def _reset(self) -> None:
        defaults = Settings()
        self._settings.appearance_mode = defaults.appearance_mode
        self._settings.always_on_top = defaults.always_on_top
        self._settings.countdown_seconds = defaults.countdown_seconds
        self._settings.loop_delay = defaults.loop_delay
        self._settings.suppress_hotkeys = defaults.suppress_hotkeys
        self._settings.hotkeys = dict(DEFAULT_HOTKEYS)

        self._appearance.set(defaults.appearance_mode.capitalize())
        if defaults.always_on_top:
            self._on_top.select()
        else:
            self._on_top.deselect()
        self._countdown.set(defaults.countdown_seconds)
        self._show_countdown(defaults.countdown_seconds)
        self._delay.set(defaults.loop_delay)
        self._show_delay(defaults.loop_delay)
        self._suppress.deselect()
        self._refresh_hotkeys()

        for field in (
            "appearance_mode",
            "always_on_top",
            "countdown_seconds",
            "loop_delay",
            "suppress_hotkeys",
            "hotkeys",
        ):
            self._on_change(field)

    def close(self) -> None:
        self._end_capture()
        try:
            self.grab_release()
        except Exception:
            pass
        if self._on_close is not None:
            self._on_close()
        self.destroy()
