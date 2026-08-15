"""A small hover label, used to name the icon buttons and their hotkeys.

CustomTkinter has no tooltip of its own, and the alternative -- writing the
hotkey into every button's caption -- would not fit the toolbar. This is the
smallest thing that works: one borrowed toplevel per tooltip, created on hover
and destroyed on leave.
"""

from __future__ import annotations

import tkinter as tk

import customtkinter as ctk

from . import theme

__all__ = ["Tooltip"]


class Tooltip:
    """Show ``text`` next to ``widget`` while the pointer rests on it."""

    #: How long the pointer must rest before the tooltip appears.
    DELAY_MS = 450
    #: Offset from the widget's bottom-left corner.
    OFFSET_X = 0
    OFFSET_Y = 6

    def __init__(self, widget: tk.Misc, text: str) -> None:
        self._widget = widget
        self._text = text
        self._after_id: str | None = None
        self._window: tk.Toplevel | None = None

        # ``add="+"`` so a tooltip never displaces a binding the widget needs.
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")
        widget.bind("<Destroy>", self._hide, add="+")

    def set_text(self, text: str) -> None:
        """Change the caption, e.g. after a hotkey has been rebound."""
        self._text = text
        if self._window is not None:
            self._hide()

    # -- internals ------------------------------------------------------
    def _schedule(self, _event: object = None) -> None:
        self._cancel()
        if self._text:
            self._after_id = self._widget.after(self.DELAY_MS, self._show)

    def _cancel(self) -> None:
        if self._after_id is not None:
            try:
                self._widget.after_cancel(self._after_id)
            except tk.TclError:  # the widget went away mid-hover
                pass
            self._after_id = None

    def _show(self) -> None:
        self._after_id = None
        if self._window is not None or not self._text:
            return
        try:
            x = self._widget.winfo_rootx() + self.OFFSET_X
            y = self._widget.winfo_rooty() + self._widget.winfo_height() + self.OFFSET_Y
        except tk.TclError:
            return

        window = tk.Toplevel(self._widget)
        window.wm_overrideredirect(True)
        window.wm_geometry(f"+{x}+{y}")
        try:
            window.attributes("-topmost", True)
        except tk.TclError:
            pass  # not every window manager allows this; the tip still shows

        frame = ctk.CTkFrame(
            window,
            corner_radius=theme.RADIUS_SM,
            fg_color=theme.SURFACE,
            border_width=1,
            border_color=theme.BORDER,
        )
        frame.pack()
        ctk.CTkLabel(
            frame,
            text=self._text,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            text_color=theme.TEXT_MUTED,
        ).pack(padx=8, pady=3)
        self._window = window

    def _hide(self, _event: object = None) -> None:
        self._cancel()
        if self._window is not None:
            try:
                self._window.destroy()
            except tk.TclError:
                pass
            self._window = None
