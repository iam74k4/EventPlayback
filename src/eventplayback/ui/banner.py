"""The message strip below the toolbar.

Messages used to be written over the status label, which meant that while one
was on screen there was no way to tell whether the application was recording,
playing or idle. They now get their own row, and severity decides how long
they stay: an error about missing input permissions is not something to show
for two seconds and then discard.
"""

from __future__ import annotations

import customtkinter as ctk

from . import theme

__all__ = ["Banner"]

_ICONS = {"info": "i", "success": "✓", "warning": "!", "error": "!"}

_COLORS = {
    "info": theme.INFO,
    "success": theme.SUCCESS,
    "warning": theme.WARNING,
    "error": theme.DANGER,
}

#: How long each severity stays before it clears itself. ``None`` means the
#: message waits for the user to dismiss it.
_DISMISS_MS: dict[str, int | None] = {
    "info": 2500,
    "success": 2500,
    "warning": 5000,
    "error": None,
}


class Banner(ctk.CTkFrame):
    """A dismissable, severity-coloured message row."""

    def __init__(self, master: ctk.CTkBaseClass | ctk.CTk) -> None:
        super().__init__(master, corner_radius=theme.RADIUS_SM, fg_color=theme.SURFACE, height=28)
        self._token = 0

        self._icon = ctk.CTkLabel(
            self, text="", width=14, font=ctk.CTkFont(size=theme.FONT_SIZE_MD, weight="bold")
        )
        self._icon.pack(side="left", padx=(8, 0))

        self._label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=theme.FONT_SIZE_MD), anchor="w", justify="left"
        )
        self._label.pack(side="left", fill="x", expand=True, padx=(6, 4), pady=4)

        self._close = ctk.CTkButton(
            self,
            text="✕",
            width=20,
            height=20,
            corner_radius=theme.RADIUS_SM,
            fg_color="transparent",
            hover_color=theme.NEUTRAL_HOVER,
            font=ctk.CTkFont(size=theme.FONT_SIZE_SM),
            command=self.hide,
        )
        self._close.pack(side="right", padx=(0, 6))

    def show(self, message: str, severity: str = "info") -> None:
        """Display ``message``. A later message always replaces an earlier one."""
        if severity not in _COLORS:
            severity = "info"
        self._token += 1
        token = self._token

        color = _COLORS[severity]
        self.configure(fg_color=theme.tint(color))
        self._icon.configure(text=_ICONS[severity], text_color=color)
        self._label.configure(text=message, text_color=theme.BANNER_TEXT[severity])
        self._close.configure(text_color=theme.BANNER_TEXT[severity])
        self.pack(fill="x", pady=(8, 0))

        timeout = _DISMISS_MS[severity]
        if timeout is not None:
            # The token stops an earlier message's timer from clearing a later
            # one that arrived while it was still on screen.
            self.after(timeout, lambda: self._expire(token))

    def hide(self) -> None:
        self._token += 1
        self.pack_forget()

    def _expire(self, token: int) -> None:
        if token == self._token:
            self.pack_forget()
