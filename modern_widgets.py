"""Small native Tk widgets used for the application's modern controls."""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk


class RoundedButton(tk.Canvas):
    """A compact pill-shaped button with the same command API as ttk.Button."""

    def __init__(self, master, text="", command=None, width=None, **kwargs):
        self._text = str(text)
        self._command = command
        self._state = "normal"
        self._font = tkfont.Font(family="Segoe UI", size=9)
        self._surface = self._frame_background(master)
        self._height = 34
        text_width = self._font.measure(self._text)
        requested = int(width) if width is not None else 0
        if requested and requested <= 40:
            requested = requested * 8 + 24
        self._width = max(96, requested, text_width + 30)

        super().__init__(
            master,
            width=self._width,
            height=self._height,
            highlightthickness=0,
            bd=0,
            relief="flat",
            background=self._surface,
            takefocus=True,
            cursor="hand2",
            **{key: value for key, value in kwargs.items() if key not in {"style"}},
        )
        self._hovered = False
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Return>", self._on_click)
        self.bind("<space>", self._on_click)
        self._draw()

    @staticmethod
    def _frame_background(master):
        try:
            return master.cget("background")
        except tk.TclError:
            return ttk.Style(master).lookup("TFrame", "background") or "#f0f0f0"

    def _draw(self):
        self.delete("all")
        disabled = self._state == "disabled"
        if disabled:
            fill, outline, foreground = "#eef2f6", "#d6dee8", "#94a3b8"
        elif self._hovered:
            fill, outline, foreground = "#e2e8f0", "#8fa8c0", "#16324f"
        else:
            fill, outline, foreground = "#f8fafc", "#b8c6d6", "#16324f"

        radius = self._height // 2
        self.create_arc(1, 1, self._height - 1, self._height - 1,
                        start=90, extent=180, fill=fill, outline=outline)
        self.create_rectangle(radius, 1, self._width - radius, self._height - 1,
                              fill=fill, outline=outline)
        self.create_arc(self._width - self._height + 1, 1, self._width - 1, self._height - 1,
                        start=270, extent=180, fill=fill, outline=outline)
        self.create_text(self._width // 2, self._height // 2, text=self._text,
                         fill=foreground, font=self._font)

    def _on_enter(self, _event):
        if self._state != "disabled":
            self._hovered = True
            self._draw()

    def _on_leave(self, _event):
        self._hovered = False
        self._draw()

    def _on_click(self, _event=None):
        if self._state != "disabled" and self._command:
            self._command()

    def configure(self, cnf=None, **kwargs):
        options = dict(cnf or {})
        options.update(kwargs)
        if "text" in options:
            self._text = str(options.pop("text"))
        if "command" in options:
            self._command = options.pop("command")
        if "state" in options:
            self._state = str(options.pop("state"))
        if options:
            super().configure(**options)
        self._draw()

    config = configure

