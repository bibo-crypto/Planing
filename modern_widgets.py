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
        button_font = kwargs.pop("font", ("Segoe UI", 10, "bold"))
        self._font = tkfont.Font(font=button_font)
        self._surface = self._frame_background(master)
        # Keep the custom buttons compact inside dense form sections. A
        # taller default makes the Ordini ELVY page overflow its notebook.
        self._height = int(kwargs.pop("height", 32))
        text_width = self._font.measure(self._text)
        requested = int(width) if width is not None else 0
        if requested and requested <= 40:
            requested = requested * 9 + 40
        self._width = max(130, requested, text_width + 40)

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
        self.bind("<Configure>", self._on_configure)
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
        draw_width = max(self._width, self.winfo_width())
        draw_height = max(self._height, self.winfo_height())
        disabled = self._state == "disabled"
        if disabled:
            fill, outline, foreground = "#dbe3ec", "#c4cfdb", "#7b8794"
        elif self._hovered:
            fill, outline, foreground = "#1d4ed8", "#1e40af", "#ffffff"
        else:
            fill, outline, foreground = "#f8fafc", "#9fb2c5", "#16324f"

        # A clean rectangular control keeps the hover/active area exactly
        # aligned with the visible button bounds.
        self.create_rectangle(1, 1, draw_width - 1, draw_height - 1,
                              fill=fill, outline=outline, width=1)
        self.create_text(draw_width // 2, draw_height // 2 - 1, text=self._text,
                         fill=foreground, font=self._font)

    def _on_configure(self, _event):
        # Grid/pack may stretch the Canvas beyond its requested width.  Redraw
        # to that exact widget boundary so hover never activates in a blank
        # area beside the visible button.
        self._draw()

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
