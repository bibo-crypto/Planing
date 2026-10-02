"""Safe lifecycle helpers for Tkinter secondary windows."""
from __future__ import annotations

import tkinter as tk
from collections.abc import Callable


def bind_escape_close(
    window: tk.Misc,
    close_callback: Callable[[], None] | None = None,
    parent: tk.Misc | None = None,
    *,
    critical: bool = False,
) -> Callable[[], None]:
    """Make WM close and Escape use the same idempotent cleanup path.

    ``critical`` is reserved for update/download dialogs where Escape must not
    interrupt the operation. The helper is deliberately tolerant of windows
    that were already destroyed by the native window manager.
    """
    closed = False

    def close() -> None:
        nonlocal closed
        if closed:
            return
        closed = True
        try:
            if close_callback is not None:
                close_callback()
            elif window.winfo_exists():
                window.destroy()
        except (tk.TclError, RuntimeError):
            pass
        finally:
            if parent is not None:
                try:
                    if parent.winfo_exists():
                        parent.focus_set()
                except tk.TclError:
                    pass

    if not critical:
        window.bind("<Escape>", lambda _event: (close(), "break")[1])
        window.protocol("WM_DELETE_WINDOW", close)
    return close
