"""
main.py — Entry point for Planing (the Delta Dyeing PDF-to-Excel Converter).

Run with:
    python main.py
"""

from __future__ import annotations

import os
from pathlib import Path
import traceback


def _write_startup_error() -> Path | None:
    """Persist a frozen/source startup traceback where the user can retrieve it."""
    try:
        if getattr(os.sys, "frozen", False):
            base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "Planing"
        else:
            base = Path(__file__).resolve().parent / "logs"
        base.mkdir(parents=True, exist_ok=True)
        path = base / "startup_error.log"
        path.write_text(traceback.format_exc(), encoding="utf-8")
        return path
    except Exception:
        return None


def main() -> None:
    from gui import ConverterApp

    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        error_path = _write_startup_error()
        try:
            from tkinter import messagebox

            location = f"\n\nسجل الخطأ: {error_path}" if error_path else ""
            messagebox.showerror(
                "Planing - Startup error",
                "تعذر تشغيل البرنامج. راجع سجل الخطأ للحصول على التفاصيل." + location,
            )
        except Exception:
            # The startup error has already been persisted; there may be no
            # usable Tk environment in which to display a dialog.
            pass
