"""
main.py — Entry point for Planing (the Delta Dyeing PDF-to-Excel Converter).

Run with:
    python main.py
"""

from __future__ import annotations


def main() -> None:
    from gui import ConverterApp

    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
