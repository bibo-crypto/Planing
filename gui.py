"""Backward-compatible launcher for the reorganized UI package."""

from ui.gui import ConverterApp


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()

