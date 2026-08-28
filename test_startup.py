import traceback
import tkinter as tk

from gui import ConverterApp


def main() -> None:
    try:
        app = ConverterApp()
        app.after(3000, app.destroy)
        app.mainloop()
    except Exception:
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
    print("startup_smoke=ok")
