"""Excel I/O helpers with user-facing file-lock errors."""
from __future__ import annotations

from pathlib import Path


def safe_save_workbook(workbook, path: Path) -> None:
    """Save an openpyxl workbook and translate Windows lock errors."""
    try:
        workbook.save(path)
    except PermissionError as exc:
        raise RuntimeError(
            f"Cannot save '{Path(path).name}'. The Excel file is open or locked. "
            "Close it in Excel (including any hidden Excel window) and try again."
        ) from exc
    except OSError as exc:
        message = str(exc).lower()
        if "permission" in message or "being used" in message or "locked" in message:
            raise RuntimeError(
                f"Cannot save '{Path(path).name}' because it is open or locked. "
                "Close the Excel file and try again."
            ) from exc
        raise
