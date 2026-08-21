"""Shared location of the last uploaded Magazino (Color Tube / VMM22) workbook.

This is a *different* file from the raw-yarn Magazino used by
magazino_cache.py / the "Magazino Filato" tab: that one tracks raw-yarn
stock by Articolo; this one carries Color Tube + kg-per-cone by Partita,
used only to fill the Biglietti "Color Tube" and "VMM22" columns.
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from utils import APP_DATA_DIR, logger

CACHE_FILE = APP_DATA_DIR / "settings" / "vmm_magazino_file_cache.json"


def save_vmm_magazino_cache(source_path: Path | str) -> None:
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "source_file": Path(source_path).name,
            "source_path": str(source_path),
            "loaded_at": datetime.now().isoformat(timespec="seconds"),
        }
        CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save Magazino (Color Tube) file cache: %s", exc)


def load_vmm_magazino_cache() -> dict[str, Any]:
    empty = {"source_file": "", "source_path": "", "loaded_at": ""}
    try:
        if CACHE_FILE.is_file():
            value = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if isinstance(value, dict):
                return value
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load Magazino (Color Tube) file cache: %s", exc)
    return empty
