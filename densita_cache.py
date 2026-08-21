"""Shared location of the last uploaded Densita' Query workbook.

Same pattern as magazino_cache.py / prezzi_cache.py: only the file *path* is
remembered (not its content), so the app can silently reload it next time
without asking the user to re-browse, and so it stays independent of which
tab last triggered the upload.
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from utils import APP_DATA_DIR, logger

CACHE_FILE = APP_DATA_DIR / "settings" / "densita_file_cache.json"


def save_densita_cache(source_path: Path | str) -> None:
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "source_file": Path(source_path).name,
            "source_path": str(source_path),
            "loaded_at": datetime.now().isoformat(timespec="seconds"),
        }
        CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save Densita' Query file cache: %s", exc)


def load_densita_cache() -> dict[str, Any]:
    empty = {"source_file": "", "source_path": "", "loaded_at": ""}
    try:
        if CACHE_FILE.is_file():
            value = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if isinstance(value, dict):
                return value
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load Densita' Query file cache: %s", exc)
    return empty
