"""Shared cache for the last selected LOTTI file."""

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from utils import APP_DATA_DIR, logger

CACHE_FILE = APP_DATA_DIR / "settings" / "lotti_file_cache.json"


def save_lotti_cache(source_path: Path | str) -> None:
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        payload: dict[str, Any] = {
            "source_file": Path(source_path).name,
            "source_path": str(source_path),
            "loaded_at": datetime.now().isoformat(timespec="seconds"),
        }
        CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save LOTTI file cache: %s", exc)


def load_lotti_cache() -> dict[str, Any]:
    empty: dict[str, Any] = {
        "source_file": "",
        "source_path": "",
        "loaded_at": "",
    }
    try:
        if CACHE_FILE.is_file():
            data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load LOTTI file cache: %s", exc)
    return empty
