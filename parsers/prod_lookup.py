"""
prod_lookup.py
Shared cache for the raw "Produzione" export file path. Situazione's "Data
prod." upload slot and Situazione Settimana's "Produzione" upload slot both
read the same underlying ERP export, just different columns from it -- so
uploading it on either page saves the path here, and the other page reloads
from this path automatically (same pattern as dfm_lookup.py's shared DFM
cache, just storing a path instead of pre-built lookup entries, since each
page parses the columns it needs itself).
"""
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from utility.utils import APP_DATA_DIR, logger

CACHE_FILE = APP_DATA_DIR / "settings" / "prod_file_cache.json"


def save_prod_cache(source_path: Path | str) -> None:
    """Remember where the Produzione file was last uploaded from."""
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        payload: dict[str, Any] = {
            "source_file": Path(source_path).name,
            "source_path": str(source_path),
            "loaded_at": datetime.now().isoformat(timespec="seconds"),
        }
        CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save Produzione file cache: %s", exc)

    try:
        from utility.path_manager import save_source
        save_source("data_prod", source_path)
        save_source("produzione", source_path)
    except Exception:
        pass


def load_prod_cache() -> dict[str, Any]:
    """
    Load the cached Produzione file location from disk.
    Returns {"source_file": "", "source_path": "", "loaded_at": ""} if no
    cache exists yet or it can't be read.
    """
    empty: dict[str, Any] = {"source_file": "", "source_path": "", "loaded_at": ""}
    try:
        if CACHE_FILE.is_file():
            data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                sp = data.get("source_path", "")
                if sp and Path(sp).is_file():
                    return data
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load Produzione file cache: %s", exc)

    # Fallback 1: check centralized path_manager
    try:
        from utility.path_manager import source_path
        for key in ("data_prod", "produzione", "prod"):
            p = source_path(key, existing_only=True)
            if p and p.is_file():
                return {"source_file": p.name, "source_path": str(p), "loaded_at": ""}
    except Exception:
        pass

    # Fallback 2: check SQLite upload_log
    try:
        import utility.situazione_db as db
        for key in ("data_prod", "produzione"):
            info = db.get_all_uploads().get(key, {})
            fp = info.get("file_path", "")
            if fp and Path(fp).is_file():
                return {"source_file": Path(fp).name, "source_path": str(fp), "loaded_at": info.get("uploaded_at", "")}
    except Exception:
        pass

    return empty
