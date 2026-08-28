"""file_cache.py — one generic implementation for "remember where the last
uploaded X file was" caches.

Every optional-file source in the app (Articoli, Densita' Query, Filato
Disponibile, Magazino, Lotti, Listini, ...) used to have its own near-
identical module: same JSON read/write, same try/except/log pattern, same
{source_file, source_path, loaded_at} shape, differing only in the cache
filename. This module is the one real implementation; each of those
per-source modules (articoli_cache.py, densita_cache.py, etc.) is now a
few-line wrapper that calls save_file_cache()/load_file_cache() with its
own cache_key, so every existing import and function name elsewhere in
the app keeps working unchanged -- only the duplicated logic moved.

Add a new cached source by adding a new small wrapper module in the same
shape as the existing ones, not by extending this file.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from utils import APP_DATA_DIR, logger
from path_manager import canonical_key

CACHE_DIR = APP_DATA_DIR / "settings"


def _cache_path(cache_key: str) -> Path:
    """Resolve all callers to the same canonical source cache."""
    return CACHE_DIR / f"{canonical_key(cache_key)}_file_cache.json"


def save_file_cache(cache_key: str, source_path: Path | str, extra: dict[str, Any] | None = None) -> None:
    """Persist which file was last used for *cache_key*. *extra*, if given,
    is merged into the saved payload (e.g. magazino_cache's optional
    pre-summarized dataframe) -- most callers don't need it."""
    cache_key = canonical_key(cache_key)
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        payload: dict[str, Any] = {
            "source_file": Path(source_path).name,
            "source_path": str(source_path),
            "loaded_at": datetime.now().isoformat(timespec="seconds"),
        }
        if extra:
            payload.update(extra)
        _cache_path(cache_key).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save %s file cache: %s", cache_key, exc)


def load_file_cache(cache_key: str) -> dict[str, Any]:
    cache_key = canonical_key(cache_key)
    empty = {"source_file": "", "source_path": "", "loaded_at": ""}
    try:
        path = _cache_path(cache_key)
        if path.is_file():
            value = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(value, dict):
                return value
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load %s file cache: %s", cache_key, exc)
    return empty
