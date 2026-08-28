"""Shared location of the last raw-yarn Magazino export.

Thin wrapper over file_cache.py -- see that module for the real
implementation. The one thing this source does that the others don't:
an optional pre-summarized dataframe gets saved alongside the path.
"""

import json
from pathlib import Path
from typing import Any

from file_cache import load_file_cache, save_file_cache
from utils import logger

_KEY = "magazino"


def save_magazino_cache(source_path: Path | str, summary=None) -> None:
    extra = None
    if summary is not None:
        try:
            # to_json converts pandas/numpy scalar values to JSON-safe types.
            extra = {"summary_rows": json.loads(summary.to_json(orient="records"))}
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not include Magazino summary in cache: %s", exc)
    save_file_cache(_KEY, source_path, extra=extra)


def load_magazino_cache() -> dict[str, Any]:
    return load_file_cache(_KEY)
