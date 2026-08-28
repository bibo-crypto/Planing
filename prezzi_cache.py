"""Shared location of the last uploaded Listini (price list) export.

Thin wrapper over file_cache.py -- see that module for the real
implementation.
"""

from pathlib import Path
from typing import Any

from file_cache import load_file_cache, save_file_cache

_KEY = "prezzi"


def save_prezzi_cache(source_path: Path | str) -> None:
    save_file_cache(_KEY, source_path)


def load_prezzi_cache() -> dict[str, Any]:
    return load_file_cache(_KEY)
