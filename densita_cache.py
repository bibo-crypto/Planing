"""Shared location of the last uploaded Densita' Query workbook (KG /
Densita source).

Thin wrapper over file_cache.py -- see that module for the real
implementation.
"""

from pathlib import Path
from typing import Any

from file_cache import load_file_cache, save_file_cache

_KEY = "densita"


def save_densita_cache(source_path: Path | str) -> None:
    save_file_cache(_KEY, source_path)


def load_densita_cache() -> dict[str, Any]:
    return load_file_cache(_KEY)
