"""Shared location of the last uploaded Articoli.xlsx, kept so Biglietti can
re-read its 'Marca' column for the ticket's Titolo -- Situazione's own
Articoli upload persists a *different* column (TITOLO) into situazione_db,
so this is a separate cache rather than reusing that path.

Thin wrapper over file_cache.py -- see that module for the real
implementation.
"""

from pathlib import Path
from typing import Any

from file_cache import load_file_cache, save_file_cache

_KEY = "articoli_marca"


def save_articoli_cache(source_path: Path | str) -> None:
    save_file_cache(_KEY, source_path)


def load_articoli_cache() -> dict[str, Any]:
    return load_file_cache(_KEY)
