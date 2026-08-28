"""Centralized persistence for shared input-file locations.

All pages use the same logical source keys.  The existing thin ``*_cache``
modules remain compatible, but delegate their storage to this registry via
:file_cache.py.  Paths are stored as strings so they survive application
restarts and are validated by callers when they are consumed.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

# One canonical key per shared source.  Aliases make the intent explicit when
# a page uses a business-facing name that differs from the cache filename.
SOURCE_KEYS: dict[str, str] = {
    "dfm": "dfm",
    # Keep compatibility with the established Marca cache filename.
    "articoli": "articoli_marca",
    "densita": "densita",
    "magazino": "magazino",
    "lotti": "lotti",
    "listini": "prezzi",
    "prezzi": "prezzi",
    # Keep compatibility with prod_lookup.py's established cache filename.
    "produzione": "prod",
    "data_prod": "prod",
    "copertura": "copertura",
    "wincoint": "wincoint",
    "uscita": "uscita",
    "qualita": "qualita",
    "data_ordine": "data_ordine",
    "dispo_bagno": "dispo_bagno",
}


def canonical_key(source: str) -> str:
    """Return the stable storage key for a page/source name."""
    normalized = str(source).strip().lower().replace(" ", "_")
    return SOURCE_KEYS.get(normalized, normalized)


def save_source(source: str, path: str | Path, *, extra: dict[str, Any] | None = None) -> None:
    """Persist a source path through the one underlying cache implementation."""
    from file_cache import save_file_cache

    save_file_cache(canonical_key(source), path, extra=extra)


def load_source(source: str) -> dict[str, Any]:
    """Load a source record from the centralized path store."""
    from file_cache import load_file_cache

    return load_file_cache(canonical_key(source))


def source_path(source: str, *, existing_only: bool = False) -> Path | None:
    """Return a stored path, optionally requiring it to still exist."""
    value = load_source(source).get("source_path", "")
    if not value:
        return None
    path = Path(value)
    if existing_only and not path.is_file():
        return None
    return path


def source_records() -> dict[str, dict[str, Any]]:
    """Return all known source records, useful for diagnostics and audits."""
    return {name: load_source(name) for name in sorted(set(SOURCE_KEYS.values()))}
