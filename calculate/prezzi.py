"""Listini loading and price lookup business rules."""
from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from utility.utils import clean_text

REQUIRED_COLUMNS = [
    "DATAFINEVAL", "DATAINIZIOVAL", "CLARTICOLO", "DESCRIZARTICOLOLI", "CLCOLORE",
    "CLDESCR", "LIVELLOLPZ", "PREZZOLPZ",
]
DISPLAY_COLUMNS = [
    "CLARTICOLO", "DESCRIZARTICOLOLI", "CLCOLORE", "CLDESCR", "LIVELLOLPZ", "PREZZOLPZ",
]
HEADERS = {
    "CLARTICOLO": "Articolo",
    "DESCRIZARTICOLOLI": "Descrizione Articolo",
    "CLCOLORE": "Codice Colore",
    "CLDESCR": "Descrizione Colore",
    "LIVELLOLPZ": "Livello",
    "PREZZOLPZ": "Prezzo",
}

# Listini is commonly a legacy .xls export (forces the slow pure-Python
# xlrd engine) and gets re-requested independently by Situazione, Prezzi,
# Biglietti, and the cross-tab background sync within the same session --
# memoize the parsed result per (path, mtime, size) so it's only actually
# re-read from disk when the file has genuinely changed. See the matching
# comment in parsers/situazione_loaders.py._read_raw for the full reasoning.
_PREZZI_CACHE: dict[tuple, tuple] = {}


def _format_codice(value) -> str:
    if pd.isna(value):
        return ""
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return clean_text(value)
    if as_float.is_integer():
        return str(int(as_float))
    return str(as_float)


def load_prezzi(path: str | Path) -> tuple[pd.DataFrame | None, list[str]]:
    cache_key = None
    try:
        stat = os.stat(path)
        cache_key = (str(path), stat.st_mtime_ns, stat.st_size)
    except OSError:
        pass
    if cache_key is not None and cache_key in _PREZZI_CACHE:
        return _PREZZI_CACHE[cache_key]

    result = _load_prezzi_uncached(path)
    if cache_key is not None and result[0] is not None:
        _PREZZI_CACHE[cache_key] = result
    return result


def _load_prezzi_uncached(path: str | Path) -> tuple[pd.DataFrame | None, list[str]]:
    try:
        df = pd.read_excel(path)
    except Exception as exc:  # noqa: BLE001
        return None, [f"Couldn't read the file: {exc}"]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return None, [f"Missing columns in the Listini file: {', '.join(missing)}"]

    df = df[df["DATAFINEVAL"].isna()]
    df = df[df["CLCOLORE"].notna()]
    start_date = pd.to_datetime(df["DATAINIZIOVAL"], format="%d/%m/%Y", errors="coerce")
    df = df.assign(_start=start_date).sort_values("_start", na_position="first")
    df = df.drop_duplicates(subset=["CLARTICOLO", "CLCOLORE", "PREZZOLPZ"], keep="last")

    out = pd.DataFrame({
        "CLARTICOLO": df["CLARTICOLO"].map(clean_text),
        "DESCRIZARTICOLOLI": df["DESCRIZARTICOLOLI"].map(clean_text),
        "CLCOLORE": df["CLCOLORE"].map(_format_codice),
        "CLDESCR": df["CLDESCR"].map(clean_text),
        "LIVELLOLPZ": pd.to_numeric(df["LIVELLOLPZ"], errors="coerce"),
        "PREZZOLPZ": pd.to_numeric(df["PREZZOLPZ"], errors="coerce"),
    })
    return out.reset_index(drop=True), []


def build_price_lookup(df: pd.DataFrame) -> dict[tuple[str, str], tuple]:
    lookup: dict[tuple[str, str], tuple] = {}
    if df is None or df.empty:
        return lookup
    for articolo, colore, livello_value, prezzo_value in df[
        ["CLARTICOLO", "CLCOLORE", "LIVELLOLPZ", "PREZZOLPZ"]
    ].itertuples(index=False, name=None):
        key = (articolo, colore)
        livello = None if pd.isna(livello_value) else livello_value
        prezzo = None if pd.isna(prezzo_value) else prezzo_value
        lookup[key] = (livello, prezzo)
    return lookup
