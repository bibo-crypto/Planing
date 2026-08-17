"""
prezzi_logic.py — Loads Delta's "Listini" (price list) export for the
Prezzi tab and applies the two filters it needs:

  1. DATAFINEVAL must be blank. A filled DATAFINEVAL means that price rate
     already has an end date -- it's no longer the current price, so it's
     excluded.
  2. CLCOLORE must NOT be blank. A blank colour code means the row has no
     colour attached to it at all, so it can't be matched to a specific
     price -- excluded.

Prices are looked up by Articolo (CLARTICOLO) + colour code (CLCOLORE)
together, since the same Articolo can have several colours, each priced
differently (different raw yarn per colour).
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from utils import clean_text

REQUIRED_COLUMNS = [
    "DATAFINEVAL", "DATAINIZIOVAL", "CLARTICOLO", "DESCRIZARTICOLOLI", "CLCOLORE",
    "CLDESCR", "LIVELLOLPZ", "PREZZOLPZ",
]

# The exact display order requested for the Prezzi Treeview/export.
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


def _format_codice(value) -> str:
    """CLCOLORE arrives as a float (1.0, 317.0, ...) -- show it as a plain code."""
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
    """
    Reads the Listini export (.xls or .xlsx) and returns
    (DataFrame[DISPLAY_COLUMNS], errors). errors is non-empty and the
    DataFrame is None when the file can't be read or is missing a
    required column.
    """
    try:
        df = pd.read_excel(path)
    except Exception as exc:  # noqa: BLE001
        return None, [f"Couldn't read the file: {exc}"]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return None, [f"Missing columns in the Listini file: {', '.join(missing)}"]

    # Rule 1: keep DATAFINEVAL blanks only.
    df = df[df["DATAFINEVAL"].isna()]
    # Rule 2: drop CLCOLORE blanks.
    df = df[df["CLCOLORE"].notna()]

    # Rule 3 (dedup): the same Articolo+Codice combo repeats across many
    # rows in this export (price-history revisions that all happen to have
    # a blank DATAFINEVAL). Collapse those down to the single most recent
    # one -- by DATAINIZIOVAL -- UNLESS the combo actually carries two (or
    # more) different prices in PREZZOLPZ, in which case each distinct
    # price is kept as its own row (never collapsed away) instead of
    # silently losing a real price difference.
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
    """
    (CLARTICOLO, CLCOLORE) -> (LIVELLOLPZ, PREZZOLPZ), for the Ordine Elvy
    extract's "Livello"/"Prezzo" columns to look up by Articolo (converted
    to its finished/C-prefixed form, see dfm_lookup.raw_to_finished_articolo)
    + Codice (COLOREDFM). load_prezzi() already dedups to at most one row
    per (Articolo, Codice, Prezzo), so a genuine duplicate here would only
    happen from two different prices -- last one wins, same as the table.
    """
    lookup: dict[tuple[str, str], tuple] = {}
    if df is None or df.empty:
        return lookup
    for _, r in df.iterrows():
        key = (r["CLARTICOLO"], r["CLCOLORE"])
        livello = None if pd.isna(r["LIVELLOLPZ"]) else r["LIVELLOLPZ"]
        prezzo = None if pd.isna(r["PREZZOLPZ"]) else r["PREZZOLPZ"]
        lookup[key] = (livello, prezzo)
    return lookup
