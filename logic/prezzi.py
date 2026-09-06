"""Listini loading and price lookup business rules."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from utils import clean_text

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
