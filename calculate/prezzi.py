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
        # Not one of DISPLAY_COLUMNS, so it never shows in the Prezzi tab or
        # its export -- kept only so detect_price_anomalies() below can tell
        # which of a repeated Articolo+Colore's surviving distinct prices
        # came first.
        "_START_DATE": df["_start"].dt.strftime("%Y-%m-%d").fillna(""),
    })
    return out.reset_index(drop=True), []


def detect_price_anomalies(df: pd.DataFrame, min_pct_change: float = 10.0) -> pd.DataFrame:
    """Flag every Articolo+Colore that has more than one surviving distinct
    price (load_prezzi() already dedupes identical prices for the same
    combo, so any group with 2+ rows here is a genuine price change, not a
    duplicate export row) and the jump from the previous price to the next
    is at least ``min_pct_change`` percent in either direction.

    Returns one row per flagged transition, largest change first, so a
    handful of genuine repricings don't get buried under small rounding-size
    changes.
    """
    columns = ["CLARTICOLO", "CLCOLORE", "CLDESCR", "old_price", "new_price", "pct_change", "changed_on"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    for (articolo, colore), group in df.groupby(["CLARTICOLO", "CLCOLORE"]):
        if len(group) < 2:
            continue
        group = group.sort_values("_START_DATE")
        prices = group["PREZZOLPZ"].tolist()
        dates = group["_START_DATE"].tolist()
        descr = group["CLDESCR"].iloc[-1]
        for i in range(1, len(prices)):
            old_price, new_price = prices[i - 1], prices[i]
            if pd.isna(old_price) or pd.isna(new_price) or old_price == 0:
                continue
            pct_change = (new_price - old_price) / old_price * 100
            if abs(pct_change) >= min_pct_change:
                rows.append({
                    "CLARTICOLO": articolo, "CLCOLORE": colore, "CLDESCR": descr,
                    "old_price": old_price, "new_price": new_price,
                    "pct_change": round(pct_change, 1), "changed_on": dates[i] or "(no date)",
                })

    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(
        "pct_change", key=lambda s: s.abs(), ascending=False
    ).reset_index(drop=True)


def validate_price_data(df: pd.DataFrame) -> list[dict[str, str]]:
    """Return actionable data-quality notices for an uploaded Listini file."""
    if df is None or df.empty:
        return [{"key": "prezzi-empty", "title": "Prezzi file is empty", "message": "The active Listini file contains no usable color prices.", "severity": "high"}]
    issues = []
    invalid_price = df["PREZZOLPZ"].notna() & (pd.to_numeric(df["PREZZOLPZ"], errors="coerce") <= 0)
    if invalid_price.any():
        issues.append({"key": "prezzi-invalid-price", "title": "Invalid prices in Prezzi", "message": f"{int(invalid_price.sum())} row(s) have a zero or negative price.", "severity": "high"})
    missing_price = df["PREZZOLPZ"].isna()
    if missing_price.any():
        issues.append({"key": "prezzi-missing-price", "title": "Missing prices in Prezzi", "message": f"{int(missing_price.sum())} active row(s) have no price.", "severity": "high"})
    for (articolo, colore), group in df.groupby(["CLARTICOLO", "CLCOLORE"]):
        prices = group["PREZZOLPZ"].dropna().unique()
        levels = group["LIVELLOLPZ"].dropna().unique()
        if len(prices) > 1:
            issues.append({"key": f"prezzi-conflict-price:{articolo}:{colore}", "title": "Conflicting color prices", "message": f"Articolo {articolo}, colore {colore} has {len(prices)} different prices.", "severity": "high"})
        if len(levels) > 1:
            issues.append({"key": f"prezzi-conflict-level:{articolo}:{colore}", "title": "Conflicting color levels", "message": f"Articolo {articolo}, colore {colore} has {len(levels)} different LVL values.", "severity": "medium"})
    return issues


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
