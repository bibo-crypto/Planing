"""LOTTI reference loading and Partita/Lotto grouping rules."""
from __future__ import annotations

import pandas as pd

RAW_ARTICOLO_PREFIXES = None


def _header_key(value):
    return " ".join(str(value).replace("\ufeff", " ").strip().upper().split())


def _find_header_row(raw, required_cols, search_rows=20):
    required_set = {_header_key(value) for value in required_cols}
    for r in range(min(search_rows, len(raw))):
        row_vals = {_header_key(v) for v in raw.iloc[r].tolist() if v is not None}
        if required_set.issubset(row_vals):
            return r
    return None


def load_lotti(path, articolo_prefix: str | tuple[str, ...] | None = RAW_ARTICOLO_PREFIXES):
    # LOTTI can have several sheets and is read from more than one tab --
    # cache each distinct articolo_prefix filter separately (same reasoning
    # as calculate.prezzi.load_prezzi's own caching).
    from utility import disk_cache
    namespace = f"lotti:{articolo_prefix!r}"
    return disk_cache.cached_load(
        namespace, path, lambda: _load_lotti_uncached(path, articolo_prefix),
        is_valid=lambda result: result[0] is not None,
    )


def _load_lotti_uncached(path, articolo_prefix: str | tuple[str, ...] | None = RAW_ARTICOLO_PREFIXES):
    required = ["MAGAZZINO", "ARTICOLO", "PARTITA", "ORDINE", "QESI", "LOTTO"]
    sheets = pd.read_excel(path, header=None, sheet_name=None)
    raw = None
    header_row = None
    for candidate in sheets.values():
        candidate_header = _find_header_row(candidate, required)
        if candidate_header is not None:
            raw = candidate
            header_row = candidate_header
            break
    if raw is None or header_row is None:
        return None, [f"Couldn't find a header row with: {', '.join(required)}"]

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [_header_key(value) for value in raw.iloc[header_row].tolist()]
    if len(df) and all(str(df.iloc[0][c]).strip().upper() == c for c in required):
        df = df.iloc[1:]
    df = df.reset_index(drop=True)
    missing = [c for c in required if c not in df.columns]
    if missing:
        return None, [f"Missing columns: {', '.join(missing)}"]

    df["MAGAZZINO"] = pd.to_numeric(df["MAGAZZINO"], errors="coerce")
    df["ORDINE"] = pd.to_numeric(df["ORDINE"], errors="coerce")
    df["QESI"] = pd.to_numeric(df["QESI"], errors="coerce")
    df["ARTICOLO"] = df["ARTICOLO"].astype(str).str.strip()
    df["PARTITA"] = df["PARTITA"].astype(str).str.strip()
    df["LOTTO"] = df["LOTTO"].astype(str).str.strip()
    keep = ((df["MAGAZZINO"] == 900910) & (df["ORDINE"] == 0)) | ((df["MAGAZZINO"] == 900160) & (df["ORDINE"] != 0))
    df = df[keep]
    df = df[df["QESI"] > 0]
    if articolo_prefix is not None:
        prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
        df = df[df["ARTICOLO"].str.startswith(prefixes)]

    out = pd.DataFrame({
        "articolo": df["ARTICOLO"], "partita": df["PARTITA"], "ordine": df["ORDINE"],
        "qesi": df["QESI"], "lotto": df["LOTTO"], "magazzino": df["MAGAZZINO"],
    })
    return out, []


def summarize_by_partita(lotti_df):
    if lotti_df is None or lotti_df.empty:
        return pd.DataFrame(columns=["partita", "lotto"])
    grouped = lotti_df.groupby("partita", as_index=False).agg(lotto=("lotto", "first"))
    return grouped[["partita", "lotto"]]


def load_lotti_snapshot(articolo_prefix=None):
    """Load normalized LOTTI data from SQLite without touching Excel."""
    from utility.source_manager import load as load_source
    df = load_source("lotti_data")
    if df is None or df.empty:
        return None, ["LOTTI is not stored in SQLite yet. Upload it once to update the database."]
    if articolo_prefix is not None and "articolo" in df.columns:
        prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
        df = df[df["articolo"].astype(str).str.startswith(prefixes)]
    return df.reset_index(drop=True), []
