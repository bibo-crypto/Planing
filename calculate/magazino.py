"""Raw-yarn warehouse loading and aggregation rules."""
from __future__ import annotations

import pandas as pd

RAW_ARTICOLO_PREFIXES = ("G130", "G170")
KEEP_MAGAZZINI = {900910, 900160}


def _to_number(series):
    return pd.to_numeric(series.astype(str).str.strip().str.replace(",", ".", regex=False), errors="coerce")


def _header_key(value):
    return " ".join(str(value).replace("\ufeff", " ").strip().upper().split())


def _find_header_row(raw, required_cols, search_rows=20):
    required_set = {_header_key(value) for value in required_cols}
    for row_index in range(min(search_rows, len(raw))):
        row_values = {_header_key(value) for value in raw.iloc[row_index].tolist() if value is not None}
        if required_set.issubset(row_values):
            return row_index
    return None


def load_magazino(path, articolo_prefix=RAW_ARTICOLO_PREFIXES):
    required = ["MAGAZZINO", "ARTICOLO", "PARTITA", "ORDINE", "ESISTENZA", "COLLI"]
    summary_raw = pd.read_excel(path, header=None)
    summary_header_row = _find_header_row(summary_raw, ["Articolo", "Partita", "Mag.rocche", "Mag.peso"])
    if summary_header_row is not None:
        header = [_header_key(value) for value in summary_raw.iloc[summary_header_row].tolist()]
        summary = summary_raw.iloc[summary_header_row + 1:].copy()
        summary.columns = header
        aliases = {
            "ARTICOLO": "articolo", "TITOLO": "titolo", "DESCRIZIONE ART": "titolo",
            "PARTITA": "partita", "MAG.ROCCHE": "mag_rocche", "MAG.PESO": "mag_peso",
        }
        summary = summary.rename(columns={name: aliases.get(name, name) for name in summary.columns})
        required_summary = {"articolo", "partita", "mag_rocche", "mag_peso"}
        if required_summary.issubset(summary.columns):
            summary["articolo"] = summary["articolo"].astype(str).str.strip()
            summary["partita"] = summary["partita"].astype(str).str.strip()
            summary["mag_rocche"] = _to_number(summary["mag_rocche"])
            summary["mag_peso"] = _to_number(summary["mag_peso"])
            if "titolo" not in summary.columns:
                summary["titolo"] = ""
            summary["titolo"] = summary["titolo"].fillna("").astype(str).str.strip()
            mask = summary["partita"].ne("")
            if articolo_prefix is not None:
                prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
                mask &= summary["articolo"].str.startswith(prefixes)
            summary = summary[mask]
            return summary[["articolo", "titolo", "partita", "mag_rocche", "mag_peso"]].reset_index(drop=True), []

    header_row = _find_header_row(summary_raw, required)
    if header_row is None:
        return None, [f"Couldn't find a header row with: {', '.join(required)}"]
    df = summary_raw.iloc[header_row + 1:].copy()
    df.columns = [_header_key(value) for value in summary_raw.iloc[header_row].tolist()]
    required = [_header_key(value) for value in required]
    if len(df) and all(str(df.iloc[0].get(column, "")).strip().upper() == column for column in required):
        df = df.iloc[1:]
    df = df.reset_index(drop=True)
    missing = [column for column in required if column not in df.columns]
    if missing:
        return None, [f"Missing columns: {', '.join(missing)}"]

    df["MAGAZZINO"] = pd.to_numeric(df["MAGAZZINO"], errors="coerce")
    df["ORDINE"] = pd.to_numeric(df["ORDINE"], errors="coerce")
    df["ESISTENZA"] = _to_number(df["ESISTENZA"])
    df["COLLI"] = _to_number(df["COLLI"])
    df["ARTICOLO"] = df["ARTICOLO"].astype(str).str.strip()
    title_column = next((column for column in ("TITOLO", "DESCRIZIONE ART") if column in df.columns), None)
    df["TITOLO"] = "" if title_column is None else df[title_column]
    df["TITOLO"] = df["TITOLO"].fillna("").astype(str).str.strip()
    df["PARTITA"] = df["PARTITA"].astype(str).str.strip()
    df = df[df["MAGAZZINO"].isin(KEEP_MAGAZZINI)]
    if articolo_prefix is not None:
        prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
        df = df[df["ARTICOLO"].str.startswith(prefixes)]
    used_up = (df["MAGAZZINO"] == 900160) & (df["ORDINE"] == 0)
    df = df[~used_up]
    df = df[df["COLLI"] != 0]
    return pd.DataFrame({
        "articolo": df["ARTICOLO"], "titolo": df["TITOLO"], "partita": df["PARTITA"],
        "ordine": df["ORDINE"], "esistenza": df["ESISTENZA"], "colli": df["COLLI"],
        "magazzino": df["MAGAZZINO"],
    }), []


def summarize_by_partita(magazino_df):
    if magazino_df is None or magazino_df.empty:
        return pd.DataFrame(columns=["articolo", "titolo", "partita", "mag_rocche", "mag_peso"])
    if "titolo" not in magazino_df.columns:
        magazino_df = magazino_df.copy()
        magazino_df["titolo"] = ""
    grouped = magazino_df.groupby("partita", as_index=False).agg(
        articolo=("articolo", "first"), titulo=("titolo", "first"),
        mag_rocche=("colli", "sum"), mag_peso=("esistenza", "sum"),
    ).rename(columns={"titulo": "titolo"})
    return grouped[["articolo", "titolo", "partita", "mag_rocche", "mag_peso"]].sort_values(["articolo", "partita"])
