"""Weekly customer/machine dyed-weight summary business rules."""
from __future__ import annotations

import pandas as pd

ARTICOLO_PREFIX_TO_CLIENTE = {
    "C010": "MED", "C011": "MED", "C170": "ELKAMAL", "C130": "ELVY",
    "C150": "SHSRABATI", "C700": "ICA",
}


def compute_batch_weights(produzione_df, dfm_df):
    df = produzione_df.copy().reset_index(drop=True)
    df["_orig_idx"] = df.index
    if dfm_df is None or dfm_df.empty:
        return pd.DataFrame(columns=["bagno", "cliente", "machine_name", "week_of_year", "sheet_date", "peso"])
    merged = df.merge(dfm_df[["partita", "bagno", "articolo"]], on="partita", how="left", suffixes=("", "_dfm"))
    merged = merged.drop_duplicates(subset="_orig_idx")
    merged = merged[merged["bagno"].notna() & (merged["bagno"] != "")]
    merged["cliente"] = merged["articolo"].astype(str).str.slice(0, 4).map(ARTICOLO_PREFIX_TO_CLIENTE).fillna("Unknown")
    return merged.groupby("bagno", as_index=False).agg(
        cliente=("cliente", "first"), machine_name=("machine_name", "first"),
        week_of_year=("week_of_year", "first"), sheet_date=("sheet_date", "first"),
        peso=("peso", "sum"),
    )


def summarize(batch_weights_df, week_of_year=None):
    df = batch_weights_df
    if week_of_year not in (None, "", "All"):
        df = df[df["week_of_year"] == int(week_of_year)]
    if df.empty:
        return pd.DataFrame(columns=["cliente", "week_of_year", "machine_name", "total_peso", "batch_count"])
    summary = df.groupby(["cliente", "week_of_year", "machine_name"], as_index=False).agg(
        total_peso=("peso", "sum"), batch_count=("peso", "count"),
    )
    return summary.sort_values(["cliente", "week_of_year", "machine_name"])


def available_weeks(batch_weights_df):
    if batch_weights_df.empty:
        return []
    return sorted(int(week) for week in batch_weights_df["week_of_year"].dropna().unique())


def week_date_ranges(batch_weights_df):
    if batch_weights_df.empty:
        return {}
    grouped = batch_weights_df.groupby("week_of_year")["sheet_date"].agg(["min", "max"])
    return {int(week): (row["min"].date(), row["max"].date()) for week, row in grouped.iterrows()}
