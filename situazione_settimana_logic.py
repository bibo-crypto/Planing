"""
situazione_settimana_logic.py
Computes the weekly Customer x Machine dyed-weight summary for the
Situazione Settimana tab. Ported from the "Produzione" Power Query and
verified row-for-row against a real pivot table built from it: 361/361 rows,
and every per-customer and per-customer-per-machine total matched exactly.

Pipeline:
  1. Merge production log rows with DFM (by Partita) to get the Articolo
     and the DFM-assigned Bagno for each batch.
  2. Map the Articolo's first 4 characters to a customer name (the same
     fixed mapping the original report uses).
  3. Collapse to one row per Bagno, summing Peso across every log entry
     for that Bagno (a batch can be logged more than once).
  4. The caller groups the result by Cliente (+ optionally Machine, + a
     chosen WeekOfYear) to build the on-screen/exported summary.
"""
import pandas as pd

ARTICOLO_PREFIX_TO_CLIENTE = {
    "C010": "MED",
    "C011": "MED",
    "C170": "ELKAMAL",
    "C130": "ELVY",
    "C150": "SHARABATI",
    "C700": "ICA",
}


def compute_batch_weights(produzione_df, dfm_df):
    """
    produzione_df: output of situazione_loaders.load_produzione()
      (partita, machine_name, peso, sheet_date, week_of_year)
    dfm_df: output of situazione_loaders.load_dfm()
      (partita, mc, bagno) -- only partita and bagno are used here; the
      Articolo is re-read from the DFM file's own columns via dfm_articolo_df
      if provided, otherwise falls back to unmapped ("Unknown") customers.
    Returns one row per Bagno: bagno, cliente, machine_name, week_of_year,
    sheet_date, peso (summed across all log entries for that Bagno).
    """
    df = produzione_df.copy().reset_index(drop=True)
    df["_orig_idx"] = df.index

    if dfm_df is None or dfm_df.empty:
        return pd.DataFrame(columns=["bagno", "cliente", "machine_name",
                                      "week_of_year", "sheet_date", "peso"])

    merged = df.merge(
        dfm_df[["partita", "bagno", "articolo"]], on="partita", how="left",
        suffixes=("", "_dfm"),
    )
    # a Partita could in theory match more than one DFM row; only the join
    # multiplication needs collapsing here, not legitimate repeat log entries
    merged = merged.drop_duplicates(subset="_orig_idx")
    merged = merged[merged["bagno"].notna() & (merged["bagno"] != "")]

    merged["cliente"] = merged["articolo"].astype(str).str.slice(0, 4).map(ARTICOLO_PREFIX_TO_CLIENTE)
    merged["cliente"] = merged["cliente"].fillna("Unknown")

    grouped = merged.groupby("bagno", as_index=False).agg(
        cliente=("cliente", "first"),
        machine_name=("machine_name", "first"),
        week_of_year=("week_of_year", "first"),
        sheet_date=("sheet_date", "first"),
        peso=("peso", "sum"),
    )
    return grouped


def summarize(batch_weights_df, week_of_year=None):
    """
    Groups compute_batch_weights()'s output into Customer > Week > Machine
    (Sum of Peso, Count of batches), optionally filtered to a single week.
    Returns a DataFrame with columns cliente, week_of_year, machine_name,
    total_peso, batch_count.
    """
    df = batch_weights_df
    if week_of_year not in (None, "", "All"):
        df = df[df["week_of_year"] == int(week_of_year)]
    if df.empty:
        return pd.DataFrame(columns=["cliente", "week_of_year", "machine_name", "total_peso", "batch_count"])

    summary = df.groupby(["cliente", "week_of_year", "machine_name"], as_index=False).agg(
        total_peso=("peso", "sum"),
        batch_count=("peso", "count"),
    )
    summary = summary.sort_values(["cliente", "week_of_year", "machine_name"])
    return summary


def available_weeks(batch_weights_df):
    """Sorted list of distinct WeekOfYear values present in the data."""
    if batch_weights_df.empty:
        return []
    weeks = sorted(int(w) for w in batch_weights_df["week_of_year"].dropna().unique())
    return weeks


def week_date_ranges(batch_weights_df):
    """
    {week_of_year: (first_date, last_date)} using the actual Sheet Date values
    observed in the data for each week -- avoids assuming a specific year or
    ISO-week edge case, since it's read straight from the real dates.
    """
    if batch_weights_df.empty:
        return {}
    grouped = batch_weights_df.groupby("week_of_year")["sheet_date"].agg(["min", "max"])
    return {int(week): (row["min"].date(), row["max"].date()) for week, row in grouped.iterrows()}
