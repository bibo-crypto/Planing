"""
magazino_logic.py
Raw yarn warehouse (Magazino) loading and aggregation, for the
"Magazino Filato" tab and for matching raw yarn to Ordini ELVY rows.

Rules (as specified):
  1. Keep only MAGAZZINO 900910 and 900160.
  2. Keep only ARTICOLO starting with the raw-yarn prefix (default "G130",
     the ELVY raw-yarn family -- change RAW_ARTICOLO_PREFIX below for a
     different client's family).
  3. MAGAZZINO 900910 with ORDINE == 0 is normal, kept as-is.
     MAGAZZINO 900160 with ORDINE == 0 means the batch is used up -> drop it.
  4. Drop any remaining row where COLLI == 0.
  5. Group what's left by PARTITA, summing COLLI (-> Mag.rocche) and
     ESISTENZA (-> Mag.peso); ARTICOLO is carried through (one Articolo per
     Partita in practice).

Verified against a real Magazino.xlsx export: rule 3 correctly dropped the
one 900160/ORDINE=0 row present, and rule 4 dropped the 3 COLLI=0 rows.
"""
import pandas as pd

RAW_ARTICOLO_PREFIXES = ("G130", "G170")   # Elvy, Kamal
KEEP_MAGAZZINI = {900910, 900160}


def _to_number(series):
    """Parse the padded, comma-decimal numeric text these exports use (e.g. '  128,0  ')."""
    return pd.to_numeric(
        series.astype(str).str.strip().str.replace(",", ".", regex=False),
        errors="coerce",
    )


def _header_key(value):
    return " ".join(str(value).replace("\ufeff", " ").strip().upper().split())


def _find_header_row(raw, required_cols, search_rows=20):
    required_set = {_header_key(value) for value in required_cols}
    for r in range(min(search_rows, len(raw))):
        row_vals = set(_header_key(v) for v in raw.iloc[r].tolist() if v is not None)
        if required_set.issubset(row_vals):
            return r
    return None


def load_magazino(path, articolo_prefix=RAW_ARTICOLO_PREFIXES):
    """
    Reads the raw Magazino export and applies filter rules 1-4.
    *articolo_prefix* selects which client's raw-yarn family/families to
    keep (rule 2) -- default keeps both Elvy (G130) and Kamal (G170) in
    one pass; pass a single string (e.g. "G170") to narrow it to one.
    Returns (DataFrame with columns [articolo, partita, ordine, esistenza,
    colli, magazzino], errors:list).
    """
    required = ["MAGAZZINO", "ARTICOLO", "PARTITA", "ORDINE", "ESISTENZA", "COLLI"]
    # The app can also consume its own exported summary (Articolo, Partita,
    # Mag.rocche, Mag.peso).  Accepting it avoids forcing users to keep the
    # original ERP export around for every conversion.
    summary_raw = pd.read_excel(path, header=None)
    summary_header_row = _find_header_row(
        summary_raw, ["Articolo", "Partita", "Mag.rocche", "Mag.peso"]
    )
    if summary_header_row is not None:
        header = [_header_key(v) for v in summary_raw.iloc[summary_header_row].tolist()]
        summary = summary_raw.iloc[summary_header_row + 1:].copy()
        summary.columns = header
        aliases = {
            "ARTICOLO": "articolo", "PARTITA": "partita",
            "MAG.ROCCHE": "mag_rocche", "MAG.PESO": "mag_peso",
        }
        columns = {name: aliases.get(name, name) for name in summary.columns}
        summary = summary.rename(columns=columns)
        required_summary = {"articolo", "partita", "mag_rocche", "mag_peso"}
        if required_summary.issubset(summary.columns):
            prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
            summary["articolo"] = summary["articolo"].astype(str).str.strip()
            summary["partita"] = summary["partita"].astype(str).str.strip()
            summary["mag_rocche"] = _to_number(summary["mag_rocche"])
            summary["mag_peso"] = _to_number(summary["mag_peso"])
            summary = summary[
                summary["articolo"].str.startswith(prefixes)
                & summary["partita"].ne("")
            ]
            return summary[["articolo", "partita", "mag_rocche", "mag_peso"]].reset_index(drop=True), []

    raw = summary_raw
    header_row = _find_header_row(raw, required)
    if header_row is None:
        return None, [f"Couldn't find a header row with: {', '.join(required)}"]

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [_header_key(value) for value in raw.iloc[header_row].tolist()]
    required = [_header_key(value) for value in required]
    if len(df) and all(str(df.iloc[0].get(c, "")).strip().upper() == c for c in required):
        df = df.iloc[1:]
    df = df.reset_index(drop=True)

    missing = [c for c in required if c not in df.columns]
    if missing:
        return None, [f"Missing columns: {', '.join(missing)}"]

    df["MAGAZZINO"] = pd.to_numeric(df["MAGAZZINO"], errors="coerce")
    df["ORDINE"] = pd.to_numeric(df["ORDINE"], errors="coerce")
    df["ESISTENZA"] = _to_number(df["ESISTENZA"])
    df["COLLI"] = _to_number(df["COLLI"])
    df["ARTICOLO"] = df["ARTICOLO"].astype(str).str.strip()
    df["PARTITA"] = df["PARTITA"].astype(str).str.strip()

    # rule 1: only these two warehouses
    df = df[df["MAGAZZINO"].isin(KEEP_MAGAZZINI)]
    # rule 2: only the raw-yarn article family
    prefixes = (articolo_prefix,) if isinstance(articolo_prefix, str) else tuple(articolo_prefix)
    df = df[df["ARTICOLO"].str.startswith(prefixes)]
    # rule 3: 900160 with ORDINE==0 means used up -> drop; 900910/ORDINE==0 is normal, keep
    used_up = (df["MAGAZZINO"] == 900160) & (df["ORDINE"] == 0)
    df = df[~used_up]
    # rule 4: COLLI == 0 -> drop
    df = df[df["COLLI"] != 0]

    out = pd.DataFrame({
        "articolo": df["ARTICOLO"],
        "partita": df["PARTITA"],
        "ordine": df["ORDINE"],
        "esistenza": df["ESISTENZA"],
        "colli": df["COLLI"],
        "magazzino": df["MAGAZZINO"],
    })
    return out, []


def summarize_by_partita(magazino_df):
    """
    Rule 5: group by Partita, summing Colli (-> mag_rocche) and Esistenza
    (-> mag_peso). Returns columns: articolo, partita, mag_rocche, mag_peso.
    """
    if magazino_df is None or magazino_df.empty:
        return pd.DataFrame(columns=["articolo", "partita", "mag_rocche", "mag_peso"])

    grouped = magazino_df.groupby("partita", as_index=False).agg(
        articolo=("articolo", "first"),
        mag_rocche=("colli", "sum"),
        mag_peso=("esistenza", "sum"),
    )
    return grouped[["articolo", "partita", "mag_rocche", "mag_peso"]].sort_values(["articolo", "partita"])
