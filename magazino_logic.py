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

RAW_ARTICOLO_PREFIX = "G130"
KEEP_MAGAZZINI = {900910, 900160}


def _to_number(series):
    """Parse the padded, comma-decimal numeric text these exports use (e.g. '  128,0  ')."""
    return pd.to_numeric(
        series.astype(str).str.strip().str.replace(",", ".", regex=False),
        errors="coerce",
    )


def _find_header_row(raw, required_cols, search_rows=6):
    required_set = set(required_cols)
    for r in range(min(search_rows, len(raw))):
        row_vals = set(str(v) for v in raw.iloc[r].tolist() if v is not None)
        if required_set.issubset(row_vals):
            return r
    return None


def load_magazino(path):
    """
    Reads the raw Magazino export and applies filter rules 1-4.
    Returns (DataFrame with columns [articolo, partita, ordine, esistenza,
    colli, magazzino], errors:list).
    """
    required = ["MAGAZZINO", "ARTICOLO", "PARTITA", "ORDINE", "ESISTENZA", "COLLI"]
    raw = pd.read_excel(path, header=None)
    header_row = _find_header_row(raw, required)
    if header_row is None:
        return None, [f"Couldn't find a header row with: {', '.join(required)}"]

    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row].tolist()
    if len(df) and all(str(df.iloc[0][c]).strip() == str(c).strip() for c in required):
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
    df = df[df["ARTICOLO"].str.startswith(RAW_ARTICOLO_PREFIX)]
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
