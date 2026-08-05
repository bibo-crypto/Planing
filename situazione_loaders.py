"""
data_loaders.py
Reads each of the 6 raw source files, finds the real header row (these ERP
exports sometimes have a title row or a duplicated header row above the
data), validates that the expected columns exist, and returns a clean
pandas DataFrame plus a list of problems (empty list = OK).
"""
import pandas as pd


def _read_raw(path, sheet_name=0):
    return pd.read_excel(path, sheet_name=sheet_name, header=None)


def _find_header_row(raw, required_cols, search_rows=20):
    """Scan the first few rows for the one that contains all required_cols."""
    required_set = {_header_key(value) for value in required_cols}
    for r in range(min(search_rows, len(raw))):
        row_vals = {_header_key(v) for v in raw.iloc[r].tolist() if v is not None}
        if required_set.issubset(row_vals):
            return r
    return None


def _header_key(value):
    """Compare Excel headers while ignoring accidental spaces/case differences."""
    return " ".join(str(value).strip().split()).casefold()


def _s(series):
    """Convert a column to string safely: missing values become '' instead of the literal text 'nan'."""
    return series.fillna("").astype(str).str.strip()


def _load_with_header(path, required_cols, sheet_name=0):
    raw = _read_raw(path, sheet_name=sheet_name)
    header_row = _find_header_row(raw, required_cols)
    if header_row is None:
        return None, [f"Expected header row was not found. The file must contain: {', '.join(required_cols)}"]

    df = raw.iloc[header_row + 1:].copy()
    raw_headers = raw.iloc[header_row].tolist()
    required_by_key = {_header_key(value): value for value in required_cols}
    df.columns = [required_by_key.get(_header_key(value), value) for value in raw_headers]

    # some exports repeat the header row right after it -- drop it if so
    if len(df) and all(str(df.iloc[0][c]).strip() == str(c).strip() for c in required_cols):
        df = df.iloc[1:]

    df = df.reset_index(drop=True)
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        return None, [f"Missing required columns: {', '.join(missing)}"]

    # some exports repeat a header-like row further down (or with slightly
    # different blank columns) -- drop any row whose first required column
    # still literally equals that column's own name.
    first_col = required_cols[0]
    df = df[df[first_col].astype(str) != str(first_col)]
    df = df.reset_index(drop=True)
    return df, []


# ---------------------------------------------------------------------------
# Each loader below returns (DataFrame or None, errors:list)
# ---------------------------------------------------------------------------

def load_wincoint_orders(path):
    """WINCOINT sheet / Table3: raw customer orders.

    In the current Wincoint export the display fields are positional:
    column F is Codice and column G is Colore.  The remaining fields are
    resolved by their header names so blank/extra columns cannot shift them.
    """
    required = ["Cliente", "Articolo", "Data",
                "Consegna", "Partita", "Ordinata", "Descrizione aggiuntiva ordine", "Bagno"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    df = df[df["Partita"].notna()]
    df = df[df["Partita"].astype(str).str.strip() != ""]
    df = df[df["Partita"].astype(str) != "Partita"]
    df = df[df["Cliente"] != "TOTALE GENERALE"]
    if len(df.columns) < 7:
        return None, ["Wincoint must contain at least columns A through G (Codice and Colore)."]

    out = pd.DataFrame({
        "cliente": _s(df["Cliente"]),
        "articolo": _s(df["Articolo"]),
        "codice": _s(df.iloc[:, 5]),        # Wincoint column F -> Codice
        "colore": _s(df.iloc[:, 6]),        # Wincoint column G -> Colore
        "ordine": _s(df["Ordine"]) if "Ordine" in df.columns else "",  # no positional fallback: wrong index risks silently duplicating Codice
        "riga": _s(df["Riga"]),
        "data": pd.to_datetime(df["Data"], errors="coerce"),
        "consegna": pd.to_datetime(df["Consegna"], errors="coerce"),
        "partita": _s(df["Partita"]),
        "rocche": df["Ordinata"],
        "comment": _s(df["Descrizione aggiuntiva ordine"]),
        "cq": _s(df["Bagno"]),
    })
    return out, []


def load_dfm(path):
    """DFM sheet / Table8: machine + batch (Bagno) info."""
    required = ["DESCRIZIONECDMACCHINA", "PARTITADFM", "NUMEROPROGRAMMANEWDFM", "ARTICOLODFM"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    df = df[df["NUMEROPROGRAMMANEWDFM"] != "NUMEROPROGRAMMANEWDFM"]
    split = df["DESCRIZIONECDMACCHINA"].astype(str).str.split(" ", n=1, expand=True)
    out = pd.DataFrame({
        "partita": _s(df["PARTITADFM"]),
        "mc": pd.to_numeric(split[0], errors="coerce"),
        "bagno": _s(df["NUMEROPROGRAMMANEWDFM"]),
        "articolo": _s(df["ARTICOLODFM"]),
    })
    out = out.sort_values("partita")
    return out, []


def load_data_prod(path):
    """Data prod. sheet: when a batch finished production (End Prod.)."""
    required = ["Partia Col", "End", "Rilavorazione"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    df = df[(df["Partia Col"].notna()) & (df["Partia Col"] != "PROVA") & (df["Rilavorazione"] == 0)]
    end_split = df["End"].astype(str).str.split(" ", n=1, expand=True)
    out = pd.DataFrame({
        "partita": _s(df["Partia Col"]),
        "end_prod": _s(end_split[0]).str.replace(".", "/", regex=False),
    })
    out = out.drop_duplicates(subset="partita")
    return out, []


def load_schedulato(path):
    """Copertura sheet (Sheet1): machine batch queue -> PlaneDate per Bagno."""
    required = ["N. Bagno", "Machine", "Batch Start"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    df = df[df["N. Bagno"].notna()].copy()
    df["N. Bagno"] = _s(df["N. Bagno"])
    df["Machine"] = _s(df["Machine"])
    df["BatchDT"] = pd.to_datetime(df["Batch Start"], errors="coerce", dayfirst=True)
    df = df.sort_values(["Machine", "BatchDT"])
    df["RowID"] = df.groupby("Machine").cumcount()

    def plan_date_for(group):
        group = group.sort_values("RowID").reset_index(drop=True)
        start_date = group.loc[0, "BatchDT"]
        dates = []
        for _, r in group.iterrows():
            if r["RowID"] == 0:
                dates.append("IN MACC.")
                continue
            if pd.isna(start_date):
                dates.append(None)
                continue
            base_days = r["RowID"] // 2
            tentative = start_date + pd.Timedelta(days=base_days)
            # skip Fridays: add one day for every Friday crossed in the range
            date_range = pd.date_range(start_date, tentative, freq="D")
            fridays = sum(1 for d in date_range if d.dayofweek == 4)  # Monday=0 .. Friday=4
            final_date = tentative + pd.Timedelta(days=fridays)
            dates.append("PROD." + final_date.strftime("%d/%m/%Y"))
        group["PlaneDate"] = dates
        return group

    df = df.groupby("Machine", group_keys=False).apply(plan_date_for)
    out = pd.DataFrame({
        "bagno": df["N. Bagno"],
        "planedate": df["PlaneDate"],
    })
    return out, []


def load_uscita(path):
    """Uscita(J) sheet: warehouse exit date per Partita (shipped)."""
    required = ["Partita", "Bagno", "Data Mandato", "Contro Magazzino"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    df = df[df["Contro Magazzino"].isin([3004, 3009, "3004", "3009"])]
    out = pd.DataFrame({
        "partita": _s(df["Partita"]),
        "data_uscita": pd.to_datetime(df["Data Mandato"], errors="coerce"),
    })
    out = out.dropna(subset=["data_uscita"]).drop_duplicates(subset="partita")
    return out, []


def load_qualita(path):
    """Qualita sheet: quality-check date per Partita."""
    required = ["Partita", "Bagno", "Data Mandato"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    out = pd.DataFrame({
        "partita": _s(df["Partita"]),
        "data_qualita": pd.to_datetime(df["Data Mandato"], errors="coerce"),
    })
    out = out.sort_values(["partita", "data_qualita"]).drop_duplicates(subset="partita")
    return out, []


def load_codes(path):
    """Optional reference table: Articolo Filato -> Titolo. Uploaded rarely."""
    required = ["Articolo Filato", "TITOLO"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    out = pd.DataFrame({
        "articolo_filato": _s(df["Articolo Filato"]),
        "titolo": _s(df["TITOLO"]),
    })
    out = out.drop_duplicates(subset="articolo_filato")
    return out, []


def load_produzione(path):
    """
    Raw production log (Situazione Settimana): one row per process event,
    with the machine, batch (Partia Col), weight (Peso), yarn code, and the
    date it was logged. Current Data prod exports call that date ``End``;
    older weekly exports may call it ``Sheet Date``. Rework entries (Codice
    ending in "RI" or "T.C") are excluded, matching the original report's
    own rework flag.
    """
    # The file uploaded in Situazione is the same Data prod export used by
    # the weekly page. It has ``End`` rather than ``Sheet Date``. Keep the
    # common columns strict, then accept either date-column name.
    required = ["Machine Name", "Partia Col", "Peso", "Codice"]
    df, errors = _load_with_header(path, required)
    if df is None:
        return None, errors
    date_column = "Sheet Date" if "Sheet Date" in df.columns else "End"
    if date_column not in df.columns:
        return None, ["Expected a production date column named End or Sheet Date"]
    df = df[df["Partia Col"].notna()]
    df = df[df["Partia Col"].astype(str) != "Partia Col"]
    is_rework = _s(df["Codice"]).str.upper().str.endswith(("RI", "T.C"))
    df = df[~is_rework]

    out = pd.DataFrame({
        "partita": _s(df["Partia Col"]),
        "machine_name": _s(df["Machine Name"]),
        "peso": pd.to_numeric(df["Peso"], errors="coerce"),
        "sheet_date": pd.to_datetime(df[date_column], errors="coerce", dayfirst=True),
    })
    out = out.dropna(subset=["partita"])
    out["week_of_year"] = out["sheet_date"].dt.isocalendar().week
    return out, []


LOADERS = {
    "wincoint": ("Customer orders (WINCOINT)", load_wincoint_orders),
    "dfm": ("Machine data (DFM)", load_dfm),
    "data_prod": ("Production data", load_data_prod),
    "copertura": ("Schedule", load_schedulato),
    "uscita": ("Warehouse exit", load_uscita),
    "qualita": ("Quality check", load_qualita),
}
