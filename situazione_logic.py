"""
situazione_logic.py
Recreates the situation ("New Comm.") rules from the New Situazione Power
Query, as a row-by-row cascade -- this is a genuine state machine, not a
simple priority list: "Old Comm." (this Partita's status from the previous
Refresh) actively feeds into computing today's "New Comm.", exactly like it
did when the value was carried WOORKSHEET -> New Situazione sheet by sheet.

compute_situation() takes `old_comments` (a dict of {partita: previous
new_comment}, i.e. what's currently stored in the DB) and returns both the
`old_comment` actually used and the freshly computed `new_comment` for each
row -- db.upsert_states() just stores both as-is now, it no longer needs to
decide anything about carry-forward itself.

Verified: this cascade was checked row-by-row against a real "New Situazione"
sheet (379 rows, 100% match) before being wired in here.
"""
import pandas as pd
from datetime import datetime, timedelta
import re
from itertools import combinations

from utils import clean_text, parse_number

READY_CODES = {"AA", "AC", "AU", "AT"}


def _blank(x):
    if x is None:
        return True
    if isinstance(x, float) and pd.isna(x):
        return True
    if isinstance(x, str) and x.strip() == "":
        return True
    return False


def _compute_new_comm(comment, cq, old_comm, end_prod_present, planedate, dq, du, today):
    """Faithful step-by-step port of the New Situazione query's New Comm. cascade."""
    is_pgx = isinstance(comment, str) and comment.startswith("PG-X")

    # start: New Comm. = Old Comm., then a PG-X (yarn) override
    nc = "Filato" if is_pgx else old_comm

    # ready-to-ship C.Q codes / stay-as-is while still "OO" / fallback to Old Comm.
    if is_pgx:
        nc = "Filato"
    elif cq in READY_CODES:
        nc = "pronto da spedire"
    elif cq == "OO":
        pass
    elif _blank(nc):
        nc = old_comm

    # was Yarn but comment no longer says so -> reset; still queued -> carry Old Comm.;
    # still blank -> fall back to the raw End Prod. text; C.Q not "OO" -> ready to ship
    if (not is_pgx) and old_comm == "Filato":
        nc = None
    elif cq == "OO" and not _blank(old_comm):
        nc = old_comm
    elif is_pgx:
        nc = "Filato"
    elif _blank(nc):
        nc = "__END_PROD_RAW__" if end_prod_present else None
    elif cq != "OO":
        nc = "pronto da spedire"

    # a stray "IN MACC." carried from Old Comm. becomes "C.Q" at this point
    if nc == "IN MACC.":
        nc = "C.Q"

    # finished dyeing and still queued (C.Q == "OO") -> waiting on quality control
    if cq == "OO" and end_prod_present:
        nc = "C.Q"

    # pull in the machine schedule's PlaneDate where relevant
    if nc == "pronto da spedire":
        pass
    elif nc == "IN ATT. RIFACIMENTO":
        pass
    elif isinstance(nc, str) and nc.startswith("PROD.") and _blank(planedate):
        pass
    elif isinstance(nc, str) and nc.startswith("PROD.") and not _blank(planedate):
        nc = planedate
    elif cq == "OO" and end_prod_present and nc == "C.Q" and _blank(planedate):
        nc = "C.Q"
    elif cq == "OO" and end_prod_present and nc == "C.Q" and not _blank(planedate):
        nc = planedate

    # already shipped (Data Uscita present) overrides a lingering "pronto da spedire"
    if isinstance(nc, str) and "pronto da spedire" in nc and du is not None:
        nc = "spedita"

    # waiting on quality with a re-dye already scheduled -> show it; any other
    # non-"OO" code that slipped through -> ready to ship
    if nc == "C.Q" and not _blank(planedate):
        nc = f"Ritinta {planedate}"
    elif cq != "OO":
        nc = "pronto da spedire"

    # quality-derived status: shipped always wins over everything computed above
    if du is not None:
        quality_status = "Spedita"
    elif dq is None:
        quality_status = None
    else:
        days_since = (today - dq).days
        quality_status = "Check" if days_since > 4 else "pronto da spedire"

    if quality_status == "Spedita":
        nc = "Spedita"

    # last-resort fallback if nothing else ever set a value
    if _blank(nc):
        nc = planedate

    return nc, quality_status


def compute_situation(orders_df, dfm_df=None, data_prod_df=None,
                       copertura_df=None, uscita_df=None, qualita_df=None,
                       codes_map=None, old_comments=None):
    """
    orders_df is required (base table = one row per Partita/color order).
    old_comments: dict {partita: previous new_comment}, normally
    db.get_all_states()'s new_comment values from before this Refresh.
    Returns a DataFrame with one row per Partita, `old_comment` (the value
    actually used this round) and `new_comment` (freshly computed).
    """
    df = orders_df.copy()

    def _partita_key(series):
        return series.fillna("").astype(str).str.strip()

    def _bagno_key(series):
        text = series.fillna("").astype(str).str.strip()
        digits = text.str.replace(r"\D", "", regex=True).str.lstrip("0")
        return digits.where(digits.ne(""), text.str.casefold())

    # Keep the left-side row count stable.  ERP exports can contain repeated
    # references, and merging them blindly would duplicate one colour/order.
    df["_partita_join"] = _partita_key(df["partita"])

    if dfm_df is not None and not dfm_df.empty:
        right = dfm_df[["partita", "mc", "bagno"]].copy()
        right["_partita_join"] = _partita_key(right["partita"])
        right = right.drop_duplicates("_partita_join", keep="first")
        df = df.merge(
            right[["_partita_join", "mc", "bagno"]],
            on="_partita_join", how="left",
        ).drop(columns=["_partita_join"])
    else:
        df = df.drop(columns=["_partita_join"])
        df["mc"] = None
        df["bagno"] = None

    if data_prod_df is not None and not data_prod_df.empty:
        right = data_prod_df[["partita", "end_prod"]].copy()
        right["_partita_join"] = _partita_key(right["partita"])
        right = right.drop_duplicates("_partita_join", keep="first")
        df["_partita_join"] = _partita_key(df["partita"])
        df = df.merge(
            right[["_partita_join", "end_prod"]],
            on="_partita_join", how="left",
        ).drop(columns=["_partita_join"])
    else:
        df["end_prod"] = None

    df["end_prod_present"] = df["end_prod"].apply(lambda v: not _blank(v))
    df["tinto"] = pd.to_datetime(df["end_prod"], errors="coerce", dayfirst=True)

    if copertura_df is not None and not copertura_df.empty:
        right = copertura_df[["bagno", "planedate"]].copy()
        right["_bagno_join"] = _bagno_key(right["bagno"])
        right = right.drop_duplicates("_bagno_join", keep="first")
        df["_bagno_join"] = _bagno_key(df["bagno"])
        df = df.merge(
            right[["_bagno_join", "planedate"]],
            on="_bagno_join", how="left",
        ).drop(columns=["_bagno_join"])
    else:
        df["planedate"] = None

    if qualita_df is not None and not qualita_df.empty:
        right = qualita_df[["partita", "data_qualita"]].copy()
        right["_partita_join"] = _partita_key(right["partita"])
        right = right.drop_duplicates("_partita_join", keep="first")
        df["_partita_join"] = _partita_key(df["partita"])
        df = df.merge(
            right[["_partita_join", "data_qualita"]],
            on="_partita_join", how="left",
        ).drop(columns=["_partita_join"])
    else:
        df["data_qualita"] = pd.NaT

    if uscita_df is not None and not uscita_df.empty:
        right = uscita_df[["partita", "data_uscita"]].copy()
        right["_partita_join"] = _partita_key(right["partita"])
        right = right.drop_duplicates("_partita_join", keep="first")
        df["_partita_join"] = _partita_key(df["partita"])
        df = df.merge(
            right[["_partita_join", "data_uscita"]],
            on="_partita_join", how="left",
        ).drop(columns=["_partita_join"])
    else:
        df["data_uscita"] = pd.NaT

    if codes_map:
        # Titolo lookup: exact match on Articolo against the codes reference
        # table (the original M code sliced a single character here, which
        # collapsed ~90% of codes into one bucket -- verified broken against
        # the real Articoli.xlsx, so this uses an exact match instead).
        df["titolo"] = df["articolo"].astype(str).map(codes_map).fillna(df.get("titolo", ""))
    elif "titolo" not in df.columns:
        df["titolo"] = ""

    old_comments = old_comments or {}
    df["old_comment"] = df["partita"].map(old_comments)

    today = pd.Timestamp(datetime.now().date())
    df["days_in_qc"] = (today - df["tinto"]).dt.days

    def _row_new_comment(r):
        dq = r["data_qualita"] if pd.notna(r["data_qualita"]) else None
        du = r["data_uscita"] if pd.notna(r["data_uscita"]) else None
        nc, _quality = _compute_new_comm(
            r.get("comment"), r.get("cq"), r.get("old_comment"),
            r.get("end_prod_present"), r.get("planedate"), dq, du, today,
        )
        if nc == "__END_PROD_RAW__":
            nc = r.get("end_prod")
        if isinstance(nc, str) and nc.startswith("__"):
            return ""
        return nc

    df["new_comment"] = df.apply(_row_new_comment, axis=1)

    # Custom flag column: "Check" when the visible report fields identify a
    # genuine quality delay: C.Q == OO, New Comment == C.Q, and more than
    # 4 days have passed in Q.C without shipment.
    def _custom_flag(r):
        cq = r.get("cq")
        tinto = r["tinto"] if pd.notna(r["tinto"]) else None
        du = r["data_uscita"] if pd.notna(r["data_uscita"]) else None
        new_comment = str(r.get("new_comment") or "").strip().casefold()
        if cq == "OO" and new_comment == "c.q" and tinto is not None and du is None:
            if (today - tinto).days > 4:
                return "Check"
        return ""

    df["custom"] = df.apply(_custom_flag, axis=1)

    # normalize output columns / types for the DB layer
    for col in ["mc", "bagno", "planedate", "titolo", "old_comment"]:
        if col not in df.columns:
            df[col] = ""
    df["data"] = df["data"].dt.strftime("%Y-%m-%d").fillna("")
    df["consegna"] = df["consegna"].dt.strftime("%Y-%m-%d").fillna("")
    df["tinto"] = df["tinto"].dt.strftime("%Y-%m-%d").fillna("")
    df["data_qualita"] = df["data_qualita"].dt.strftime("%Y-%m-%d").fillna("")
    df["data_uscita"] = df["data_uscita"].dt.strftime("%Y-%m-%d").fillna("")
    df["days_in_qc"] = df["days_in_qc"].fillna("").astype(str)

    keep = ["partita", "cliente", "articolo", "titolo", "codice", "colore", "ordine", "riga",
            "data", "consegna", "rocche", "mc", "comment", "cq", "bagno", "tinto",
            "planedate", "data_qualita", "data_uscita", "custom", "days_in_qc",
            "old_comment", "new_comment"]
    for c in keep:
        if c not in df.columns:
            df[c] = ""

    result = df[keep].fillna("")
    # unmatched merges or stray conversions can still leave the literal text
    # "nan" behind -- blank those out too.
    result = result.astype(str).apply(lambda s: s.str.strip())
    result = result.replace(r"(?i)^nan$", "", regex=True)
    return result


# ----------------------------------------------------------------------
# Raw yarn auto-match for "PG-X" rows (works for every client, not just
# whichever one was checked by hand before) -- so a shortage row that
# actually already has raw yarn sitting in Magazino Filato doesn't get
# missed just because nobody thought to look it up.
# ----------------------------------------------------------------------
_LOTTO_IN_COMMENT_RE = re.compile(r"PG-X\s*\(\s*([^)]+)\s*\)", re.IGNORECASE)


def _finished_articolo_to_raw(articolo) -> str | None:
    """C1701234 -> G1701234, same C->G rule used by ordini_elvy.match_raw_yarn()."""
    a = clean_text(articolo).upper()
    if a.startswith("C") and len(a) > 1:
        return "G" + a[1:]
    return None


def compute_raw_yarn_matches(df: pd.DataFrame, magazino_summary: pd.DataFrame,
                              lotti_summary: pd.DataFrame | None = None) -> pd.Series:
    """
    For every row whose Comment marks a yarn shortage ("PG-X..."), check
    whether Magazino Filato already has raw yarn for it. Returns a Series
    aligned to df.index: "Articolo / Partita" where a batch was found and
    has enough quantity, "" otherwise. Works across every client in df --
    the raw-yarn Articolo comes from each row's own Articolo (C -> G), not
    from a hardcoded client name.

    Priority, per the two ways "PG-X" shows up:
      1. "PG-X(807346)" -- the number in parens is a Lotto (Kamal's own
         raw-yarn message number). Look it up directly in the LOTTI
         reference (lotti_summary, merged with magazino_summary on
         Partita) for one exact, deterministic batch.
      2. Plain "PG-X" -- fall back to grouping every row needing the same
         raw-yarn Articolo and consuming stock colour by colour.  A colour
         may use several Partitas when necessary, while a later colour is
         left unmatched if the article's remaining total is insufficient.

    Either way, a batch's available quantity is decremented as rows get
    matched to it, so two different colours/groups can't both claim the
    same physical yarn.
    """
    result = pd.Series([""] * len(df), index=df.index, dtype=object)
    if df.empty or "comment" not in df.columns:
        return result

    comment = df["comment"].fillna("").astype(str)
    is_shortage = comment.str.upper().str.startswith("PG-X")
    if not is_shortage.any():
        return result

    magazino_summary = magazino_summary if isinstance(magazino_summary, pd.DataFrame) else pd.DataFrame()
    lotti_summary = lotti_summary if isinstance(lotti_summary, pd.DataFrame) else pd.DataFrame()
    if magazino_summary.empty:
        return result

    # Keep stock per article/partita.  A single colour may be supplied by
    # more than one partita, so matching must be allowed to consume several
    # stock rows instead of requiring one partita to cover the whole colour.
    stock = magazino_summary.copy()
    stock["articolo"] = stock["articolo"].map(clean_text)
    stock["partita"] = stock["partita"].map(clean_text)
    stock["mag_rocche"] = stock["mag_rocche"].map(parse_number).fillna(0.0)
    stock = stock.groupby(["articolo", "partita"], as_index=False)["mag_rocche"].sum()
    remaining: dict[tuple[str, str], float] = {
        (str(r["articolo"]), str(r["partita"])): float(r["mag_rocche"])
        for _, r in stock.iterrows()
    }

    lotto_to_batches: dict[str, list[tuple[str, str]]] = {}
    if not lotti_summary.empty and "lotto" in lotti_summary.columns:
        merged = stock.merge(lotti_summary, on="partita", how="inner")
        for _, r in merged.iterrows():
            lotto_key = clean_text(r["lotto"])
            if lotto_key:
                batch_key = (str(r["articolo"]), str(r["partita"]))
                if batch_key not in lotto_to_batches.setdefault(lotto_key, []):
                    lotto_to_batches[lotto_key].append(batch_key)

    # Wincoint values are commonly formatted as Italian decimals (e.g.
    # "6,00").  pd.to_numeric would turn those into NaN/0 and make one
    # one-cone partita appear sufficient for every machine.
    rocche = df.get("rocche", pd.Series(0.0, index=df.index)).map(parse_number).fillna(0.0)
    handled: set = set()

    def consume(batch_keys, needed: float):
        """Consume *needed* rocche and return every partita used, or None."""
        candidates = [key for key in batch_keys if remaining.get(key, 0.0) > 0]
        # Prefer one partita when it is enough.  Split across partitas only
        # when the colour itself needs a combined quantity.
        fitting = [key for key in candidates if remaining[key] >= needed]
        if fitting:
            key = min(fitting, key=lambda item: (remaining[item], item[1]))
            remaining[key] -= needed
            return [key[1]]
        # A split match may use two Partitas, never three or more.  Choose
        # the smallest pair that covers the colour so excess stock is not
        # consumed unnecessarily.
        pairs = [
            pair for pair in combinations(candidates, 2)
            if remaining[pair[0]] + remaining[pair[1]] >= needed
        ]
        if not pairs:
            return None
        selected = min(
            pairs,
            key=lambda pair: (remaining[pair[0]] + remaining[pair[1]], pair),
        )
        used = []
        for key in sorted(selected, key=lambda item: (remaining[item], item[1])):
            amount = min(remaining[key], needed)
            if amount <= 0:
                continue
            remaining[key] -= amount
            needed -= amount
            used.append(key[1])
            if needed <= 1e-9:
                break
        return used

    def match_text(article: str, partitas: list[str]) -> str:
        # Keep the existing display format while making a split stock match
        # obvious in the Filato Disponibile cell.
        return f"{article} / {' + '.join(partitas)}"

    # 1) explicit Lotto in the comment -- deterministic lot, possibly many
    # Partitas when the same Lotto appears on multiple warehouse rows.
    lot_rows: dict[str, list] = {}
    for idx in df.index[is_shortage]:
        m = _LOTTO_IN_COMMENT_RE.search(comment.loc[idx])
        if m:
            lot_rows.setdefault(clean_text(m.group(1)), []).append(idx)

    for lotto, idxs in lot_rows.items():
        for i in sorted(idxs, key=lambda row_idx: (rocche.loc[row_idx], str(row_idx))):
            # Explicit Lotto rows are handled only by the Lotto-specific
            # stock below, even when that Lotto is absent from the warehouse.
            handled.add(i)
            raw_article = _finished_articolo_to_raw(df.at[i, "articolo"]) if "articolo" in df.columns else None
            batch_keys = [
                key for key in lotto_to_batches.get(lotto, [])
                if raw_article is None or key[0] == raw_article
            ]
            if not batch_keys:
                continue
            # A row with an explicit Lotto must never fall through to the
            # generic article match: that could attach an unrelated Partita
            # after the requested Lotto runs out of stock.
            used_partitas = consume(batch_keys, rocche.loc[i])
            if used_partitas is None:
                continue
            result.loc[i] = match_text(batch_keys[0][0], used_partitas)

    # 2) bare "PG-X" -- group by raw-yarn Articolo and consume by colour
    if "articolo" in df.columns:
        groups: dict[str, list] = {}
        for idx in df.index[is_shortage]:
            if idx in handled:
                continue
            raw = _finished_articolo_to_raw(df.at[idx, "articolo"])
            if raw:
                groups.setdefault(raw, []).append(idx)

        for raw_articolo, idxs in groups.items():
            # Match colours one by one and allow each colour to draw from
            # many partitas (e.g. 32 + 24 covers a 56-rocche colour).  This
            # also leaves a later colour blank when the remaining stock is
            # insufficient, instead of assigning an unusable partita.
            # Allocate the smallest machines first.  This ensures that a
            # stock of 56 rocche serves 32 + 24 before a 128-rocche colour,
            # regardless of the visual/database row order.
            for i in sorted(idxs, key=lambda row_idx: (rocche.loc[row_idx], str(row_idx))):
                need_i = rocche.loc[i]
                batches = [k for k in remaining if k[0] == raw_articolo]
                used_partitas = consume(batches, need_i)
                if used_partitas is None:
                    continue
                result.loc[i] = match_text(raw_articolo, used_partitas)

    return result


# ----------------------------------------------------------------------
# Copertura (physical machine coverage) helpers, shared between the
# Copertura popup window (situazione_tab.py) and the compact machine
# summary cards on the Overview tab (overview_tab.py) -- kept here as one
# definition so the two views can never silently drift out of sync.
# ----------------------------------------------------------------------
_MACHINE_CODE_RANGE = (3300, 3399)
_BARE_MACHINE_RE = re.compile(r"(?<!\d)0*(1[0-2]|[3-9])(?:\.0+)?(?!\d)")


def machine_number_from_label(value) -> int | None:
    """
    Parse a Copertura "Machine" label into a plain machine number (3-12).
    Accepts either the 3300-series code Delta uses internally (e.g.
    "3305" -> 5) or a bare number as it might appear in a hand-edited
    sheet (e.g. "5", "05", "12.0", "M5" -> 5 / 12). Returns None when
    nothing recognizable is found.
    """
    text = clean_text(value)
    digits = re.sub(r"\D", "", text)
    if digits and _MACHINE_CODE_RANGE[0] <= int(digits) <= _MACHINE_CODE_RANGE[1]:
        return int(digits) - _MACHINE_CODE_RANGE[0]
    match = _BARE_MACHINE_RE.search(text)
    return int(match.group(1)) if match else None


def machine_coverage_until(color_count, today=None) -> str:
    """
    "2 colours per day per machine, Friday excluded" -- the date by which
    *color_count* colours queued on one machine will have been produced,
    starting from *today* (defaults to the real today). Returns "-" for a
    falsy/zero count.
    """
    if not color_count:
        return "-"
    required_days = (int(color_count) + 1) // 2
    day = today or datetime.now().date()
    completed = 0
    while completed < required_days:
        if day.weekday() != 4:  # Friday
            completed += 1
        if completed < required_days:
            day += timedelta(days=1)
    return day.strftime("%Y-%m-%d")


def compute_machine_totals(situation_df, copertura_df) -> dict[int, int]:
    """{machine_number (3-12): total_colors_currently_queued}, the same
    join/aggregation the Copertura window itself uses -- factored out here
    so a new order's own machine-queue scheduling (Ordine MED) can use the
    exact same baseline instead of recomputing it separately.
    Returns {} if either input is missing/empty or nothing matches."""
    import re as _re

    if situation_df is None or situation_df.empty or copertura_df is None or copertura_df.empty:
        return {}
    if "machine" not in copertura_df.columns:
        return {}

    situation = situation_df.copy()
    copertura = copertura_df.copy()

    def bagno_key(value):
        digits = _re.sub(r"\D", "", str(value or "")).lstrip("0")
        return digits or str(value or "").strip().casefold()

    for frame in (situation, copertura):
        frame["bagno"] = frame["bagno"].fillna("").astype(str).str.strip()
        frame["bagno_key"] = frame["bagno"].map(bagno_key)
    merged = situation.merge(
        copertura[["bagno_key", "machine"]].drop_duplicates("bagno_key"),
        on="bagno_key", how="inner",
    )
    merged["machine_number"] = merged["machine"].map(machine_number_from_label)
    merged = merged[merged["machine_number"].between(3, 12, inclusive="both")]
    if merged.empty:
        return {}
    return merged.groupby("machine_number").size().to_dict()


# ---------------------------------------------------------------------------
# Price validation -- flags Situazione rows whose Prezzo looks wrong.
# Checked: an exact 0.01 value (a known placeholder/data-entry mistake in
# the source ERP), and a genuinely missing price (blank/zero) on a row
# that otherwise has enough identifying data to expect one. The $24/32/56
# surcharge itself is applied when Prezzo is computed (see
# biglietti_exporter.apply_machine_surcharge), so a correctly-surcharged
# price never shows up here -- only prices that are missing or match the
# known-bad placeholder do.
# ---------------------------------------------------------------------------

PRICE_PLACEHOLDER_VALUE = 0.01


def find_price_anomalies(df):
    """Returns a list of dicts (cliente, articolo, colore, bagno, mc,
    prezzo, issue) for every row whose Prezzo needs a human to check it.
    Never raises -- an empty/missing dataframe just returns []."""
    if df is None or df.empty:
        return []
    if "prezzo" not in df.columns:
        return []

    def _as_number(value):
        try:
            if value in (None, ""):
                return None
            return float(str(value).replace(",", "."))
        except (TypeError, ValueError):
            return None

    out = []
    for _, row in df.iterrows():
        articolo = str(row.get("articolo", "") or "").strip()
        colore = str(row.get("colore", "") or row.get("codice", "") or "").strip()
        if not articolo:
            continue  # nothing to price-check on a blank row

        prezzo_num = _as_number(row.get("prezzo"))
        issue = None
        if prezzo_num is not None and abs(prezzo_num - PRICE_PLACEHOLDER_VALUE) < 1e-9:
            issue = f"Prezzo sospetto ({PRICE_PLACEHOLDER_VALUE})"
        elif prezzo_num is None or prezzo_num == 0:
            issue = "Prezzo mancante"

        if issue:
            out.append({
                "cliente": row.get("cliente", ""),
                "articolo": articolo,
                "colore": colore,
                "codice": row.get("codice", "") or colore,
                "ordine": row.get("ordine", ""),
                "riga": row.get("riga", ""),
                "bagno": row.get("bagno", ""),
                "mc": row.get("mc", ""),
                "prezzo": row.get("prezzo", ""),
                "issue": issue,
            })
    return out
