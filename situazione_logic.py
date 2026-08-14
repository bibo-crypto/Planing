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
from datetime import datetime
import re

from utils import clean_text

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

    if dfm_df is not None and not dfm_df.empty:
        df = df.merge(dfm_df[["partita", "mc", "bagno"]], on="partita", how="left")
    else:
        df["mc"] = None
        df["bagno"] = None

    if data_prod_df is not None and not data_prod_df.empty:
        df = df.merge(data_prod_df[["partita", "end_prod"]], on="partita", how="left")
    else:
        df["end_prod"] = None

    df["end_prod_present"] = df["end_prod"].apply(lambda v: not _blank(v))
    df["tinto"] = pd.to_datetime(df["end_prod"], errors="coerce", dayfirst=True)

    if copertura_df is not None and not copertura_df.empty:
        df = df.merge(copertura_df[["bagno", "planedate"]], on="bagno", how="left")
    else:
        df["planedate"] = None

    if qualita_df is not None and not qualita_df.empty:
        df = df.merge(qualita_df[["partita", "data_qualita"]], on="partita", how="left")
    else:
        df["data_qualita"] = pd.NaT

    if uscita_df is not None and not uscita_df.empty:
        df = df.merge(uscita_df[["partita", "data_uscita"]], on="partita", how="left")
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
        return nc

    df["new_comment"] = df.apply(_row_new_comment, axis=1)

    # Custom flag column: "Check" when a batch has already gone through
    # quality (C.Q is something other than "OO"), more than 3 days have
    # passed since Data Qualita, and it still hasn't shipped (no Data Uscita).
    def _custom_flag(r):
        cq = r.get("cq")
        dq = r["data_qualita"] if pd.notna(r["data_qualita"]) else None
        du = r["data_uscita"] if pd.notna(r["data_uscita"]) else None
        if cq != "OO" and dq is not None and du is None:
            if (today - dq).days > 3:
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
         raw-yarn Articolo and checking whether one batch's available
         Mag.rocche covers the GROUP's total Rocche, not just this single
         row's, since several colours often draw from the same batch.

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

    remaining: dict[tuple[str, str], float] = {
        (str(r["articolo"]), str(r["partita"])): float(r["mag_rocche"])
        for _, r in magazino_summary.iterrows()
    }

    lotto_to_batch: dict[str, tuple[str, str]] = {}
    if not lotti_summary.empty and "lotto" in lotti_summary.columns:
        merged = magazino_summary.merge(lotti_summary, on="partita", how="inner")
        for _, r in merged.iterrows():
            lotto_key = clean_text(r["lotto"])
            if lotto_key:
                lotto_to_batch[lotto_key] = (str(r["articolo"]), str(r["partita"]))

    rocche = pd.to_numeric(df.get("rocche", 0), errors="coerce").fillna(0.0)
    handled: set = set()

    # 1) explicit Lotto in the comment -- deterministic, one Partita
    lot_rows: dict[str, list] = {}
    for idx in df.index[is_shortage]:
        m = _LOTTO_IN_COMMENT_RE.search(comment.loc[idx])
        if m:
            lot_rows.setdefault(clean_text(m.group(1)), []).append(idx)

    for lotto, idxs in lot_rows.items():
        batch_key = lotto_to_batch.get(lotto)
        if batch_key is None:
            continue
        needed = sum(rocche.loc[i] for i in idxs)
        if remaining.get(batch_key, 0.0) >= needed:
            remaining[batch_key] -= needed
            for i in idxs:
                result.loc[i] = f"{batch_key[0]} / {batch_key[1]}"
                handled.add(i)

    # 2) bare "PG-X" -- group by raw-yarn Articolo, cover the group's total
    if "articolo" in df.columns:
        groups: dict[str, list] = {}
        for idx in df.index[is_shortage]:
            if idx in handled:
                continue
            raw = _finished_articolo_to_raw(df.at[idx, "articolo"])
            if raw:
                groups.setdefault(raw, []).append(idx)

        for raw_articolo, idxs in groups.items():
            needed = sum(rocche.loc[i] for i in idxs)
            batches = [(k, v) for k, v in remaining.items() if k[0] == raw_articolo and v > 0]
            covering = next((b for b in sorted(batches, key=lambda kv: kv[1]) if b[1] >= needed), None)
            if covering is not None:
                remaining[covering[0]] -= needed
                for i in idxs:
                    result.loc[i] = f"{covering[0][0]} / {covering[0][1]}"
                continue
            # can't cover the whole group in one batch -> cover row by row
            # with the smallest batch that fits each one individually
            for i in idxs:
                need_i = rocche.loc[i]
                batches = [(k, v) for k, v in remaining.items() if k[0] == raw_articolo and v > 0]
                fit = next((b for b in sorted(batches, key=lambda kv: kv[1]) if b[1] >= need_i), None)
                if fit is None:
                    continue
                remaining[fit[0]] -= need_i
                result.loc[i] = f"{fit[0][0]} / {fit[0][1]}"

    return result
