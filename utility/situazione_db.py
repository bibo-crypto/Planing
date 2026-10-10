"""
db.py
SQLite persistence layer for the Dyeing Situation Tracker.

Tables:
- partita_state : one row per Partita (batch). Holds the current computed
  situation plus the rolling old_comment / new_comment history that used to
  live in the Old Situazione -> WOORKSHEET -> New Situazione Excel chain.
- codes         : Articoli reference table (Articolo Filato -> Titolo).
  Uploaded rarely; persists until the user re-uploads it.
- upload_log    : last validated upload per source, so the app can show
  green/red status per source tab and know what "no changes" means.
"""
import sqlite3
import json
import hashlib
import os
from io import StringIO
from datetime import datetime

from utility.utils import APP_DATA_DIR

DB_DIR = APP_DATA_DIR / "situazione"
DB_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = str(DB_DIR / "dyeing_tracker.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS partita_state (
    partita         TEXT PRIMARY KEY,
    cliente         TEXT,
    articolo        TEXT,
    titolo          TEXT,
    codice          TEXT,
    colore          TEXT,
    ordine          TEXT,
    riga            TEXT,
    data            TEXT,
    consegna        TEXT,
    rocche          TEXT,
    mc              TEXT,
    comment         TEXT,
    cq              TEXT,
    bagno           TEXT,
    tinto           TEXT,
    planedate       TEXT,
    data_qualita    TEXT,
    data_uscita     TEXT,
    custom          TEXT,
    days_in_qc      TEXT,
    ritardo_consegna TEXT,
    old_comment     TEXT,
    new_comment     TEXT,
    prezzo          TEXT,
    last_seen_at    TEXT,
    row_hash        TEXT
);

CREATE TABLE IF NOT EXISTS codes (
    articolo_filato TEXT PRIMARY KEY,
    titolo          TEXT
);

-- partita is already indexed (it's the primary key). The rest are
-- columns the app actually filters/searches/sorts by (Situazione
-- Generale's own search box, and the delivery/exit-date-driven Overview
-- cards) -- on a large history these were full table scans without an
-- index.
CREATE INDEX IF NOT EXISTS idx_partita_state_articolo ON partita_state(articolo);
CREATE INDEX IF NOT EXISTS idx_partita_state_bagno ON partita_state(bagno);
CREATE INDEX IF NOT EXISTS idx_partita_state_cliente ON partita_state(cliente);
CREATE INDEX IF NOT EXISTS idx_partita_state_data_uscita ON partita_state(data_uscita);
CREATE INDEX IF NOT EXISTS idx_partita_state_consegna ON partita_state(consegna);

CREATE TABLE IF NOT EXISTS upload_log (
    source_name     TEXT PRIMARY KEY,
    file_name       TEXT,
    file_path       TEXT,
    uploaded_at     TEXT,
    row_count       INTEGER,
    status          TEXT,
    message         TEXT,
    data_json       TEXT
);
CREATE TABLE IF NOT EXISTS frame_cache (
    source_name TEXT PRIMARY KEY,
    data_json TEXT NOT NULL,
    saved_at TEXT NOT NULL
);
-- Versioned row-level snapshots.  frame_cache remains as a backward-
-- compatible fallback for old installations, while new writes use these
-- tables so metadata and rows can be invalidated/queryed independently.
CREATE TABLE IF NOT EXISTS source_snapshot (
    source_name TEXT PRIMARY KEY,
    schema_version INTEGER NOT NULL DEFAULT 1,
    source_path TEXT,
    source_mtime_ns INTEGER,
    row_count INTEGER NOT NULL DEFAULT 0,
    saved_at TEXT NOT NULL,
    data_fingerprint TEXT,
    column_meta TEXT
);
CREATE TABLE IF NOT EXISTS source_snapshot_row (
    source_name TEXT NOT NULL,
    row_number INTEGER NOT NULL,
    row_json TEXT NOT NULL,
    PRIMARY KEY (source_name, row_number),
    FOREIGN KEY (source_name) REFERENCES source_snapshot(source_name) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_source_snapshot_row_source ON source_snapshot_row(source_name);
"""


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn = get_conn()
    conn.executescript(SCHEMA)
    conn.commit()
    try:
        conn.execute("ALTER TABLE source_snapshot ADD COLUMN data_fingerprint TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE source_snapshot ADD COLUMN column_meta TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE partita_state ADD COLUMN custom TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    try:
        conn.execute("ALTER TABLE partita_state ADD COLUMN ritardo_consegna TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    try:
        conn.execute("ALTER TABLE upload_log ADD COLUMN file_path TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    try:
        # Prezzo Ord. -- the price actually entered in Wincoint, read from
        # the source file (see situazione_loaders.load_wincoint_orders).
        # Existing installs get this column added on next launch; older
        # rows simply read back with an empty prezzo until their next Refresh.
        conn.execute("ALTER TABLE partita_state ADD COLUMN prezzo TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    conn.close()


def save_upload(source_name, file_name, row_count, status, message, data_json=None, file_path=None):
    conn = get_conn()
    conn.execute(
        """INSERT INTO upload_log (source_name, file_name, file_path, uploaded_at, row_count, status, message, data_json)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)
           ON CONFLICT(source_name) DO UPDATE SET
             file_name=excluded.file_name,
             file_path=excluded.file_path,
             uploaded_at=excluded.uploaded_at,
             row_count=excluded.row_count,
             status=excluded.status,
             message=excluded.message,
             data_json=excluded.data_json""",
        (source_name, file_name, file_path, datetime.now().isoformat(timespec="seconds"),
         row_count, status, message, data_json),
    )
    conn.commit()
    conn.close()


def get_upload(source_name):
    conn = get_conn()
    row = conn.execute("SELECT * FROM upload_log WHERE source_name=?", (source_name,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_uploads():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM upload_log").fetchall()
    conn.close()
    return {r["source_name"]: dict(r) for r in rows}


def frame_fingerprint(frame):
    """Stable content fingerprint for an uploaded/normalized DataFrame."""
    import pandas as pd
    if frame is None:
        return ""
    normalized = frame
    try:
        normalized = frame.copy()
        normalized.columns = [str(c) for c in normalized.columns]
        # Hash values + column names + dtypes.  This compares the actual
        # normalized data, not merely the source path or mtime.
        value_hash = pd.util.hash_pandas_object(normalized, index=True).values.tobytes()
        meta = json.dumps(
            {"columns": list(normalized.columns), "dtypes": [str(x) for x in normalized.dtypes]},
            ensure_ascii=False, separators=(",", ":"),
        ).encode("utf-8")
        return hashlib.sha256(meta + value_hash).hexdigest()
    except Exception:
        payload = normalized.to_json(orient="records", date_format="iso", default_handler=str)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def snapshot_fingerprint(source_name):
    conn = get_conn()
    row = conn.execute("SELECT data_fingerprint FROM source_snapshot WHERE source_name=?", (str(source_name),)).fetchone()
    conn.close()
    return row[0] if row and row[0] else ""


# Columns that hold dates in the loaders' output. Used only to restore snapshots
# written before dtypes were recorded (column_meta is NULL for those).
_LEGACY_DATETIME_COLUMNS = {"data", "consegna", "data_uscita", "data_qualita", "batch_start", "batchdt", "sheet_date"}


def _column_meta(frame):
    return json.dumps(
        {"columns": [str(c) for c in frame.columns], "dtypes": {str(c): str(t) for c, t in frame.dtypes.items()}},
        ensure_ascii=False, separators=(",", ":"),
    )


def save_frame_cache(source_name, frame, source_path=None):
    """Persist a normalized source frame in SQLite as the canonical local copy.

    Returns True when the stored data changed (new source or different
    content) and False when the frame is identical to what was already saved.
    In the identical case the rows are left alone (only the source path/time
    are refreshed), so re-uploading the same file costs almost nothing;
    otherwise the old rows are replaced as a whole inside one transaction.
    """
    if frame is None:
        return False
    source_name = str(source_name)
    fingerprint = frame_fingerprint(frame)
    previous = snapshot_fingerprint(source_name)
    changed = fingerprint != previous
    mtime_ns = None
    if source_path:
        try:
            mtime_ns = os.stat(source_path).st_mtime_ns
        except OSError:
            pass
    saved_at = datetime.now().isoformat(timespec="seconds")
    conn = get_conn()
    try:
        if not changed:
            conn.execute(
                "UPDATE source_snapshot SET source_path=?, source_mtime_ns=?, saved_at=? WHERE source_name=?",
                (str(source_path) if source_path else None, mtime_ns, saved_at, source_name),
            )
            conn.commit()
            return False
        conn.execute("DELETE FROM source_snapshot_row WHERE source_name=?", (source_name,))
        conn.execute(
            "INSERT INTO source_snapshot(source_name,schema_version,source_path,source_mtime_ns,row_count,saved_at,data_fingerprint,column_meta) "
            "VALUES(?,?,?,?,?,?,?,?) "
            "ON CONFLICT(source_name) DO UPDATE SET schema_version=excluded.schema_version, source_path=excluded.source_path, "
            "source_mtime_ns=excluded.source_mtime_ns, row_count=excluded.row_count, saved_at=excluded.saved_at, "
            "data_fingerprint=excluded.data_fingerprint, column_meta=excluded.column_meta",
            (source_name, 2, str(source_path) if source_path else None, mtime_ns, int(len(frame)), saved_at,
             fingerprint, _column_meta(frame)),
        )
        rows = [(source_name, i, json.dumps(row, ensure_ascii=False, default=str))
                for i, row in enumerate(frame.to_dict(orient="records"))]
        if rows:
            conn.executemany("INSERT INTO source_snapshot_row(source_name,row_number,row_json) VALUES(?,?,?)", rows)
        # source_snapshot_row is canonical; remove the legacy duplicate blob.
        conn.execute("DELETE FROM frame_cache WHERE source_name=?", (source_name,))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def snapshot_meta(source_name):
    """(source_path, source_mtime_ns) a snapshot was built from; {} if there is none."""
    conn = get_conn()
    try:
        row = conn.execute(
            "SELECT source_path, source_mtime_ns FROM source_snapshot WHERE source_name=?", (str(source_name),)
        ).fetchone()
    finally:
        conn.close()
    return {"source_path": row["source_path"], "source_mtime_ns": row["source_mtime_ns"]} if row else {}


def replace_snapshot(source_name, frame, source_path=None):
    """Alias used by upload reconciliation code; returns whether data changed."""
    return save_frame_cache(source_name, frame, source_path=source_path)


def _restore_frame(records, meta_json):
    """Rebuild a DataFrame from snapshot rows, restoring the original column order and date columns."""
    import pandas as pd
    meta = {}
    if meta_json:
        try:
            meta = json.loads(meta_json) or {}
        except (TypeError, ValueError):
            meta = {}
    columns = [c for c in meta.get("columns", [])]
    dtypes = meta.get("dtypes", {}) if isinstance(meta.get("dtypes", {}), dict) else {}
    frame = pd.DataFrame(records, columns=columns or None) if columns else pd.DataFrame(records)
    if columns:
        for column in columns:
            if column not in frame.columns:
                frame[column] = None
        frame = frame[columns]
    for column in frame.columns:
        dtype = dtypes.get(str(column), "")
        is_date = dtype.startswith("datetime64") if dtypes else str(column).casefold() in _LEGACY_DATETIME_COLUMNS
        if is_date:
            # Rows hold dates as text ("2026-09-01 00:00:00" or "NaT"); the
            # Situazione maths needs real datetimes (.dt accessors, subtraction).
            frame[column] = pd.to_datetime(frame[column], errors="coerce")
    return frame


def load_frame_cache(source_name):
    """Return a cached pandas DataFrame, or None when no snapshot exists."""
    import pandas as pd
    conn = get_conn()
    source_name = str(source_name)
    snapshot = conn.execute("SELECT schema_version, column_meta FROM source_snapshot WHERE source_name=?", (source_name,)).fetchone()
    if snapshot:
        rows = conn.execute("SELECT row_json FROM source_snapshot_row WHERE source_name=? ORDER BY row_number", (source_name,)).fetchall()
        conn.close()
        try:
            return _restore_frame([json.loads(row["row_json"]) for row in rows], snapshot["column_meta"])
        except (TypeError, ValueError, json.JSONDecodeError):
            return None
    row = conn.execute("SELECT data_json FROM frame_cache WHERE source_name=?", (source_name,)).fetchone()
    conn.close()
    if not row:
        return None
    try:
        return pd.read_json(StringIO(row["data_json"]), orient="records")
    except (TypeError, ValueError):
        return None


def save_codes(df):
    """df must have columns: articolo_filato, titolo"""
    conn = get_conn()
    conn.execute("DELETE FROM codes")
    conn.executemany(
        "INSERT OR REPLACE INTO codes (articolo_filato, titolo) VALUES (?, ?)",
        list(df[["articolo_filato", "titolo"]].itertuples(index=False, name=None)),
    )
    conn.commit()
    conn.close()


def load_codes():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM codes").fetchall()
    conn.close()
    return {r["articolo_filato"]: r["titolo"] for r in rows}


def get_all_states():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM partita_state").fetchall()
    conn.close()
    return {r["partita"]: dict(r) for r in rows}


def upsert_states(rows):
    """
    rows: list of dicts matching partita_state columns. Each row must already
    carry the `old_comment` and `new_comment` decided by
    situazione_logic.compute_situation() (which uses the previous
    new_comment, fetched via get_all_states(), as its `old_comment` input).
    This function just persists them and reports counts for the UI.
    Returns (added_count, updated_count, unchanged_count)
    """
    conn = get_conn()
    existing = {r["partita"]: dict(r) for r in conn.execute("SELECT * FROM partita_state").fetchall()}

    added, updated, unchanged = 0, 0, 0
    now = datetime.now().isoformat(timespec="seconds")
    for row in rows:
        partita = row["partita"]
        if not str(partita or "").strip():
            continue  # never persist a row with no Partita -- it can't be tracked meaningfully
        old_comment = row.get("old_comment", "")
        new_comment = row.get("new_comment", "")
        prev = existing.get(partita)

        if prev is None:
            added += 1
        elif prev["new_comment"] == new_comment:
            unchanged += 1
        else:
            updated += 1

        conn.execute(
            """INSERT INTO partita_state
               (partita, cliente, articolo, titolo, codice, colore, ordine, riga, data, consegna,
                rocche, mc, comment, cq, bagno, tinto, planedate, data_qualita, data_uscita, custom,
                days_in_qc, ritardo_consegna, old_comment, new_comment, prezzo, last_seen_at, row_hash)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(partita) DO UPDATE SET
                 cliente=excluded.cliente, articolo=excluded.articolo, titolo=excluded.titolo,
                 codice=excluded.codice, colore=excluded.colore, ordine=excluded.ordine,
                 riga=excluded.riga, data=excluded.data, consegna=excluded.consegna,
                 rocche=excluded.rocche, mc=excluded.mc, comment=excluded.comment,
                 cq=excluded.cq, bagno=excluded.bagno, tinto=excluded.tinto,
                 planedate=excluded.planedate, data_qualita=excluded.data_qualita,
                 data_uscita=excluded.data_uscita, custom=excluded.custom,
                 days_in_qc=excluded.days_in_qc,
                 ritardo_consegna=excluded.ritardo_consegna,
                 old_comment=excluded.old_comment, new_comment=excluded.new_comment,
                 prezzo=excluded.prezzo,
                 last_seen_at=excluded.last_seen_at, row_hash=excluded.row_hash
            """,
            (partita, row.get("cliente", ""), row.get("articolo", ""), row.get("titolo", ""),
             row.get("codice", ""), row.get("colore", ""), row.get("ordine", ""), row.get("riga", ""),
             row.get("data", ""), row.get("consegna", ""), row.get("rocche", ""), row.get("mc", ""),
             row.get("comment", ""), row.get("cq", ""), row.get("bagno", ""), row.get("tinto", ""),
             row.get("planedate", ""), row.get("data_qualita", ""), row.get("data_uscita", ""),
             row.get("custom", ""), row.get("days_in_qc", ""), row.get("ritardo_consegna", ""),
             old_comment, new_comment, row.get("prezzo", ""), now, row.get("row_hash", "")),
        )

    conn.commit()
    conn.close()
    return added, updated, unchanged


def remove_states_not_in(partite):
    """Remove saved batches that no longer exist in the current Wincoint file."""
    keep = {str(p).strip() for p in partite if str(p).strip()}
    conn = get_conn()
    if keep:
        placeholders = ",".join("?" for _ in keep)
        cursor = conn.execute(
            f"DELETE FROM partita_state WHERE partita NOT IN ({placeholders})",
            tuple(keep),
        )
    else:
        cursor = conn.execute("DELETE FROM partita_state")
    removed = cursor.rowcount
    conn.commit()
    conn.close()
    return removed
