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
import os
from datetime import datetime

from utils import APP_DATA_DIR

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
    old_comment     TEXT,
    new_comment     TEXT,
    last_seen_at    TEXT,
    row_hash        TEXT
);

CREATE TABLE IF NOT EXISTS codes (
    articolo_filato TEXT PRIMARY KEY,
    titolo          TEXT
);

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
"""


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.executescript(SCHEMA)
    conn.commit()
    try:
        conn.execute("ALTER TABLE partita_state ADD COLUMN custom TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    try:
        conn.execute("ALTER TABLE upload_log ADD COLUMN file_path TEXT")
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
                days_in_qc, old_comment, new_comment, last_seen_at, row_hash)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(partita) DO UPDATE SET
                 cliente=excluded.cliente, articolo=excluded.articolo, titolo=excluded.titolo,
                 codice=excluded.codice, colore=excluded.colore, ordine=excluded.ordine,
                 riga=excluded.riga, data=excluded.data, consegna=excluded.consegna,
                 rocche=excluded.rocche, mc=excluded.mc, comment=excluded.comment,
                 cq=excluded.cq, bagno=excluded.bagno, tinto=excluded.tinto,
                 planedate=excluded.planedate, data_qualita=excluded.data_qualita,
                 data_uscita=excluded.data_uscita, custom=excluded.custom,
                 days_in_qc=excluded.days_in_qc,
                 old_comment=excluded.old_comment, new_comment=excluded.new_comment,
                 last_seen_at=excluded.last_seen_at, row_hash=excluded.row_hash
            """,
            (partita, row.get("cliente", ""), row.get("articolo", ""), row.get("titolo", ""),
             row.get("codice", ""), row.get("colore", ""), row.get("ordine", ""), row.get("riga", ""),
             row.get("data", ""), row.get("consegna", ""), row.get("rocche", ""), row.get("mc", ""),
             row.get("comment", ""), row.get("cq", ""), row.get("bagno", ""), row.get("tinto", ""),
             row.get("planedate", ""), row.get("data_qualita", ""), row.get("data_uscita", ""),
             row.get("custom", ""), row.get("days_in_qc", ""), old_comment, new_comment, now,
             row.get("row_hash", "")),
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
