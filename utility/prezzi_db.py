"""SQLite-backed cache for the normalized Listini dataframe."""
from __future__ import annotations

import json
import sqlite3
from io import StringIO
from pathlib import Path

import pandas as pd

from utility.orders_db import DB_PATH


def _connect(db_path: Path | None = None) -> sqlite3.Connection:
    path = Path(db_path or DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.execute("""CREATE TABLE IF NOT EXISTS prezzi_frame_cache (
        cache_key TEXT PRIMARY KEY,
        fingerprint TEXT NOT NULL,
        payload TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )""")
    return connection


def load_prezzi_frame(
    fingerprint: tuple[object, ...], db_path: Path | None = None,
) -> pd.DataFrame | None:
    fingerprint_json = json.dumps(fingerprint, ensure_ascii=False, separators=(",", ":"))
    connection = None
    try:
        connection = _connect(db_path)
        row = connection.execute(
            "SELECT fingerprint, payload FROM prezzi_frame_cache WHERE cache_key='listini'"
        ).fetchone()
        if row is None or row[0] != fingerprint_json:
            return None
        return pd.read_json(StringIO(row[1]), orient="records", convert_dates=False)
    except (OSError, sqlite3.Error, TypeError, ValueError):
        return None
    finally:
        if connection is not None:
            connection.close()


def save_prezzi_frame(
    fingerprint: tuple[object, ...], frame: pd.DataFrame, db_path: Path | None = None,
) -> None:
    fingerprint_json = json.dumps(fingerprint, ensure_ascii=False, separators=(",", ":"))
    payload = frame.to_json(orient="records", date_format="iso", double_precision=15)
    connection = _connect(db_path)
    try:
        connection.execute(
            """INSERT INTO prezzi_frame_cache(cache_key, fingerprint, payload, updated_at)
               VALUES('listini', ?, ?, CURRENT_TIMESTAMP)
               ON CONFLICT(cache_key) DO UPDATE SET
                   fingerprint=excluded.fingerprint,
                   payload=excluded.payload,
                   updated_at=excluded.updated_at""",
            (fingerprint_json, payload),
        )
        connection.commit()
    finally:
        connection.close()