"""SQLite-first source manager.

Excel files are import/update inputs only. Runtime consumers should ask this
module for normalized source data; the original workbook path is metadata and
is never required for normal startup.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import pandas as pd

from utility import situazione_db as db


# Snapshot name -> the shared "last uploaded file" cache that every page writes to
# when the operator picks that kind of file (module, loader function).
_PATH_CACHES = {
    "magazino_summary": ("utility.magazino_cache", "load_magazino_cache"),
    "lotti_summary": ("utility.lotti_cache", "load_lotti_cache"),
    "lotti_data": ("utility.lotti_cache", "load_lotti_cache"),
    "listini": ("utility.prezzi_cache", "load_prezzi_cache"),
    "densita_lookup": ("utility.densita_cache", "load_densita_cache"),
    "articoli_marca": ("utility.articoli_cache", "load_articoli_cache"),
}


def _registered_path(name: str) -> str:
    entry = _PATH_CACHES.get(name)
    if not entry:
        return ""
    try:
        import importlib
        module = importlib.import_module(entry[0])
        return str(getattr(module, entry[1])().get("source_path", "") or "")
    except Exception:
        return ""


def is_current(name: str) -> bool:
    """False when the file registered for *name* is newer than the snapshot built from it.

    Any page can register a new file (the shared path caches are written by every
    upload), but only some of them also refresh the SQLite snapshot. Comparing the
    snapshot's recorded source (path + modification time) with the registered file
    catches an upload made elsewhere AND a same-name file overwritten by today's
    export. If the file is gone, or nothing is registered, the snapshot is simply
    trusted -- the workbook is optional once its data is in SQLite.
    """
    registered = _registered_path(name)
    if not registered or not os.path.isfile(registered):
        return True
    try:
        meta = db.snapshot_meta(name)
    except Exception:
        return True
    if not meta or not meta.get("source_path"):
        return True        # no recorded origin to compare with
    same_path = os.path.normcase(os.path.normpath(str(meta["source_path"]))) == \
        os.path.normcase(os.path.normpath(registered))
    try:
        same_time = int(meta.get("source_mtime_ns") or 0) == os.stat(registered).st_mtime_ns
    except OSError:
        return True
    return same_path and same_time


def load(name: str) -> pd.DataFrame | None:
    """Return the latest normalized snapshot from SQLite, or None.

    None is also returned for a snapshot that is out of date with respect to the
    file registered for it (see is_current): callers already fall back to reading
    that file, which then refreshes the snapshot.
    """
    try:
        if not is_current(name):
            return None
        return db.load_frame_cache(name)
    except Exception:
        return None


def save(name: str, frame: pd.DataFrame, source_path: str | Path | None = None) -> bool:
    """Persist a normalized source snapshot and return whether it changed."""
    if frame is None:
        return False
    return bool(db.save_frame_cache(name, frame, source_path=source_path))


def metadata(name: str) -> dict[str, Any]:
    """Return upload metadata without requiring the original file to exist."""
    try:
        return db.get_upload(name) or {}
    except Exception:
        return {}


def source_name(name: str) -> str:
    return str(metadata(name).get("file_name", ""))
