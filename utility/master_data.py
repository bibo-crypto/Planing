"""Local reference data store.

This is deliberately JSON-backed until the central database phase is approved.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from utility.utils import APP_DATA_DIR

PATH = APP_DATA_DIR / "settings" / "master_data.json"
DEFAULTS = {
    "customers": [],
    "machines": [
        {"code": "3301", "number": "11", "rocche": "6"},
        {"code": "3310", "number": "12", "rocche": "24"},
        {"code": "3306", "number": "9", "rocche": "32"},
        {"code": "3302", "number": "10", "rocche": "56"},
        {"code": "3307", "number": "7", "rocche": "72"},
        {"code": "3303", "number": "8", "rocche": "128"},
        {"code": "3308", "number": "5", "rocche": "192"},
        {"code": "3304", "number": "6", "rocche": "384"},
        {"code": "3309", "number": "3", "rocche": "672"},
    ],
}


def load() -> dict[str, list[dict[str, Any]]]:
    try:
        data = json.loads(PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        data = {}
    result = {key: list(data.get(key, defaults)) for key, defaults in DEFAULTS.items()}
    # Drop the old, generic machine rows from the first prototype.
    result["machines"] = [m for m in result["machines"] if m.get("number") and m.get("rocche")]
    return result


def save(data: dict[str, list[dict[str, Any]]]) -> None:
    PATH.parent.mkdir(parents=True, exist_ok=True)
    PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
