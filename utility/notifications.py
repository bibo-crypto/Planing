"""Small persisted notification store used by the desktop UI.

This intentionally stays JSON-backed for now.  It can be moved to the future
central database without changing the pages that create or display notices.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from utility.utils import APP_DATA_DIR

_PATH = APP_DATA_DIR / "settings" / "notifications.json"


def _read() -> list[dict]:
    try:
        data = json.loads(_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, ValueError):
        return []


def _write(items: list[dict]) -> None:
    _PATH.parent.mkdir(parents=True, exist_ok=True)
    _PATH.write_text(json.dumps(items[-500:], ensure_ascii=False, indent=2), encoding="utf-8")


def list_open() -> list[dict]:
    return [item for item in _read() if not item.get("resolved")]


def add(key: str, title: str, message: str, page: str, severity: str = "medium", partita_colore: str = "") -> bool:
    items = _read()
    if any(item.get("key") == key and not item.get("resolved") for item in items):
        return False
    if not partita_colore or partita_colore.strip() in ("", "None", "nan"):
        import re
        partita_colore = ""
        # Try extracting Partita/Bagno or color code from message text
        m = re.search(r"(?i)(?:Partita|color code|code)\s+([A-Z0-9_\-\/]+)", message)
        if m:
            partita_colore = m.group(1).strip()
        else:
            parts = key.split(":")
            if len(parts) >= 3 and parts[0].startswith("situazione-price"):
                cand = parts[3] if len(parts) > 3 and not parts[3].replace(".", "").isdigit() else parts[2]
                if cand and cand not in ("None", "nan"):
                    partita_colore = cand
        if not partita_colore:
            partita_colore = "N/A"
    items.append({
        "key": key, "title": title, "message": message, "page": page,
        "partita_colore": partita_colore,
        "severity": severity, "created_at": datetime.now().isoformat(timespec="seconds"),
        "resolved": False,
    })
    _write(items)
    return True


def resolve(key: str) -> None:
    items = _read()
    changed = False
    for item in items:
        if item.get("key") == key and not item.get("resolved"):
            item["resolved"] = True
            changed = True
    if changed:
        _write(items)
