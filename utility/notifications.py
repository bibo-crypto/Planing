"""Persisted, categorized notification store used by the desktop UI."""
from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

from utility.utils import APP_DATA_DIR

_PATH = APP_DATA_DIR / "settings" / "notifications.json"
CATEGORIES = ("Urgent", "Raw Yarn", "Prices", "Quality", "Delay", "Files", "System", "Other")


def _read() -> list[dict]:
    try:
        data = json.loads(_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, ValueError):
        return []


def _write(items: list[dict]) -> None:
    _PATH.parent.mkdir(parents=True, exist_ok=True)
    _PATH.write_text(json.dumps(items[-1000:], ensure_ascii=False, indent=2), encoding="utf-8")


def category_for(key: str, title: str = "", page: str = "") -> str:
    text = f"{key} {title} {page}".casefold()
    if any(x in text for x in ("price", "prezzi", "prezzo")):
        return "Prices"
    if any(x in text for x in ("pg-x", "filato", "yarn", "raw")):
        return "Raw Yarn"
    if any(x in text for x in ("quality", "qualita", "c.q", "cq")):
        return "Quality"
    if any(x in text for x in ("delay", "ritard", "consegna", "delivery")):
        return "Delay"
    if any(x in text for x in ("file", "upload", "excel", "listini")):
        return "Files"
    if any(x in text for x in ("error", "failed", "system", "save")):
        return "System"
    return "Other"


def list_all() -> list[dict]:
    items = _read()
    for item in items:
        item.setdefault("category", category_for(item.get("key", ""), item.get("title", ""), item.get("page", "")))
        item.setdefault("status", "resolved" if item.get("resolved") else "open")
    return items


def list_open(category: str = "All") -> list[dict]:
    now = datetime.now()
    result = []
    for item in list_all():
        if item.get("resolved") or item.get("status") == "resolved":
            continue
        until = item.get("snoozed_until", "")
        if until:
            try:
                if datetime.fromisoformat(until) > now:
                    continue
            except ValueError:
                pass
        if category != "All" and item.get("category", "Other") != category:
            continue
        result.append(item)
    return sorted(result, key=lambda item: (item.get("severity") != "high", item.get("created_at", "")), reverse=False)


def add_many(entries: Iterable[dict]) -> int:
    items = _read()
    known_keys = {item.get("key") for item in items if not item.get("resolved")}
    now = datetime.now().isoformat(timespec="seconds")
    import re
    added = 0
    for entry in entries:
        key = str(entry["key"])
        if key in known_keys:
            continue
        title = str(entry["title"])
        message = str(entry["message"])
        page = str(entry.get("page", ""))
        severity = str(entry.get("severity", "medium"))
        partita_colore = str(entry.get("partita_colore", ""))
        if not partita_colore or partita_colore in ("None", "nan"):
            match = re.search(r"(?i)(?:Partita|color code|code)\s+([A-Z0-9_\-/]+)", message)
            partita_colore = match.group(1).strip() if match else ""
            if not partita_colore:
                parts = key.split(":")
                if len(parts) >= 3 and parts[0].startswith("situazione-price"):
                    candidate = parts[3] if len(parts) > 3 and not parts[3].replace(".", "").isdigit() else parts[2]
                    if candidate and candidate not in ("None", "nan"):
                        partita_colore = candidate
            partita_colore = partita_colore or "N/A"
        category = str(entry.get("category", ""))
        item_category = category if category in CATEGORIES else category_for(key, title, page)
        items.append({
            "key": key, "title": title, "message": message, "page": page,
            "partita_colore": partita_colore, "category": item_category,
            "severity": severity, "status": "open",
            "created_at": now, "resolved": False, "snoozed_until": "",
        })
        known_keys.add(key)
        added += 1
    if added:
        _write(items)
    return added


def add(key: str, title: str, message: str, page: str, severity: str = "medium", partita_colore: str = "", category: str = "") -> bool:
    return bool(add_many([{
        "key": key, "title": title, "message": message, "page": page,
        "severity": severity, "partita_colore": partita_colore, "category": category,
    }]))


def resolve(key: str) -> None:
    update_status(key, "resolved")


def update_status(key: str, status: str) -> None:
    items = _read()
    changed = False
    for item in items:
        if item.get("key") == key and not item.get("resolved"):
            item["status"] = status
            item["resolved"] = status == "resolved"
            if status != "snoozed":
                item["snoozed_until"] = ""
            changed = True
    if changed:
        _write(items)


def snooze(key: str, hours: int = 24) -> None:
    items = _read()
    until = (datetime.now() + timedelta(hours=hours)).isoformat(timespec="seconds")
    changed = False
    for item in items:
        if item.get("key") == key and not item.get("resolved"):
            item["status"] = "snoozed"
            item["snoozed_until"] = until
            changed = True
    if changed:
        _write(items)
