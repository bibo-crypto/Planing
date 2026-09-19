"""Pure helpers for the PG-X raw-yarn demand report and its schedule."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Any

import pandas as pd


def _number(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def pg_x_demand(records: Iterable[Any]) -> pd.DataFrame:
    """Aggregate open PG-X rows by Titolo, retaining article/client detail."""
    rows = []
    for record in records:
        raw_batch = str(getattr(record, "raw_batch", "") or "").strip().upper().replace(" ", "")
        if raw_batch not in {"", "X", "PG-X", "PGX"}:
            continue
        rows.append({
            "Titolo": str(getattr(record, "title", "") or getattr(record, "description", "") or "").strip(),
            "Articolo": str(getattr(record, "article", "") or "").strip(),
            "Cliente": str(getattr(record, "customer_name", "") or "").strip(),
            "Rocche": _number(getattr(record, "quantity_cones", 0)),
            "KG": _number(getattr(record, "kg", 0)),
            "Partita Col": str(getattr(record, "colored_batch", "") or "").strip(),
        })
    columns = ["Titolo", "Articolo", "Cliente", "Rocche", "KG", "Partita Col"]
    if not rows:
        return pd.DataFrame(columns=columns)
    detail = pd.DataFrame(rows)
    return (detail.groupby(["Titolo", "Articolo", "Cliente"], dropna=False, as_index=False)
            .agg({"Rocche": "sum", "KG": "sum", "Partita Col": lambda values: ", ".join(str(v) for v in values if str(v).strip())})
            .sort_values(["Titolo", "Articolo", "Cliente"], kind="stable")
            .reset_index(drop=True))


def schedule_due(schedule: dict, now: datetime | None = None) -> bool:
    """Return whether a daily/weekly report is due, without mutating state."""
    if not schedule or not schedule.get("enabled") or not schedule.get("recipient"):
        return False
    now = now or datetime.now()
    if str(schedule.get("frequency", "daily")).lower() == "weekly":
        day = int(schedule.get("weekday", now.weekday()))
        if now.weekday() != day:
            return False
    target = str(schedule.get("time", "08:00"))
    try:
        hour, minute = (int(part) for part in target.split(":", 1))
    except (ValueError, TypeError):
        return False
    if (now.hour, now.minute) != (hour, minute):
        return False
    last = str(schedule.get("last_sent", ""))
    return last != now.date().isoformat()


def report_path(shared_excel: Path) -> Path:
    return shared_excel.with_name("PG-X Orders Report.xlsx")


def export_report(frame: pd.DataFrame, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(path, index=False, sheet_name="PG-X Demand")
    return path


def schedule_stamp(now: datetime | None = None) -> str:
    return (now or datetime.now()).date().isoformat()


def today() -> date:
    return datetime.now().date()
