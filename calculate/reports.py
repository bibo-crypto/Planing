"""Cross-cutting reports built on top of the persisted Situazione state."""
from __future__ import annotations

from datetime import datetime

import pandas as pd


def _parse_date(value: str):
    value = str(value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None


def compute_on_time_delivery(states: dict) -> pd.DataFrame:
    """One row per Cliente: how many Partite have actually shipped
    (Data Uscita filled in) and, of those, what fraction shipped on or
    before the promised Consegna date.

    A Partita only counts once both Consegna and Data Uscita are known --
    still-open batches (no Data Uscita yet) aren't judged either way, since
    they haven't missed anything yet.
    """
    records = []
    for row in states.values():
        consegna = _parse_date(row.get("consegna"))
        uscita = _parse_date(row.get("data_uscita"))
        if consegna is None or uscita is None:
            continue
        delay_days = (uscita - consegna).days
        records.append({
            "cliente": row.get("cliente") or "(no client)",
            "partita": row.get("partita"),
            "on_time": delay_days <= 0,
            "delay_days": delay_days,
        })

    columns = ["cliente", "shipped", "on_time", "late", "on_time_pct", "avg_delay_days"]
    if not records:
        return pd.DataFrame(columns=columns)

    df = pd.DataFrame(records)
    summary = df.groupby("cliente").agg(
        shipped=("partita", "count"),
        on_time=("on_time", "sum"),
    ).reset_index()
    summary["late"] = summary["shipped"] - summary["on_time"]
    summary["on_time_pct"] = (summary["on_time"] / summary["shipped"] * 100).round(1)

    # Average delay across the LATE shipments only -- mixing in the on-time
    # ones (negative/zero delay) would water down the number that actually
    # matters: "when we're late, by how much?"
    late_only = df[~df["on_time"]]
    avg_delay = late_only.groupby("cliente")["delay_days"].mean().round(1)
    summary["avg_delay_days"] = summary["cliente"].map(avg_delay).fillna(0.0)

    return summary.sort_values("on_time_pct", ascending=True).reset_index(drop=True)[columns]


