"""Cross-cutting reports built on top of the persisted Situazione state
(utility/situazione_db.py) -- on-time delivery scoring and the Partita
timeline/audit trail. Pure dataframe logic, no Tkinter.
"""
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


def format_partita_timeline(history_rows: list[dict]) -> pd.DataFrame:
    """Turn raw partita_history rows into a display-ready timeline: one row
    per event, with a plain-language description of what changed since the
    previous snapshot.
    """
    from calculate.situazione import compute_delay_days, _parse_delivery_date

    columns = ["changed_at", "event", "comment", "bagno", "tinto", "data_qualita", "data_uscita", "days_in_qc", "ritardo"]
    if not history_rows:
        return pd.DataFrame(columns=columns)

    tracked = [
        ("bagno", "Assigned to Bagno {new}"),
        ("tinto", "Dyed on {new}"),
        ("data_qualita", "Passed Q.C. on {new}"),
        ("data_uscita", "Shipped on {new}"),
    ]
    out_rows = []
    prev: dict = {}
    today = pd.Timestamp(datetime.now().date())

    for i, row in enumerate(history_rows):
        events = []
        if i == 0:
            events.append("First seen")
        else:
            for field, template in tracked:
                new_val = str(row.get(field) or "").strip()
                old_val = str(prev.get(field) or "").strip()
                if new_val and new_val != old_val:
                    events.append(template.format(new=new_val))
            if row.get("new_comment") != prev.get("new_comment"):
                events.append(f"Status -> {row.get('new_comment') or '(cleared)'}")
        if not events:
            events.append("Updated")

        days_qc = str(row.get("days_in_qc") or "").strip()
        if not days_qc:
            tinto_dt = _parse_delivery_date(row.get("tinto"))
            if tinto_dt is not None:
                days_qc = str((today - tinto_dt).days)
            else:
                days_qc = ""

        ritardo_val = str(row.get("ritardo_consegna") or row.get("ritardo") or "").strip()
        if not ritardo_val:
            ritardo_val = compute_delay_days(row)

        out_rows.append({
            "changed_at": row.get("changed_at"),
            "event": "; ".join(events),
            "comment": row.get("new_comment"),
            "bagno": row.get("bagno"),
            "tinto": row.get("tinto"),
            "data_qualita": row.get("data_qualita"),
            "data_uscita": row.get("data_uscita"),
            "days_in_qc": days_qc,
            "ritardo": ritardo_val,
        })
        prev = row
    return pd.DataFrame(out_rows, columns=columns)
