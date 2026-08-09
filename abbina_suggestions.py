"""Suggestions for combining dyeing batches from the Situazione table."""

from __future__ import annotations

import math
import re
from collections import defaultdict

import pandas as pd


MACHINE_CAPACITIES = [6, 24, 32, 56, 72, 128, 192, 384, 672]
MACHINE_NUMBERS = {
    6: 11, 24: 12, 32: 9, 56: 10, 72: 7,
    128: 8, 192: 5, 384: 6, 672: 3,
}
SPECIAL_TITLES = {"30/1", "31/1", "36/1", "9.5/1", "23/1", "15,4/1",
                  "26,8/1", "38/1", "8,3/1", "35,4/1"}


def _text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _title_plies(title):
    """Return the ply/count after '/' (1, 2, etc.) from a Titolo."""
    match = re.match(r"^\s*[^/]+/\s*([0-9]+)", _text(title))
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def titles_compatible(left, right):
    """Return whether two yarn titles may be dyed in one bath."""
    a = _text(left).replace(" ", "").casefold()
    b = _text(right).replace(" ", "").casefold()
    if not a or not b:
        return False
    if a in {x.casefold() for x in SPECIAL_TITLES} or b in {x.casefold() for x in SPECIAL_TITLES}:
        return a == b
    plies_a, plies_b = _title_plies(a), _title_plies(b)
    if plies_a is None or plies_b is None:
        return a == b
    return plies_a == plies_b


def _number(value):
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def build_suggestions(df: pd.DataFrame, max_extra_percent: float = 0.20) -> pd.DataFrame:
    """Build one detail row per batch that can be combined with another batch.

    Groups above the 30% raw-addition limit are retained and marked as a
    warning so the planner can see the alternative without accidentally
    treating it as a normal recommendation.
    """
    columns = ["titolo", "codice", "colore", "rocche", "partita", "bagno",
               "abbina", "tot_rocche", "mc_target", "polmoni", "extra_percent", "motivo",
               "new_comment"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)

    work = df.copy().reset_index(drop=True)
    for col in ("titolo", "codice", "colore", "partita", "bagno", "mc"):
        if col not in work:
            work[col] = ""
    if "rocche" not in work:
        work["rocche"] = 0

    # Abbina is only safe for virgin batches that are still waiting to be
    # dyed: C.Q must still be OO and New Comment must be a production status.
    # This excludes quality-checked, shipped, and re-dye (Ritinta) batches.
    if "cq" not in work:
        work["cq"] = ""
    if "new_comment" not in work:
        work["new_comment"] = ""
    eligible = (
        work["cq"].map(_text).str.upper().eq("OO")
        & work["new_comment"].map(_text).str.upper().str.startswith("PROD")
    )
    work = work[eligible].reset_index(drop=True)
    if work.empty:
        return pd.DataFrame(columns=columns)

    # Same codice and same dye colour are required. Within that bucket, make
    # compatible title families (odd/even, with the explicit exceptions).
    buckets = defaultdict(list)
    for index, row in work.iterrows():
        code = _text(row.get("codice"))
        colour = _text(row.get("colore"))
        if code:
            buckets[(code, colour)].append(index)

    output = []
    for (code, colour), indexes in buckets.items():
        families = []
        for index in indexes:
            title = work.at[index, "titolo"]
            family = next((group for group in families if titles_compatible(title, work.at[group[0], "titolo"])), None)
            if family is None:
                family = []
                families.append(family)
            family.append(index)

        for family in families:
            if len(family) < 2:
                continue

            # Rows sharing one Bagno are already in the same dyeing bath.
            # Treat that Bagno as one batch, so its quantity is not counted
            # twice and a group made only of one Bagno is not suggested.
            batch_totals = defaultdict(float)
            for index in family:
                bagno = _text(work.at[index, "bagno"])
                batch_key = f"bagno:{bagno}" if bagno else f"row:{index}"
                batch_totals[batch_key] += max(0.0, _number(work.at[index, "rocche"]))
            if len(batch_totals) < 2:
                continue
            total = sum(batch_totals.values())
            if total <= 0:
                continue
            target = next((capacity for capacity in MACHINE_CAPACITIES if capacity >= total), None)
            if target is None:
                continue
            polmoni = max(0, math.ceil((target - total) / 4))
            extra_percent = (polmoni * 4) / total
            machine_number = MACHINE_NUMBERS[target]
            polmoni_text = f" + {polmoni} Polmoni" if polmoni else ""
            recommendation = f"M/C {machine_number} ({target}){polmoni_text}"
            motivo = f"{len(family)} lotti: {total:g} rocche -> {target:g}"
            if polmoni:
                motivo += f"; extra {polmoni * 4:g} ({extra_percent:.1%})"
            if extra_percent > max_extra_percent + 1e-9:
                motivo += f"; ⚠ oltre {max_extra_percent:.0%}"
            for index in family:
                row = work.loc[index]
                output.append({
                    "titolo": _text(row.get("titolo")), "codice": code, "colore": colour,
                    "rocche": row.get("rocche", ""), "partita": _text(row.get("partita")),
                    "bagno": _text(row.get("bagno")), "abbina": recommendation,
                    "tot_rocche": total, "mc_target": target, "polmoni": polmoni,
                    "extra_percent": extra_percent, "motivo": motivo,
                    "new_comment": _text(row.get("new_comment")),
                })
    return pd.DataFrame(output, columns=columns)
