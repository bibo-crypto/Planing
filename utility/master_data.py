"""Local reference data store.

This is deliberately JSON-backed until the central database phase is approved.
"""
from __future__ import annotations

import json
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
    # Delave colours (Colore starting with "#") get their own Articolo,
    # which is unrelated to the raw yarn's own Articolo -- unlike a normal
    # colour, where the raw yarn's Articolo is always the same digits with
    # "C" swapped for "G" (C010032S -> G010032S). Every place that resolves
    # a finished colour's Articolo back to its raw yarn (Show Orders'
    # PG-X/shortage matching, the Filato x Tinturia export, and the Edit
    # PG-X dialog) needs this table checked first, before falling back to
    # the plain C->G rule. One row per Delave Articolo actually seen.
    "delave_map": [],
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
    _delave_maps_cache.clear()


_delave_maps_cache: dict[str, dict[str, str]] = {}


def _delave_maps() -> tuple[dict[str, str], dict[str, str]]:
    """(delave Articolo -> raw Articolo, raw Articolo -> delave Articolo),
    rebuilt from disk only after the first call or after save() invalidates it."""
    if not _delave_maps_cache:
        forward: dict[str, str] = {}
        backward: dict[str, str] = {}
        for row in load().get("delave_map", []):
            delave = str(row.get("delave_articolo", "")).strip().upper()
            raw = str(row.get("raw_articolo", "")).strip().upper()
            if delave and raw:
                forward[delave] = raw
                backward.setdefault(raw, delave)
        _delave_maps_cache["forward"] = forward
        _delave_maps_cache["backward"] = backward
    return _delave_maps_cache["forward"], _delave_maps_cache["backward"]


def raw_articolo_for(articolo: str) -> str:
    """Resolve a finished colour's Articolo to its raw-yarn Articolo.

    Checks the Delave override table first (see DEFAULTS["delave_map"]),
    since a Delave colour's Articolo does not share its digits with the
    raw yarn's the way a normal colour's does. Falls back to the plain
    C -> G digit-preserving rule used everywhere else (C010032S ->
    G010032S) when there is no override, or the value is not C-prefixed.
    """
    a = str(articolo or "").strip().upper()
    if not a:
        return ""
    forward, _backward = _delave_maps()
    if a in forward:
        return forward[a]
    if a.startswith("C") and len(a) > 1:
        return "G" + a[1:]
    return a


def finished_articolo_for(raw_articolo: str) -> str:
    """Reverse of raw_articolo_for(): guess a finished colour's Articolo
    from a raw yarn's Articolo, e.g. for suggesting Articolo choices from
    Magazino Filato stock in the Edit PG-X dialog. Delave override first,
    then the plain G -> C rule."""
    a = str(raw_articolo or "").strip().upper()
    if not a:
        return ""
    _forward, backward = _delave_maps()
    if a in backward:
        return backward[a]
    if a.startswith("G") and len(a) > 1:
        return "C" + a[1:]
    return a
