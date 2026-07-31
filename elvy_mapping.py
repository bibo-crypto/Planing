"""
elvy_mapping.py — Persisted lookup table mapping Elvy article codes to
Delta article codes.

Elvy orders reference items using their own "Article No" numbering — the
same field already extracted as OrderRow.article_no from Purchase Order
PDFs. Delta uses its own internal article codes for the same items, and
there's no way to derive one from the other automatically. This module
lets the user maintain that mapping by hand (via the "Elvy" tab in the
GUI) and looks it up when exporting Purchase Orders, so a row's "Article
No" (Elvy) is joined to its equivalent "Articolo Delta" — similar to a
spreadsheet VLOOKUP / merge-query by Article No.

Storage: a simple JSON file in the same writable per-user AppData
directory used for settings/logs (see utils.APP_DATA_DIR), keyed by the
Elvy article code so each Elvy code maps to exactly one Delta code.
"""

from __future__ import annotations

import json

from utils import APP_DATA_DIR, clean_text, logger

MAPPING_FILE = APP_DATA_DIR / "settings" / "elvy_mapping.json"

# Seed values for a first run.  Once the file exists, user edits and deletions
# are preserved exactly; defaults are not merged back into an existing file.
DEFAULT_ELVY_MAPPING: dict[str, str] = {
    "420": "G130008S", "423": "G130054S", "425": "G130136S",
    "427": "G130008S", "501": "G130157S", "601": "G130078S",
    "607": "G130078S", "700": "G130179S", "709": "G130098S",
    "800": "G130026S", "810": "G130027S", "821": "G130025S",
    "822": "G130348S", "823": "G130025S", "825": "G130365S",
    "828": "G130273S", "831": "G130348S", "840": "G130394S",
    "850": "G130947S", "852": "G130947S", "860": "G130229S",
    "875": "G130154S", "940": "G130055S", "4026": "G130015S",
    "4039": "G130034S", "4052": "G130036S", "8050": "G130038S",
    "9714": "G130097S", "9817": "G130027S", "9838": "G130025S",
}

def load_elvy_mapping() -> dict[str, str]:
    """
    Load the persisted Elvy→Delta article mapping.
    Returns an empty dict if the file doesn't exist yet or can't be read.
    """
    try:
        if MAPPING_FILE.is_file():
            on_disk = json.loads(MAPPING_FILE.read_text(encoding="utf-8"))
            if isinstance(on_disk, dict):
                return {str(k): str(v) for k, v in on_disk.items()}
        save_elvy_mapping(DEFAULT_ELVY_MAPPING.copy())
        return DEFAULT_ELVY_MAPPING.copy()
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load Elvy mapping: %s", exc)
    return {}


def save_elvy_mapping(mapping: dict[str, str]) -> None:
    """Persist *mapping* to disk. Creates the directory if needed."""
    try:
        MAPPING_FILE.parent.mkdir(parents=True, exist_ok=True)
        MAPPING_FILE.write_text(
            json.dumps(mapping, indent=2, ensure_ascii=False, sort_keys=True),
            encoding="utf-8",
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save Elvy mapping: %s", exc)


def add_elvy_mapping(articolo_elvy: str, articolo_delta: str) -> dict[str, str]:
    """
    Add or update one (Articolo Elvy -> Articolo Delta) entry and persist
    it. Returns the full updated mapping.
    """
    mapping = load_elvy_mapping()
    key = clean_text(articolo_elvy)
    if key:
        mapping[key] = clean_text(articolo_delta)
        save_elvy_mapping(mapping)
    return mapping


def delete_elvy_mapping(articolo_elvy: str) -> dict[str, str]:
    """Remove one entry (if present) and persist. Returns the updated mapping."""
    mapping = load_elvy_mapping()
    mapping.pop(clean_text(articolo_elvy), None)
    save_elvy_mapping(mapping)
    return mapping


def lookup_articolo_delta(article_no: str, mapping: dict[str, str]) -> str:
    """
    Look up *article_no* (an OrderRow's Article No, i.e. Articolo Elvy) in
    *mapping* and return the matching Articolo Delta, or "" if not found.
    """
    return mapping.get(clean_text(article_no), "")
