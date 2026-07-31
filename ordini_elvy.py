"""
ordini_elvy.py — Derives the "Ordini ELVY" sheet from parsed Purchase Order
rows (see pdf_parser.OrderRow), for the Elvy raw-yarn dyeing workflow.

Column mapping (fixed values confirmed with the user; header labels and
types match EXCEL_PER_ORDINE_VENDITA_EGITTO.xlsx, the real ERP import
template, so this sheet can be pasted straight into it)
--------------------------------------------------------------------------
CLIENTE              — always 3009 (number).
ARTICOLO             — the row's raw code (already G-prefixed via the Elvy
                        mapping) converted to "C" — it's entering dyeing.
COLORE                — copied from the row's DFM colour lookup (number).
Q.TA                 — copied as-is (number).
CONSEGNA             — today + 14 days (real date).
COMMENTO             — "PG-X-PO-<POD No>-<year>", year from the row's Po Date.
LAVORANTE            — always 900901 (number).
LAV. SUCC            — always 900161 (number).
DATA RICONSEGNA      — CONSEGNA - 1 day (real date).
MAG. GREGGIO         — always 900923 (number).
SIGLA DISPOSIZIONE   — always "D".
BAGNO PREPOSTO       — always blank.
GRUPPO MACCHINA      — the physical machine number for the row's Abbina
                        cone count (see abbina_calculator.MACHINE_CODES),
                        as a number.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from abbina_calculator import MACHINE_CODES
from pdf_parser import OrderRow
from utils import clean_text

CLIENTE_CODE = 3009
LAVORANTE_CODE = 900901
LAVORANTE_SUCCESSIVO_CODE = 900161
STOCK_MAG_CODE = 900923
SIGLA_DISPO_CODE = "D"
CONSEGNA_LEAD_DAYS = 14

RE_ABBINA_CONES = re.compile(r"Machine\s+(\d+)\s*Cones?", re.IGNORECASE)
RE_YEAR = re.compile(r"(\d{4})")


@dataclass
class OrdiniElvyRow:
    """One derived row for the "Ordini ELVY" sheet."""
    cliente: int = CLIENTE_CODE
    articolo_delta: str = ""
    coloredfm: int | None = None
    quantity_cones: float | None = None
    data_consegna: datetime | None = None
    commento: str = ""
    lavorante: int = LAVORANTE_CODE
    lavorante_successivo: int = LAVORANTE_SUCCESSIVO_CODE
    data_riconsegna: datetime | None = None
    stock_mag: int = STOCK_MAG_CODE
    sigla_dispo: str = SIGLA_DISPO_CODE
    bagno_proposto: str = ""
    macchina: int | None = None


def _to_c_code(articolo_delta: str) -> str:
    """Convert a G-prefixed raw code to its C-prefixed dyeing-entry form."""
    articolo_delta = clean_text(articolo_delta)
    if articolo_delta.upper().startswith("G"):
        return "C" + articolo_delta[1:]
    return ""


def _extract_year(po_date: str) -> str:
    """Pull the 4-digit year out of a Po Date string like '19/4/2026'."""
    m = RE_YEAR.search(clean_text(po_date))
    return m.group(1) if m else ""


def _to_int(value: str) -> int | None:
    value = clean_text(value)
    try:
        return int(value)
    except ValueError:
        return None


def _machine_code_for_abbina(abbina: str) -> int | None:
    """Look up the MACCHINA code from an Abbina value like 'Machine 72 Cones'."""
    m = RE_ABBINA_CONES.search(clean_text(abbina))
    if not m:
        return None
    return MACHINE_CODES.get(int(m.group(1)))


def build_ordini_elvy_rows(rows: list[OrderRow]) -> list[OrdiniElvyRow]:
    """Build one OrdiniElvyRow per OrderRow, using today's date as the basis
    for CONSEGNA / DATA RICONSEGNA."""
    today = datetime.now()
    data_consegna_dt = today + timedelta(days=CONSEGNA_LEAD_DAYS)
    data_riconsegna_dt = data_consegna_dt - timedelta(days=1)

    out: list[OrdiniElvyRow] = []
    for r in rows:
        year = _extract_year(r.po_date)
        commento = f"PG-X-PO-{clean_text(r.pod_no)}-{year}" if year else f"PG-X-PO-{clean_text(r.pod_no)}"

        out.append(OrdiniElvyRow(
            articolo_delta=_to_c_code(r.articolo_delta),
            coloredfm=_to_int(r.coloredfm),
            quantity_cones=r.quantity_cones,
            data_consegna=data_consegna_dt,
            commento=commento,
            data_riconsegna=data_riconsegna_dt,
            macchina=_machine_code_for_abbina(r.abbina),
        ))
    return out


# ---------------------------------------------------------------------------
# Updating an existing external ERP file in place
# ---------------------------------------------------------------------------

# Header text (normalized: stripped, case-insensitive) -> OrdiniElvyRow
# attribute. Matches EXCEL_PER_ORDINE_VENDITA_EGITTO.xlsx's own headers, so
# the target file's column ORDER doesn't have to match ours exactly — only
# the header text does.
_HEADER_TO_ATTR = {
    "cliente": "cliente",
    "articolo": "articolo_delta",
    "colore": "coloredfm",
    "q.ta": "quantity_cones",
    "consegna": "data_consegna",
    "commento": "commento",
    "lavorante": "lavorante",
    "lav. succ": "lavorante_successivo",
    "data riconsegna": "data_riconsegna",
    "mag. greggio": "stock_mag",
    "sigla disposizione": "sigla_dispo",
    "bagno preposto": "bagno_proposto",
    "gruppo macchina": "macchina",
}


def update_existing_ordini_file(target_path: Path, ordini_rows: list[OrdiniElvyRow]) -> int:
    """
    Open the existing ERP import file at *target_path*, clear every data
    row (everything from row 2 down — the header row is left untouched),
    write *ordini_rows* starting at row 2 matched to that file's own
    header text, save, and close.

    Returns the number of rows written. Raises if the file can't be
    opened (e.g. it's currently open in Excel) or has no recognizable
    header row.
    """
    import openpyxl  # local import: this module doesn't need openpyxl otherwise

    wb = openpyxl.load_workbook(target_path)
    ws = wb.active

    header_cells = next(ws.iter_rows(min_row=1, max_row=1))
    col_attr: dict[int, str] = {}
    for cell in header_cells:
        if cell.value is None:
            continue
        key = str(cell.value).strip().lower()
        attr = _HEADER_TO_ATTR.get(key)
        if attr:
            col_attr[cell.column] = attr

    if not col_attr:
        wb.close()
        raise ValueError(
            "None of this file's column headers match the expected Ordini "
            "ELVY headers (CLIENTE, ARTICOLO, COLORE, ...) — is this the "
            "right file?"
        )

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for row_offset, row_data in enumerate(ordini_rows):
        row_idx = 2 + row_offset
        for col_idx, attr in col_attr.items():
            value = getattr(row_data, attr, None)
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            if isinstance(value, datetime):
                cell.number_format = "DD/MM/YYYY"

    wb.save(target_path)
    wb.close()
    return len(ordini_rows)
