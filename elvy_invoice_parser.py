"""
elvy_invoice_parser.py — Parses Elvy's raw-yarn "Invoice to Delta" PDFs.

These are a third document type, distinct from both the Purchase Order and
Bolla formats: a 2-page PDF where Elvy sends Delta raw (griege) yarn to be
dyed.

    Page 1 ("Proforma invoice"): Inv. no, Date, and one row per Pos with
        Yarn code (Elvy's article number), NM/Ne, composition, an optional
        Lot reference, Price/Net/Gross/Value.
    Page 2 ("Packing list"): the same Pos entries again, each with its own
        sub-table broken down by pallet/lot, ending in a "Total" line whose
        last number is that Pos's total No. of cones.

Text-extraction quirk
----------------------
This PDF's font/content-stream ordering confuses pdfplumber's normal
extract_words() — it produces nonsense words that splice together
characters from unrelated lines (confirmed by inspection: e.g. the header
lines came out as "ETBB" then "lvhlou..."). The raw page.chars list is
positioned correctly, though, so this module bypasses extract_words()
entirely and reconstructs lines directly from characters, grouped by
rounded top (y) position and split into words wherever the x-gap between
consecutive characters exceeds a small threshold.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pdfplumber

from utils import clean_text, logger, parse_number

Char = dict

LINE_GAP_TOLERANCE = 2.0  # x-gap (pt) above which a space is inserted


@dataclass
class ElvyInvoiceRow:
    """One line-item from an Elvy raw-yarn invoice + its packing-list total."""
    inv_no: str = ""
    inv_date: str = ""
    pos: int = 0
    yarn_code: str = ""
    articolo_delta: str = ""
    nm: str = ""
    ne: str = ""
    yarn_type: str = ""
    lot: str = ""
    price_usd: float | None = None
    net_qty_kg: float | None = None
    gross_qty_kg: float | None = None
    value_usd: float | None = None
    rocche: float | None = None


# ---------------------------------------------------------------------------
# Character-level line reconstruction
# ---------------------------------------------------------------------------

def _reconstruct_lines(page: "pdfplumber.page.Page") -> list[tuple[float, str]]:
    """
    Return [(top, line_text), ...] built directly from page.chars, sorted
    top-to-bottom, with spaces inserted at word-sized x-gaps. See module
    docstring for why this bypasses extract_words().
    """
    chars: list[Char] = [c for c in page.chars if c["text"].strip()]
    chars.sort(key=lambda c: (round(c["top"], 1), c["x0"]))

    grouped: dict[float, list[Char]] = {}
    for c in chars:
        key = round(c["top"], 1)
        grouped.setdefault(key, []).append(c)

    lines: list[tuple[float, str]] = []
    for top in sorted(grouped.keys()):
        row = sorted(grouped[top], key=lambda c: c["x0"])
        parts: list[str] = []
        prev_x1: float | None = None
        for c in row:
            if prev_x1 is not None and c["x0"] - prev_x1 > LINE_GAP_TOLERANCE:
                parts.append(" ")
            parts.append(c["text"])
            prev_x1 = c["x1"]
        lines.append((top, "".join(parts)))
    return lines


def _parse_number(text: str) -> float | None:
    """Parse an invoice number after removing the optional currency sign."""
    return parse_number(clean_text(text).replace("$", ""))


# ---------------------------------------------------------------------------
# Page 1 — Proforma invoice
# ---------------------------------------------------------------------------

RE_HEADER_INV_NO = re.compile(r"Inv\.?\s*no\s*([\d\-]+)", re.IGNORECASE)
RE_HEADER_DATE = re.compile(r"Date\s*:\s*([\d\-]+)", re.IGNORECASE)
RE_YARN_CODE = re.compile(r"Yarn\s*code\s*(\d+)", re.IGNORECASE)
RE_NM_NE = re.compile(r"NM\s*(\S+)\s*Ne\s*(\S+)", re.IGNORECASE)
# The Pos/Color/Price/Qty/Value row: starts with the Pos number, then the
# "000" Color Lot# placeholder, then price/net/gross/$value.
RE_POS_DATA_ROW = re.compile(
    r"^\s*(\d{1,3})\s+\d+\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+\$?\s*([\d.,]+)"
)


def _parse_page1(lines: list[tuple[float, str]]) -> tuple[str, str, list[dict]]:
    """
    Parse the Proforma invoice page. Returns (inv_no, inv_date, blocks),
    where each block is a dict of raw fields for one Pos (not yet resolved
    against the Elvy article mapping).
    """
    full_text = "\n".join(t for _, t in lines)
    inv_no = ""
    inv_date = ""
    m = RE_HEADER_INV_NO.search(full_text)
    if m:
        inv_no = m.group(1)
    m = RE_HEADER_DATE.search(full_text)
    if m:
        inv_date = m.group(1)

    blocks: list[dict] = []
    current: dict | None = None

    for _top, text in lines:
        m = RE_YARN_CODE.search(text)
        if m:
            current = {
                "yarn_code": m.group(1),
                "nm": "", "ne": "",
                "extra_lines": [],
                "price": None, "net": None, "gross": None, "value": None,
            }
            blocks.append(current)
            continue

        if current is None:
            continue

        m = RE_NM_NE.search(text)
        if m:
            current["nm"] = m.group(1)
            current["ne"] = m.group(2)
            continue

        m = RE_POS_DATA_ROW.match(text)
        if m:
            current["price"] = _parse_number(m.group(2))
            current["net"] = _parse_number(m.group(3))
            current["gross"] = _parse_number(m.group(4))
            current["value"] = _parse_number(m.group(5))
            continue

        if text.strip().lower().startswith("total"):
            current = None
            continue

        # Anything else between the NM/Ne line and the next block boundary
        # is composition text (Yarn Type) and/or a Lot reference line.
        current["extra_lines"].append(text.strip())

    for block in blocks:
        extra = block.pop("extra_lines")
        block["yarn_type"] = extra[0] if extra else ""
        block["lot"] = " ".join(extra[1:]) if len(extra) > 1 else ""

    return inv_no, inv_date, blocks


# ---------------------------------------------------------------------------
# Page 2 — Packing list (for the Rocche / No. of cones total per Pos)
# ---------------------------------------------------------------------------

RE_TOTAL_LINE = re.compile(
    r"^\s*Total\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$", re.IGNORECASE
)


def _parse_page2(lines: list[tuple[float, str]]) -> list[float | None]:
    """
    Parse the Packing list page. Returns the list of Rocche (total No. of
    cones) values, one per Pos block, in page order.
    """
    cones: list[float | None] = []
    in_block = False

    for _top, text in lines:
        if RE_YARN_CODE.search(text):
            in_block = True
            continue
        if not in_block:
            continue
        m = RE_TOTAL_LINE.match(text)
        if m:
            cones.append(_parse_number(m.group(3)))
            in_block = False

    return cones


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

class ElvyInvoiceParser:
    """Parses one Elvy raw-yarn invoice (Proforma invoice + Packing list) PDF."""

    def __init__(self, pdf_path: Path) -> None:
        self.pdf_path = Path(pdf_path)

    def parse(self) -> list[ElvyInvoiceRow]:
        logger.info("Opening Elvy Invoice PDF: %s", self.pdf_path.name)

        with pdfplumber.open(self.pdf_path) as pdf:
            if not pdf.pages:
                return []
            page1_lines = _reconstruct_lines(pdf.pages[0])
            inv_no, inv_date, blocks = _parse_page1(page1_lines)

            cones: list[float | None] = []
            if len(pdf.pages) > 1:
                page2_lines = _reconstruct_lines(pdf.pages[1])
                cones = _parse_page2(page2_lines)

        if len(cones) != len(blocks):
            logger.warning(
                "  %s: %d item(s) on page 1 but %d Rocche total(s) found on "
                "page 2 — matching what's available in order; check the "
                "output for a possible mismatch.",
                self.pdf_path.name, len(blocks), len(cones),
            )

        rows: list[ElvyInvoiceRow] = []
        for i, block in enumerate(blocks):
            rows.append(ElvyInvoiceRow(
                inv_no=inv_no,
                inv_date=inv_date,
                pos=i + 1,
                yarn_code=block["yarn_code"],
                nm=block["nm"],
                ne=block["ne"],
                yarn_type=block["yarn_type"],
                lot=block["lot"],
                price_usd=block["price"],
                net_qty_kg=block["net"],
                gross_qty_kg=block["gross"],
                value_usd=block["value"],
                rocche=cones[i] if i < len(cones) else None,
            ))

        logger.info("  %d row(s) parsed — %s", len(rows), self.pdf_path.name)
        return rows
