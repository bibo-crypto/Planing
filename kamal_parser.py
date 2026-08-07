"""
kamal_parser.py
Parses Kamal Textile's order letters ("جواب") -- a different PDF layout
from Elvy's Purchase Orders (no Article No column; colours are identified
by an optional COLOREDFM code and/or yarn count "نمرة الغزل", which
doubles as the Titolo used to find the matching Articolo in DFM).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pdfplumber

from utils import clean_text, logger

RE_DYE_TYPE = re.compile(r"\b(REACTIVE|VAT)\b", re.IGNORECASE)
_ARABIC_INDIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
_ARABIC_CHAR_RANGE = re.compile(r"[\u0600-\u06FF]")


def _fix_rtl(text: str) -> str:
    """
    pdfplumber extracts Arabic table text in visual (mirrored) order --
    e.g. 'محمد علي' comes out as 'يلع دمحم'. Fix it per line: reverse the
    word order, and reverse the characters within each Arabic word (but
    not numbers/Latin words, which are already in the right order).
    """
    if not text:
        return text

    def _is_arabic(token: str) -> bool:
        return bool(_ARABIC_CHAR_RANGE.search(token))

    fixed_lines = []
    for line in text.split("\n"):
        if not _ARABIC_CHAR_RANGE.search(line):
            fixed_lines.append(line)  # no Arabic here -- already in correct order, leave as-is
            continue
        tokens = line.split(" ")[::-1]
        fixed_lines.append(" ".join(t[::-1] if _is_arabic(t) else t for t in tokens))
    return "\n".join(fixed_lines)


@dataclass
class KamalOrderRow:
    """One colour line from a Kamal order letter."""
    reply_no: str = ""
    order_date: str = ""          # as printed, Western-digit form, e.g. "2026/07/20"
    dye_type: str = ""            # "REACTIVE" or "VAT"
    notes: str = ""
    sender_name: str = ""
    sender_msg_no: str = ""
    coloredfm: str = ""
    titolo: str = ""              # "نمرة الغزل", e.g. "80/2" -- blank for many rows
    colore_name: str = ""         # "اللون", Kamal's own colour label
    peso_kg: float = 0.0          # "الوزن", in kg


def _to_western_digits(text: str) -> str:
    return clean_text(text).translate(_ARABIC_INDIC_DIGITS)


def _extract_reply_no(page) -> str:
    for table in page.extract_tables() or []:
        for row in table:
            cells = [clean_text(c) for c in row if c]
            if any("مقر باوج" in c or "باوج" in c and "مقر" in c for c in cells):
                # cells look like ['1/2', '98', 'باوج مقر'] -- the reply
                # number is the numeric cell that isn't the page marker "x/y"
                for c in cells:
                    if c.isdigit():
                        return c
    return ""


def _extract_order_date(page) -> str:
    for table in page.extract_tables() or []:
        for row in table:
            cells = [clean_text(c) for c in row if c]
            if any("خيرات" in c for c in cells):
                for c in cells:
                    western = _to_western_digits(c)
                    if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", western):
                        return western
    return ""


def _extract_dye_type(text: str) -> str:
    m = RE_DYE_TYPE.search(text or "")
    return m.group(1).upper() if m else ""


def _extract_year(order_date: str) -> str:
    m = re.search(r"(20\d{2})", order_date or "")
    return m.group(1) if m else str(datetime.now().year)


def _is_data_table(table: list) -> bool:
    """The colour table has a header row containing 'نزولا' (weight, reversed)."""
    for row in table[:1]:
        cells = " ".join(clean_text(c) for c in row if c)
        if "نزولا" in cells or "COLOREDFM" in cells:
            return True
    return False


class KamalParser:
    """Parses a single Kamal order-letter PDF into a list of KamalOrderRow."""

    def __init__(self, pdf_path: Path):
        self.pdf_path = pdf_path

    def parse(self) -> list[KamalOrderRow]:
        logger.info("Opening Kamal order PDF: %s", self.pdf_path.name)
        rows: list[KamalOrderRow] = []

        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                reply_no = _extract_reply_no(page)
                order_date = _extract_order_date(page)
                dye_type = _extract_dye_type(text)

                for table in page.extract_tables() or []:
                    if not _is_data_table(table):
                        continue
                    for raw_row in table[1:]:  # skip the header row
                        cells = [clean_text(c) for c in raw_row]
                        while len(cells) < 8:
                            cells.append("")
                        notes, sender_name, sender_msg_no, coloredfm, _sample, titolo, colore_name, peso = cells[:8]
                        notes = _fix_rtl(notes)
                        sender_name = _fix_rtl(sender_name)
                        colore_name = _fix_rtl(colore_name)

                        if not colore_name and not peso:
                            continue  # skip fully blank rows

                        try:
                            peso_kg = float(_to_western_digits(peso)) if peso else 0.0
                        except ValueError:
                            peso_kg = 0.0

                        rows.append(KamalOrderRow(
                            reply_no=reply_no,
                            order_date=order_date,
                            dye_type=dye_type,
                            notes=notes,
                            sender_name=sender_name,
                            sender_msg_no=_to_western_digits(sender_msg_no),
                            coloredfm=_to_western_digits(coloredfm),
                            titolo=titolo,
                            colore_name=colore_name,
                            peso_kg=peso_kg,
                        ))

        logger.info("  %d row(s) parsed — %s", len(rows), self.pdf_path.name)
        return rows
