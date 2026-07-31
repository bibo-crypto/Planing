"""
bolla_parser.py — Parses "Bolla" (Italian delivery-note / packing-list) PDFs.

This is a different document type from the main dyeing Purchase Orders:
one row per pallet/box, plus a small totals summary block at the end of
each document.

Header line (repeats on every page):
    Pallet Scatola Disposizione Articolo Descrizione Colore Partita
    Rocche KgNetto KgLordo Famiglia

Totals block (appears once per document, usually on the last page):
    Totale N.Colli N.Pallet TaraTub. TaraScat. Rocche Netto Lordo
    <values...>
    Totale Generale ,00
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pdfplumber

from utils import clean_text, logger

Word = dict  # a pdfplumber word dict: {"text", "x0", "x1", "top", ...}

# ---------------------------------------------------------------------------
# Header / label regexes
# ---------------------------------------------------------------------------

RE_BOLLA_HEADER = re.compile(
    r"Bolla\s+Numero:\s*(?P<no>.+?)\s+del\s+(?P<date>\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)

ITEM_HEADER_LABELS = {
    "Pallet": "pallet",
    "Scatola": "scatola",
    "Disposizione": "disposizione",
    "Articolo": "articolo",
    "Descrizione": "descrizione",
    "Colore": "colore",
    "Partita": "partita",
    "Rocche": "rocche",
    "KgNetto": "kg_netto",
    "KgLordo": "kg_lordo",
    "Famiglia": "famiglia",
}

TOTALS_HEADER_LABELS = {
    "N.Colli": "n_colli",
    "N.Pallet": "n_pallet",
    "TaraTub.": "tara_tub",
    "TaraScat.": "tara_scat",
    "Rocche": "rocche",
    "Netto": "netto",
    "Lordo": "lordo",
}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class BollaRow:
    """One Pallet/box line item from a Bolla document."""
    bolla_no: str = ""
    bolla_date: str = ""
    pallet: str = ""
    scatola: str = ""
    disposizione: str = ""
    articolo: str = ""
    descrizione: str = ""
    colore: str = ""
    partita: str = ""
    rocche: float | None = None
    kg_netto: float | None = None
    kg_lordo: float | None = None
    famiglia: str = ""


@dataclass
class BollaTotals:
    """The 'Totale' summary line for one Bolla document."""
    bolla_no: str = ""
    bolla_date: str = ""
    n_colli: float | None = None
    n_pallet: float | None = None
    tara_tub: float | None = None
    tara_scat: float | None = None
    rocche: float | None = None
    netto: float | None = None
    lordo: float | None = None


@dataclass
class BollaGroupedRow:
    """
    One grouped summary row: all BollaRow items sharing the same
    (Articolo, Descrizione, Colore, Partita) combination, collapsed into a
    single row with Rocche and KgNetto summed across them.
    """
    articolo: str = ""
    descrizione: str = ""
    colore: str = ""
    partita: str = ""
    rocche: float | None = None
    kg_netto: float | None = None


def group_rows_by_lot(rows: "list[BollaRow]") -> "list[BollaGroupedRow]":
    """
    Collapse *rows* into one row per distinct (Articolo, Descrizione,
    Colore, Partita) combination, summing Rocche and KgNetto across every
    row that shares all four values exactly.

    Intended to be called on the full set of rows being exported together
    (a single document, or several merged documents) — if the same lot
    shows up across multiple merged Bolla files, its rows are summed
    together too, since the grouping key doesn't include bolla_no.
    """
    order: list[tuple[str, str, str, str]] = []
    sums: dict[tuple[str, str, str, str], dict[str, float]] = {}

    for r in rows:
        key = (r.articolo, r.descrizione, r.colore, r.partita)
        if key not in sums:
            sums[key] = {"rocche": 0.0, "kg_netto": 0.0}
            order.append(key)
        sums[key]["rocche"] += r.rocche or 0.0
        sums[key]["kg_netto"] += r.kg_netto or 0.0

    return [
        BollaGroupedRow(
            articolo=key[0],
            descrizione=key[1],
            colore=key[2],
            partita=key[3],
            rocche=sums[key]["rocche"],
            kg_netto=sums[key]["kg_netto"],
        )
        for key in order
    ]


# ---------------------------------------------------------------------------
# Number parsing (European format: '.' thousands, ',' decimal)
# ---------------------------------------------------------------------------

def parse_euro_number(text: str) -> float | None:
    """
    Parse a European-format number such as '260,30' or '1.111,30' to a
    float. Returns None when the text isn't numeric.
    """
    text = clean_text(text)
    if not text:
        return None
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Spatial helpers (same line-grouping / column-band approach used by the
# main pdf_parser.py, kept self-contained here since the two document
# types have unrelated layouts)
# ---------------------------------------------------------------------------

def _group_into_lines(words: list[Word], y_tolerance: float = 3.0) -> list[list[Word]]:
    """Cluster words into visual lines by y-proximity, sorted left→right."""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: list[list[Word]] = []
    current: list[Word] = [sorted_words[0]]
    current_top = sorted_words[0]["top"]
    for w in sorted_words[1:]:
        if abs(w["top"] - current_top) <= y_tolerance:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda ww: ww["x0"]))
            current = [w]
            current_top = w["top"]
    lines.append(sorted(current, key=lambda ww: ww["x0"]))
    return lines


def _build_slots(
    header_words: list[Word],
    label_map: dict[str, str],
    page_width: float,
) -> list[tuple[float, float, str]]:
    """
    Build (x_start, x_end, field_name) column bands from header word
    positions. Each band runs from its own header's x0 to the next
    header's x0 (or the page edge for the last one).
    """
    entries = [
        (w["x0"], label_map[w["text"]])
        for w in header_words
        if w["text"] in label_map
    ]
    entries.sort(key=lambda e: e[0])
    slots: list[tuple[float, float, str]] = []
    for i, (x0, field) in enumerate(entries):
        x_end = entries[i + 1][0] if i + 1 < len(entries) else page_width
        slots.append((x0, x_end, field))
    return slots


def _assign_to_slots(
    words: list[Word], slots: list[tuple[float, float, str]]
) -> dict[str, str]:
    """Assign each word to the column band whose range contains its x0."""
    out: dict[str, list[str]] = {field: [] for _, _, field in slots}
    for w in words:
        for x0, x_end, field in slots:
            if x0 <= w["x0"] < x_end:
                out[field].append(w["text"])
                break
    return {k: clean_text(" ".join(v)) for k, v in out.items()}


def _build_right_edge_slots(
    header_words: list[Word], label_map: dict[str, str]
) -> list[tuple[float, str]]:
    """
    Build (x1, field_name) reference points from header word right edges.
    Used for right-aligned numeric columns (like the totals line), where a
    value's own x1 lines up with its header's x1 far more reliably than its
    x0 does — a short number's left edge drifts right and can land in the
    wrong x0-band, but its right edge stays anchored to the column.
    """
    return sorted(
        (
            (w["x1"], label_map[w["text"]])
            for w in header_words
            if w["text"] in label_map
        ),
        key=lambda e: e[0],
    )


def _assign_by_nearest_x1(
    words: list[Word], ref_points: list[tuple[float, str]]
) -> dict[str, str]:
    """Assign each word to the column whose header x1 is closest to its own x1."""
    out: dict[str, list[str]] = {field: [] for _, field in ref_points}
    for w in words:
        _, field = min(ref_points, key=lambda rp: abs(rp[0] - w["x1"]))
        out[field].append(w["text"])
    return {k: clean_text(" ".join(v)) for k, v in out.items()}


def _is_dashes(text: str) -> bool:
    return bool(text) and set(text) == {"-"}


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

class BollaParser:
    """Parses one Bolla PDF into item rows plus its totals summary."""

    def __init__(self, pdf_path: Path) -> None:
        self.pdf_path = Path(pdf_path)

    def parse(self) -> tuple[list[BollaRow], BollaTotals | None]:
        rows: list[BollaRow] = []
        totals: BollaTotals | None = None
        bolla_no = ""
        bolla_date = ""

        logger.info("Opening Bolla PDF: %s", self.pdf_path.name)

        with pdfplumber.open(self.pdf_path) as pdf:
            full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            m = RE_BOLLA_HEADER.search(full_text)
            if m:
                bolla_no = clean_text(m.group("no"))
                bolla_date = clean_text(m.group("date"))
            else:
                logger.warning(
                    "  Could not find 'Bolla Numero ... del ...' header in %s",
                    self.pdf_path.name,
                )

            for page in pdf.pages:
                words = page.extract_words(
                    x_tolerance=1.5, y_tolerance=3, keep_blank_chars=False
                )
                if not words:
                    continue

                lines = _group_into_lines(words)
                item_slots: list[tuple[float, float, str]] | None = None
                totals_slots: list[tuple[float, float, str]] | None = None
                totals_ref_points: list[tuple[float, str]] | None = None

                for line in lines:
                    texts = [w["text"] for w in line]
                    first_text = texts[0]

                    # Item table header (repeats on every page)
                    if "Pallet" in texts and "Scatola" in texts and "Articolo" in texts:
                        item_slots = _build_slots(line, ITEM_HEADER_LABELS, page.width)
                        continue

                    # Totals header ("Totale N.Colli N.Pallet ...")
                    if "N.Colli" in texts and "N.Pallet" in texts:
                        totals_ref_points = _build_right_edge_slots(line, TOTALS_HEADER_LABELS)
                        totals_slots = _build_slots(line, TOTALS_HEADER_LABELS, page.width)
                        continue

                    # Dashed separator line under the item header
                    if _is_dashes(first_text):
                        continue

                    # "Totale Generale ,00" footer line — nothing to extract
                    if first_text == "Totale":
                        continue

                    # Totals VALUES line: identified by position, not order —
                    # its leftmost word sits inside the totals column band,
                    # never in the Pallet/Scatola/Disposizione/Articolo zone.
                    # (Layouts sometimes interleave this line between item
                    # rows rather than placing it immediately after the
                    # totals header, so we can't rely on line order here.)
                    if totals_slots is not None:
                        band_start = totals_slots[0][0]
                        if min(w["x0"] for w in line) >= band_start - 5:
                            vals = _assign_by_nearest_x1(line, totals_ref_points)
                            totals = BollaTotals(
                                bolla_no=bolla_no,
                                bolla_date=bolla_date,
                                n_colli=parse_euro_number(vals.get("n_colli", "")),
                                n_pallet=parse_euro_number(vals.get("n_pallet", "")),
                                tara_tub=parse_euro_number(vals.get("tara_tub", "")),
                                tara_scat=parse_euro_number(vals.get("tara_scat", "")),
                                rocche=parse_euro_number(vals.get("rocche", "")),
                                netto=parse_euro_number(vals.get("netto", "")),
                                lordo=parse_euro_number(vals.get("lordo", "")),
                            )
                            continue

                    # Item row — starts with a plain integer Pallet number
                    if item_slots is not None and re.match(r"^\d+$", first_text):
                        vals = _assign_to_slots(line, item_slots)
                        rows.append(
                            BollaRow(
                                bolla_no=bolla_no,
                                bolla_date=bolla_date,
                                pallet=vals.get("pallet", ""),
                                scatola=vals.get("scatola", ""),
                                disposizione=vals.get("disposizione", ""),
                                articolo=vals.get("articolo", ""),
                                descrizione=vals.get("descrizione", ""),
                                colore=vals.get("colore", ""),
                                partita=vals.get("partita", ""),
                                rocche=parse_euro_number(vals.get("rocche", "")),
                                kg_netto=parse_euro_number(vals.get("kg_netto", "")),
                                kg_lordo=parse_euro_number(vals.get("kg_lordo", "")),
                                famiglia=vals.get("famiglia", ""),
                            )
                        )

        logger.info(
            "  %d item row(s) parsed%s — %s",
            len(rows),
            " (totals found)" if totals else " (no totals block found)",
            self.pdf_path.name,
        )
        return rows, totals
