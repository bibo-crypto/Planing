"""
pdf_parser.py — Coordinate-aware, rule-based parser for Delta Dyeing Purchase Order PDFs.

Architecture
============
Uses pdfplumber's extract_words() to obtain every word with its exact
(x0, top, x1, bottom) bounding box.  Column boundaries are derived from the
PDF's own header row — no fixed line numbers, no fixed coordinates, and no
fixed token counts appear anywhere in the logic.

Per-page pipeline
-----------------
1.  Extract all words with bounding boxes.
2.  Find the "Pos." anchor word (table header marker).
3.  Locate the first actual data row below the anchor (dynamic; not a fixed offset).
4.  Collect header-zone words: anchor.top − LOOK_ABOVE_POS  →  data_start − ε.
5.  Build column slots:
      • pos_slot: from 0 to pos_word.x1  (narrow — just the POS number area)
      • right-side slots (colour, qty_cones, qty_kg, price_usd, value_usd,
        ship_date, abbina): detected from header keywords; each slot's x_left =
        the header phrase's x0; x_right = next column header's x0.
      • article_area: implicit gap between pos_slot.x_right and colour_slot.x_left.
        Contains Article No (first word per row) and Article Description (rest).
6.  Scan left-side words only for POD No / Po Date metadata.
7.  Collect data words: [data_start, total_y).
8.  Group into visual lines by y-proximity.
9.  Walk lines:
        Machine line  → flush row, update Abbina
        POS number in pos_slot → flush row, start new row accumulator
        Otherwise     → continuation: extend article_description only
10. Flush final row.

Extending for new layouts
-------------------------
• New right-side column  → add entry to RIGHT_SCHEMA.
• Renamed header keyword → add the alias to the existing RIGHT_SCHEMA entry.
• New POD/Date label     → extend RE_POD or RE_DATE.
• Header more than LOOK_ABOVE_POS pt above "Pos."  → increase that constant.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple

import pdfplumber

from utils import clean_text, divide_pos, logger, parse_number, split_article_description


# ---------------------------------------------------------------------------
# Column schema — right-side columns only
# ---------------------------------------------------------------------------
# Article No. and Article Description are NOT listed here; they share the same
# x-region in the PDF header and are handled as an implicit "article_area"
# slot between the pos_slot and the colour_slot.
#
# Order in the list matters only for tie-breaking (first unmatched wins).
# Each tuple: (field_name, [keywords_to_match_in_header_phrase_lowercase])
# ---------------------------------------------------------------------------

RIGHT_SCHEMA: list[tuple[str, list[str]]] = [
    ("colour",    ["colour", "color"]),
    ("qty_cones", ["/cones", "cones"]),
    ("qty_kg",    ["/kg"]),
    ("price_usd", ["price in $", "price in", "price"]),
    ("value_usd", ["value in usd", "value in", "value"]),
    ("ship_date", ["ship. date", "ship date", "ship."]),
    ("abbina",    ["abbina"]),
    ("reference_to", ["reference to", "ref. to", "ref to"]),
]

# ---------------------------------------------------------------------------
# Tuning constants
# ---------------------------------------------------------------------------

# How far above the "Pos." word to start looking for header words (pt).
# Handles PDFs that place "Article No." slightly above the main header line.
LOOK_ABOVE_POS: float = 20.0

# Horizontal gap (pt) below which two words on the same line are merged into
# one phrase for keyword matching.
PHRASE_GAP_THRESHOLD: float = 10.0

# Words whose `top` values differ by ≤ this are placed on the same visual line.
LINE_Y_TOLERANCE: float = 4.0

# pdfplumber word-extraction tolerances.
WORD_X_TOL: float = 3.0
WORD_Y_TOL: float = 3.0

# Metadata scanning: only words with x0 < this value are considered for POD
# No / Po Date extraction (keeps right-side address text out of captures).
META_LEFT_ZONE_X: float = 250.0

# ---------------------------------------------------------------------------
# Compiled regexes
# ---------------------------------------------------------------------------

RE_POD = re.compile(r"POD\s*No[\s:\-]*(.+)", re.IGNORECASE)
RE_DATE = re.compile(
    r"(?:Po\s*Date|Purchase\s*Order\s*Date)[\s:\-]*(.+)", re.IGNORECASE
)
# Used to extract just the date value when the capture group contains extra text
RE_DATE_VALUE = re.compile(r"\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}")

RE_TABLE_END = re.compile(r"Total\s*:", re.IGNORECASE)
RE_POS_NUMBER = re.compile(r"^\d+$")           # e.g. 10, 20, 30 …
RE_MACHINE = re.compile(r"Machine\s+\S+\s+Cone", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------

@dataclass
class OrderRow:
    """One line-item from a Delta Dyeing Purchase Order PDF."""

    pod_no: str = ""
    po_date: str = ""
    pos: int | None = None
    article_no: str = ""
    # Looked up from the user-maintained Elvy mapping (see elvy_mapping.py)
    # by article_no, after parsing — left "" if no mapping is found.
    articolo_delta: str = ""
    # Looked up from the DFM colour reference (see dfm_lookup.py) by
    # article_no + colour, after parsing — left "" if no match is found.
    coloredfm: str = ""
    cldescr: str = ""
    article_description: str = ""
    yarn: str = ""
    nm: str = ""
    ne: str = ""
    dye_type: str = ""
    lot: str = ""
    colour: str = ""
    quantity_cones: float | None = None
    quantity_kg: float | None = None
    price_usd: float | None = None
    value_usd: float | None = None
    ship_date: str = ""
    abbina: str = ""
    # New PDF column: a free-text reference the customer attaches to the
    # order (e.g. 'Last Patch') -- shown as its own column and folded into
    # Commento after the year.
    reference_to: str = ""
    # "prima volta tint." when this exact Articolo Delta + COLOREDFM pair
    # has no prior entry at all in the DFM reference (see dfm_lookup.py's
    # is_first_time_dyeing) -- left "" otherwise.
    check_articolo: str = ""
    # Livello/Prezzo -- looked up from the Prezzi tab's Listini data by
    # (finished Articolo, COLOREDFM). None when there's no price match.
    livello: float | None = None
    prezzo: float | None = None


class ColumnSlot(NamedTuple):
    """Horizontal x-range belonging to one logical column."""

    name: str
    x_left: float    # inclusive
    x_right: float   # exclusive


# ---------------------------------------------------------------------------
# Word type alias
# ---------------------------------------------------------------------------

Word = dict   # pdfplumber word dict: text, x0, x1, top, bottom, …


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

class PDFParser:
    """
    Extracts :class:`OrderRow` objects from a Delta Dyeing PO PDF.

    Usage::

        rows = PDFParser(Path("order.pdf")).parse()
    """

    def __init__(self, pdf_path: Path) -> None:
        self.pdf_path = pdf_path

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def parse(self) -> list[OrderRow]:
        logger.info("Opening PDF: %s", self.pdf_path.name)
        try:
            return self._parse_pdf()
        except Exception as exc:
            logger.error("Failed to parse %s: %s", self.pdf_path.name, exc)
            raise RuntimeError(
                f"Could not parse {self.pdf_path.name}: {exc}"
            ) from exc

    # ------------------------------------------------------------------
    # Top-level pipeline
    # ------------------------------------------------------------------

    def _parse_pdf(self) -> list[OrderRow]:
        all_rows: list[OrderRow] = []
        pod_no = ""
        po_date = ""
        current_abbina = ""

        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                logger.debug("  Page %d", page_num)

                words: list[Word] = page.extract_words(
                    x_tolerance=WORD_X_TOL,
                    y_tolerance=WORD_Y_TOL,
                    keep_blank_chars=False,
                    use_text_flow=False,
                )
                if not words:
                    continue

                # ── 1. Metadata (POD No, Po Date) ──────────────────────────
                pod_no, po_date = self._scan_metadata(words, pod_no, po_date)

                # ── 2. Table anchor ────────────────────────────────────────
                pos_word = self._find_pos_anchor(words)
                if pos_word is None:
                    logger.debug("    No 'Pos.' anchor on page %d — skipping", page_num)
                    continue

                # ── 3. Dynamic data-start detection ────────────────────────
                data_start_top = self._find_data_start(words, pos_word)
                if data_start_top is None:
                    logger.warning(
                        "    Could not detect data start on page %d — skipping", page_num
                    )
                    continue

                logger.debug(
                    "    Pos. anchor y=%.1f  data starts y=%.1f",
                    pos_word["top"], data_start_top,
                )

                # ── 4. Header zone ─────────────────────────────────────────
                header_top = pos_word["top"] - LOOK_ABOVE_POS
                header_words = [
                    w for w in words
                    if header_top <= w["top"] < data_start_top
                ]

                # ── 5. Column slot detection ───────────────────────────────
                slots = self._build_column_slots(header_words, pos_word, page.width)
                if slots is None:
                    logger.warning("    Column detection failed on page %d — skipping", page_num)
                    continue
                _log_slots(slots)

                # ── 6. Table end ───────────────────────────────────────────
                total_y = self._find_total_y(words, data_start_top) or page.height

                # ── 7. Data words ──────────────────────────────────────────
                data_words = [
                    w for w in words
                    if data_start_top <= w["top"] < total_y
                ]

                # ── 8–10. Lines → rows ────────────────────────────────────
                lines = _group_into_lines(data_words, LINE_Y_TOLERANCE)
                page_rows, current_abbina, row_tops = self._lines_to_rows(
                    lines, slots, pod_no, po_date, current_abbina
                )

                # ── 11. Abbina from PDF annotations ─────────────────────────
                # "Machine X cones" labels are frequently added by hand as PDF
                # sticky-note / FreeText annotations (e.g. via Acrobat) rather
                # than as real text in the content stream. extract_words()
                # never sees these, so they must be read separately via
                # page.annots and matched to rows by vertical position.
                self._apply_annotation_abbinas(
                    page_rows, row_tops, page, data_start_top, total_y
                )

                all_rows.extend(page_rows)

        logger.info("  %d row(s) parsed — %s", len(all_rows), self.pdf_path.name)
        return all_rows

    # ------------------------------------------------------------------
    # Step 1 — metadata
    # ------------------------------------------------------------------

    def _scan_metadata(
        self, words: list[Word], pod_no: str, po_date: str
    ) -> tuple[str, str]:
        """
        Scan left-side words only (x0 < META_LEFT_ZONE_X) for POD No and
        Po Date.  Restricting to the left zone prevents right-side address
        text that shares the same visual line from corrupting the captures.
        """
        left_words = [w for w in words if w["x0"] < META_LEFT_ZONE_X]
        lines = _group_into_lines(left_words, LINE_Y_TOLERANCE)

        for line in lines:
            text = " ".join(w["text"] for w in line)

            if not pod_no:
                m = RE_POD.search(text)
                if m:
                    pod_no = clean_text(m.group(1))

            if not po_date:
                m = RE_DATE.search(text)
                if m:
                    captured = clean_text(m.group(1))
                    # If extra text follows the date, extract just the date token
                    dm = RE_DATE_VALUE.search(captured)
                    po_date = dm.group(0) if dm else captured

            if pod_no and po_date:
                break

        return pod_no, po_date

    # ------------------------------------------------------------------
    # Step 2 — anchor
    # ------------------------------------------------------------------

    @staticmethod
    def _find_pos_anchor(words: list[Word]) -> Word | None:
        """Return the 'Pos.' / 'Pos' header word (leftmost, topmost match)."""
        candidates = [
            w for w in words
            if w["text"].lower().rstrip(".") == "pos"
        ]
        if not candidates:
            return None
        # Pick the topmost (smallest top value); leftmost breaks ties
        return min(candidates, key=lambda w: (w["top"], w["x0"]))

    # ------------------------------------------------------------------
    # Step 3 — dynamic data-start detection
    # ------------------------------------------------------------------

    @staticmethod
    def _find_data_start(words: list[Word], pos_word: Word) -> float | None:
        """
        Return the top-y of the first data row by scanning for a bare integer
        (10, 20, 30 …) that appears in the POS column's x-range below the
        anchor word.

        Using the anchor's own x1 as the right boundary of the POS area
        ensures that article numbers or description tokens starting just to
        the right of the POS column are not mistaken for POS values.
        """
        pos_col_x_right = pos_word["x1"] + 5   # small tolerance
        below_y = pos_word["bottom"]

        candidates: list[float] = []
        for w in words:
            if w["top"] <= below_y:
                continue
            if w["x0"] > pos_col_x_right:
                continue
            if RE_POS_NUMBER.match(w["text"]):
                candidates.append(w["top"])

        return min(candidates) if candidates else None

    # ------------------------------------------------------------------
    # Step 5 — column slot detection
    # ------------------------------------------------------------------

    def _build_column_slots(
        self,
        header_words: list[Word],
        pos_word: Word,
        page_width: float,
    ) -> dict[str, ColumnSlot] | None:
        """
        Derive column x-boundaries from the header zone.

        pos_slot  : narrow — from 0 to pos_word.x1.
        right-side: each slot's x_left = the header phrase's own x0;
                    x_right = the next slot's x0 (or page_width for the last).
        article_area: implicit gap between pos_slot.x_right and
                      colour_slot.x_left.

        This approach is robust because:
        • It does NOT depend on midpoints between distant headers.
        • Article No. and Article Description (which overlap in x) are not
          treated as separate measurable columns.
        • Every right-side column is anchored to its own header text.
        """
        phrases = _build_phrases(header_words, PHRASE_GAP_THRESHOLD, LINE_Y_TOLERANCE)

        # Match phrases → right-side schema; record (phrase_x0, phrase_x1)
        matched: dict[str, tuple[float, float]] = {}
        for phrase_words in phrases:
            phrase_text = " ".join(w["text"] for w in phrase_words).lower().strip()
            px0 = phrase_words[0]["x0"]
            px1 = phrase_words[-1]["x1"]

            for field_name, keywords in RIGHT_SCHEMA:
                if field_name in matched:
                    continue
                for kw in keywords:
                    if kw in phrase_text:
                        matched[field_name] = (px0, px1)
                        break

        if "colour" not in matched:
            logger.warning("    'colour' column not detected — cannot build article area")
            return None

        # Sort right-side columns by x0 position
        sorted_right = sorted(matched.items(), key=lambda t: t[1][0])

        # Build slots: x_left = phrase x0; x_right = next phrase x0 (or page_width)
        slots: dict[str, ColumnSlot] = {}
        for i, (name, (px0, _px1)) in enumerate(sorted_right):
            if i < len(sorted_right) - 1:
                x_right = sorted_right[i + 1][1][0]
            else:
                x_right = page_width
            slots[name] = ColumnSlot(name, px0, x_right)

        # pos_slot: narrow strip ending at pos_word.x1
        pos_x_right = pos_word["x1"]
        slots["pos"] = ColumnSlot("pos", 0.0, pos_x_right)

        # article_area: gap between pos_slot and colour_slot
        colour_x_left = slots["colour"].x_left
        slots["article_area"] = ColumnSlot("article_area", pos_x_right, colour_x_left)

        return slots

    # ------------------------------------------------------------------
    # Table end
    # ------------------------------------------------------------------

    @staticmethod
    def _find_total_y(words: list[Word], below_y: float) -> float | None:
        """Return the top-y of the 'Total :' line, or None."""
        # Reconstruct lines below the header zone
        data_words = [w for w in words if w["top"] >= below_y]
        lines = _group_into_lines(data_words, LINE_Y_TOLERANCE)
        for line in lines:
            text = " ".join(w["text"] for w in line)
            if RE_TABLE_END.search(text):
                return line[0]["top"]
        return None

    # ------------------------------------------------------------------
    # Lines → OrderRow objects
    # ------------------------------------------------------------------

    def _lines_to_rows(
        self,
        lines: list[list[Word]],
        slots: dict[str, ColumnSlot],
        pod_no: str,
        po_date: str,
        current_abbina: str,
    ) -> tuple[list[OrderRow], str, list[float]]:
        """
        Walk visual lines and build :class:`OrderRow` objects.

        Rules
        -----
        - Machine line  -> flush current row; update current_abbina.
        - POS line      -> flush current row; open new accumulator; assign all
                          right-side fields from the main row line.
        - Continuation  -> extend article_description only (words in
                          article_area); ignore words in right-side slots.
                          This prevents "C.W 21" (which appears in the ship
                          date column position on continuation lines) from
                          overriding the already-captured ship date.

        Also returns ``row_tops``: the top-y of each row's POS line, in the
        same order as the returned rows. Used by ``_apply_annotation_abbinas``
        to match PDF annotations (which are not part of the text layer) back
        to the row they visually sit next to.
        """
        rows: list[OrderRow] = []
        row_tops: list[float] = []
        acc: dict[str, list[str]] = {}
        acc_top: float | None = None

        pos_slot = slots.get("pos")
        art_slot = slots.get("article_area")

        def flush() -> None:
            if acc:
                row = _assemble_row(acc, pod_no, po_date, current_abbina)
                if row is not None:
                    rows.append(row)
                    row_tops.append(acc_top if acc_top is not None else 0.0)
            acc.clear()

        colour_slot = slots.get("colour")

        def _is_colour_word(w: Word) -> bool:
            """
            A word belongs to the colour column if its right edge reaches or
            crosses the colour column's left boundary.  This captures colour-
            code tokens (e.g. "009.Red") that start slightly to the left of the
            header anchor but visually sit in the colour column area.
            """
            if colour_slot is None:
                return False
            return w["x1"] >= colour_slot.x_left and w["x0"] < colour_slot.x_right

        def _article_area_words(line_words: list[Word]) -> list[Word]:
            """Words that belong to the article area (not colour-overlapping)."""
            if art_slot is None:
                return []
            return sorted(
                [w for w in line_words
                 if _in_slot(w, art_slot) and not _is_colour_word(w)],
                key=lambda w: w["x0"],
            )

        for line in lines:
            line_text = " ".join(w["text"] for w in line)

            # ── Machine / Abbina ────────────────────────────────────────
            if RE_MACHINE.search(line_text):
                flush()
                current_abbina = line_text.strip()
                logger.debug("    Abbina → %s", current_abbina)
                continue

            # ── New data row? ───────────────────────────────────────────
            if pos_slot:
                pos_words_on_line = [
                    w for w in line if _in_slot(w, pos_slot)
                ]
            else:
                pos_words_on_line = []

            is_new_row = bool(
                pos_words_on_line
                and RE_POS_NUMBER.match(pos_words_on_line[0]["text"])
            )

            if is_new_row:
                flush()
                acc = defaultdict(list)
                acc_top = pos_words_on_line[0]["top"]

                # POS
                acc["pos"].append(pos_words_on_line[0]["text"])

                # Article area: first word = article_no, rest = description.
                # Words whose right edge overlaps the colour column are excluded
                # here and collected below as colour words instead.
                art_words = _article_area_words(line)
                if art_words:
                    acc["article_no"].append(art_words[0]["text"])
                    acc["article_description"].extend(
                        w["text"] for w in art_words[1:]
                    )

                # Right-side columns.
                # Colour uses right-edge overlap (_is_colour_word) so that
                # codes like "009.Red" that straddle the column boundary are
                # captured together with the colour name ("Water Color").
                for field_name in (
                    "colour", "qty_cones", "qty_kg",
                    "price_usd", "value_usd", "ship_date", "abbina",
                ):
                    slot = slots.get(field_name)
                    if not slot:
                        continue
                    for w in line:
                        if field_name == "colour":
                            if _is_colour_word(w):
                                acc[field_name].append(w["text"])
                        else:
                            if _in_slot(w, slot):
                                acc[field_name].append(w["text"])

            elif acc:
                # ── Continuation line ───────────────────────────────────
                # Only extend article_description; ignore right-side columns.
                # Same colour-overlap exclusion applies.
                art_words = _article_area_words(line)
                acc["article_description"].extend(
                    w["text"] for w in art_words
                )

        flush()
        return rows, current_abbina, row_tops

    # ------------------------------------------------------------------
    # Step 11 — Abbina from PDF annotations
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_annotation_abbinas(
        rows: list[OrderRow],
        row_tops: list[float],
        page: "pdfplumber.page.Page",
        data_start_top: float,
        total_y: float,
    ) -> None:
        """
        Read "Machine X cones" labels from PDF annotations (sticky notes /
        FreeText comments) and assign them to rows.

        In real-world Delta Dyeing POs, these labels are frequently added by
        hand in Acrobat as annotations rather than typed into the page's text
        layer, so ``extract_words()`` never sees them. ``page.annots``
        exposes their text (``contents``) and position (``top``) separately.

        Each annotation is matched to the row it visually sits next to (by
        y-position), then applied to every row in that row's consecutive
        same-colour group — mirroring how a single "Machine 32 cones" label
        placed once next to a pair of same-colour rows is meant to describe
        both of them.
        """
        if not rows or not row_tops:
            return

        annots = page.annots or []
        machine_annots: list[tuple[float, str]] = []
        for a in annots:
            contents = a.get("contents")
            top = a.get("top")
            if not contents or top is None:
                continue
            if not (data_start_top <= top < total_y):
                continue
            if RE_MACHINE.search(contents):
                machine_annots.append((top, clean_text(contents)))

        if not machine_annots:
            return

        groups = _group_consecutive_by_colour(rows)

        for annot_top, text in machine_annots:
            anchor_idx = None
            for i, row_top in enumerate(row_tops):
                next_top = row_tops[i + 1] if i + 1 < len(row_tops) else total_y
                if row_top <= annot_top < next_top:
                    anchor_idx = i
                    break
            if anchor_idx is None:
                # Annotation sits above every row's own top (rare) — fall
                # back to the nearest row at or above it.
                above = [i for i, t in enumerate(row_tops) if t <= annot_top]
                anchor_idx = above[-1] if above else None
            if anchor_idx is None:
                continue

            for group in groups:
                if anchor_idx in group:
                    for idx in group:
                        rows[idx].abbina = text
                    break


# ---------------------------------------------------------------------------
# Row assembly
# ---------------------------------------------------------------------------

def _assemble_row(
    acc: dict[str, list[str]],
    pod_no: str,
    po_date: str,
    abbina: str,
) -> OrderRow | None:
    """Convert a text accumulator to an :class:`OrderRow`. Returns None if invalid."""

    def get(name: str) -> str:
        return " ".join(acc.get(name, [])).strip()

    pos_raw = get("pos")
    if not RE_POS_NUMBER.match(pos_raw):
        logger.debug("    Skipped acc with invalid pos=%r", pos_raw)
        return None

    article_description = get("article_description")
    parts = split_article_description(article_description)

    return OrderRow(
        pod_no=pod_no,
        po_date=po_date,
        pos=divide_pos(pos_raw),
        article_no=get("article_no"),
        article_description=article_description,
        yarn=parts["yarn"],
        nm=parts["nm"],
        ne=parts["ne"],
        dye_type=parts["dye_type"],
        lot=parts["lot"],
        colour=get("colour"),
        quantity_cones=parse_number(get("qty_cones")),
        quantity_kg=parse_number(get("qty_kg")),
        price_usd=parse_number(get("price_usd")),
        value_usd=parse_number(get("value_usd")),
        ship_date=get("ship_date"),
        abbina=abbina,
        reference_to=get("reference_to"),
    )


# ---------------------------------------------------------------------------
# Spatial helpers
# ---------------------------------------------------------------------------

def _group_into_lines(words: list[Word], y_tolerance: float) -> list[list[Word]]:
    """
    Cluster words into visual lines by y-proximity.
    Each returned line is sorted left → right by x0.
    """
    if not words:
        return []

    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: list[list[Word]] = []
    current_line: list[Word] = [sorted_words[0]]
    current_top: float = sorted_words[0]["top"]

    for word in sorted_words[1:]:
        if abs(word["top"] - current_top) <= y_tolerance:
            current_line.append(word)
        else:
            lines.append(sorted(current_line, key=lambda w: w["x0"]))
            current_line = [word]
            current_top = word["top"]

    if current_line:
        lines.append(sorted(current_line, key=lambda w: w["x0"]))

    return lines


def _build_phrases(
    words: list[Word],
    gap_threshold: float,
    y_tolerance: float,
) -> list[list[Word]]:
    """
    Group words into visual lines then merge horizontally adjacent words
    (gap < gap_threshold) into phrases for keyword matching.
    """
    lines = _group_into_lines(words, y_tolerance)
    phrases: list[list[Word]] = []

    for line in lines:
        if not line:
            continue
        current: list[Word] = [line[0]]
        for word in line[1:]:
            gap = word["x0"] - current[-1]["x1"]
            if gap < gap_threshold:
                current.append(word)
            else:
                phrases.append(current)
                current = [word]
        phrases.append(current)

    return phrases


def _group_consecutive_by_colour(rows: list[OrderRow]) -> list[list[int]]:
    """
    Return groups of row indices, where each group is a run of consecutive
    rows sharing the same ``colour`` value (exact match). Mirrors the
    grouping rule used by :class:`~abbina_calculator.AbbinaCalculator`.
    """
    groups: list[list[int]] = []
    i = 0
    while i < len(rows):
        colour = rows[i].colour
        start = i
        while i < len(rows) and rows[i].colour == colour:
            i += 1
        groups.append(list(range(start, i)))
    return groups


def _in_slot(word: Word, slot: ColumnSlot) -> bool:
    """True when the word's x-centre falls within the slot's x-range."""
    x_centre = (word["x0"] + word["x1"]) / 2
    return slot.x_left <= x_centre < slot.x_right


def _log_slots(slots: dict[str, ColumnSlot]) -> None:
    for slot in sorted(slots.values(), key=lambda s: s.x_left):
        logger.debug(
            "    slot %-22s  x=[%6.1f, %6.1f)",
            slot.name, slot.x_left, slot.x_right,
        )
