"""
ordine_kamal.py
Builds "Ordine Kamal" rows from parsed Kamal order-letter rows
(kamal_parser.KamalOrderRow), matching each colour's Articolo (C170xxxx)
via its Titolo ("نمرة الغزل") against the DFM reference -- Kamal's PDFs
don't carry an Article No column the way Elvy's do.

Column structure mirrors OrdiniElvyRow exactly (same order, same fixed
defaults) with only CLIENTE changed to Kamal's code -- COLORE holds the
COLOREDFM code, same convention as Elvy's own "COLORE" column, not a text
colour name.

COMMENTO format (as specified): "PG-X-ORD.nr.<year>(LETTER <reply no>)"
-- X gets replaced with the matched raw yarn Partita the same way as
Elvy orders, via ordini_elvy.match_raw_yarn().
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta

from abbina_calculator import MACHINE_CODES, _smallest_fitting_machine
from dfm_lookup import find_articolo_by_titolo
from kamal_parser import KamalOrderRow
from ordini_elvy import (
    CONSEGNA_LEAD_DAYS,
    LAVORANTE_CODE,
    LAVORANTE_SUCCESSIVO_CODE,
    SIGLA_DISPO_CODE,
    STOCK_MAG_CODE,
)
from utils import clean_text, logger

CLIENTE_CODE = 3017
KAMAL_ARTICLE_PREFIX = "C170"


@dataclass
class OrdineKamalRow:
    """One derived row for the "Ordine Kamal" sheet -- same column set as
    OrdiniElvyRow, same order, same fixed-field defaults."""
    cliente: int = CLIENTE_CODE
    articolo_delta: str = ""
    coloredfm: str = ""            # "COLORE" column -- the COLOREDFM code, same as Elvy
    peso_kg: float = 0.0           # "Q.TA" column -- weight in kg for Kamal (agreed)
    data_consegna: datetime | None = None
    commento: str = ""
    abbina: str = ""
    lavorante: int = LAVORANTE_CODE
    lavorante_successivo: int = LAVORANTE_SUCCESSIVO_CODE
    data_riconsegna: datetime | None = None
    stock_mag: int = STOCK_MAG_CODE
    sigla_dispo: str = SIGLA_DISPO_CODE
    bagno_proposto: str = ""
    macchina: int | None = None
    sender_msg_no: str = ""        # "رقم رسالة الخام" -- used to match Lotto in LOTTI
    colore_name: str = ""          # "اللون" -- Kamal's own colour label, used as the
                                    # GRUPPO MACCHINA grouping key when COLOREDFM is blank

    @property
    def quantity_cones(self) -> float:
        """Compatibility alias for the shared ERP writer/matching helpers.

        Kamal requests are measured in kilograms, but the common Ordini
        helpers use ``quantity_cones`` as their quantity field.  Exposing the
        alias keeps the shared header mapping intact without changing the
        Kamal workbook's public ``peso_kg`` field.
        """
        return self.peso_kg


def _machine_group_key(row: OrdineKamalRow) -> tuple[str, str] | None:
    """
    Grouping key used to sum weight and pick a machine for GRUPPO MACCHINA.

    COLOREDFM is the primary key. When a row has no COLOREDFM (blank),
    fall back to grouping by Colore ("اللون", Kamal's own colour label)
    instead of skipping the row -- rows sharing the same Colore still get
    their Peso (kg) summed together and rounded to one machine; a Colore
    that appears on its own is rounded using just its own weight. Rows
    with neither COLOREDFM nor Colore have no grouping key and are left
    unassigned.
    """
    color = clean_text(row.coloredfm)
    if color:
        return ("coloredfm", color)
    colore = clean_text(row.colore_name)
    if colore:
        return ("colore", colore)
    return None


def assign_ordine_kamal_machines(rows: list[OrdineKamalRow]) -> None:
    """Assign GRUPPO MACCHINA consistently for workbook and ERP exports."""
    totals: dict[tuple[str, str], float] = {}
    for row in rows:
        key = _machine_group_key(row)
        if key is None:
            continue
        totals[key] = totals.get(key, 0.0) + (row.peso_kg or 0.0)

    for row in rows:
        key = _machine_group_key(row)
        if key is None:
            continue
        capacity = _smallest_fitting_machine(totals[key])
        row.macchina = MACHINE_CODES.get(capacity)
        row.abbina = f"Machine {capacity} Cones"


def build_ordine_kamal_rows(rows: list[KamalOrderRow], dfm_c170_entries: list[dict]) -> list[OrdineKamalRow]:
    """
    dfm_c170_entries: dfm_lookup.load_dfm_entries_by_prefix("C170") --
    passed in rather than loaded here so the caller only loads it once per
    conversion run.
    """
    today = datetime.now()
    data_consegna_dt = today + timedelta(days=CONSEGNA_LEAD_DAYS)
    data_riconsegna_dt = data_consegna_dt - timedelta(days=1)

    out: list[OrdineKamalRow] = []
    unmatched = 0

    for r in rows:
        articolo = find_articolo_by_titolo(dfm_c170_entries, r.titolo) if r.titolo else ""
        if r.titolo and not articolo:
            unmatched += 1

        year = _extract_year(r.order_date)
        commento = f"PG-X-ORD.nr.{year}(LETTER {clean_text(r.reply_no)})"

        out.append(OrdineKamalRow(
            articolo_delta=articolo,
            coloredfm=r.coloredfm,
            peso_kg=r.peso_kg,
            data_consegna=data_consegna_dt,
            commento=commento,
            data_riconsegna=data_riconsegna_dt,
            sender_msg_no=clean_text(r.sender_msg_no),
            colore_name=clean_text(r.colore_name),
        ))

    if unmatched:
        logger.warning("Ordine Kamal: %d row(s) had a Titolo but no matching C170 Articolo in DFM", unmatched)

    return out


def match_by_lotto(rows: list[OrdineKamalRow], lotti_summary) -> int:
    """
    Fills in the Partita number in place of "X" in "PG-X-..." by matching
    each row's sender_msg_no (Kamal's own "رقم رسالة الخام") against the
    Lotto column in the LOTTI reference (lotti_logic.summarize_by_partita()
    output: columns partita, lotto) -- this is Kamal's own deterministic
    reference for which raw yarn batch to use, so it takes priority over
    quantity-based matching (ordini_elvy.match_raw_yarn).
    Returns how many rows were matched.
    """
    if lotti_summary is None or lotti_summary.empty:
        return 0

    lotto_to_partita: dict[str, str] = {}
    for _, r in lotti_summary.iterrows():
        lotto_to_partita.setdefault(str(r["lotto"]).strip(), str(r["partita"]).strip())

    matched = 0
    for row in rows:
        if "PG-X" not in (row.commento or ""):
            continue
        msg_no = clean_text(row.sender_msg_no)
        if not msg_no:
            continue
        partita = lotto_to_partita.get(msg_no)
        if partita:
            row.commento = row.commento.replace("PG-X", f"PG-{partita}")
            matched += 1

    return matched


def _extract_year(order_date: str) -> str:
    for part in clean_text(order_date).replace("-", "/").split("/"):
        if len(part) == 4 and part.isdigit():
            return part
    return str(datetime.now().year)
