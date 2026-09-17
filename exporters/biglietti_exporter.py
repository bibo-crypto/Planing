"""Extract order data into the Italian dyeing-ticket workbook and Word form.

The source workbook is the ERP export (``Sheet1``) plus the manually completed
``Doispo-Bagno`` and optional raw-yarn sheet.  The code deliberately matches
headers by normalized text and accepts either one or two header rows.

Reference logic: the machine table, Polmoni rule, Commento split, KG/Densita
formula and Color Tube/VMM22 lookup below were reverse-engineered from the
Power Query ("Densita' Quer", "Magazino", "codes", "PESO ROCCHE") embedded in
a real Delta export -- see the extraction notes shared with the user for the
exact M code these mirror. Anything derived from a file the user has not
uploaded yet (Densita' Query workbook, Magazino Color Tube workbook) degrades
gracefully to a blank cell rather than raising.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from calculate.constants import MACHINE_CAPACITIES


def _clean(v: Any) -> str:
    return " ".join(str(v or "").replace("\xa0", " ").split())


def _key(v: Any) -> str:
    # ``\w`` keeps Arabic headers (e.g. وزن and تحضير خام) as well as
    # the synthetic positional keys used for blank ERP headers.
    return re.sub(r"[^\w]+", " ", _clean(v).lower(), flags=re.UNICODE).strip()


def _number(v: Any) -> float | int | None:
    s = _clean(v).replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        n = float(s)
        return int(n) if n.is_integer() else n
    except ValueError:
        return None


def _order_number(v: Any) -> str:
    s = _clean(v)
    if "/" in s:
        s = s.rsplit("/", 1)[-1]
    return s.lstrip("0") or "0"


def _first_match(text: str, pattern: str) -> str:
    m = re.search(pattern, text, re.I)
    return m.group(1) if m else ""


# ---------------------------------------------------------------------------
# M/C (machine) lookup -- Rocche total per Bagno group -> physical machine
# number. Values taken from the reference Power Query's machine table.
# ---------------------------------------------------------------------------
# Mapping remains exporter-specific; the capacity list is centralized.
MACHINE_SIZE_TABLE: dict[int, int] = {
    capacity: machine for capacity, machine in zip(
        MACHINE_CAPACITIES, (11, 12, 9, 10, 7, 8, 5, 6, 3)
    )
}
# Seen only on MED orders in the reference query (7 Rocche -> machine 11).
# Kept separate since it overlaps oddly with the 6 -> 11 rule; flagged for
# the user to confirm once real MED data is available.
MACHINE_SIZE_TABLE_MED_EXTRA: dict[int, int] = {7: 11}

# A nearest-size match is only accepted within this tolerance (fraction of
# the nearest table size); anything further is left blank for a human to
# check, per "لو غير مطابقه بيقربها لأقرب ماكينه ولو بعيد سيبه فاضي".
_MACHINE_MATCH_TOLERANCE = 0.15


def _machine_for_count(count: float | int | None, extra: dict[int, int] | None = None) -> str:
    """Return a machine number; non-positive or missing count means blank."""
    if count is None or count <= 0:
        return ""
    table = dict(MACHINE_SIZE_TABLE)
    if extra:
        table.update(extra)
    if count in table:
        return str(table[count])
    nearest = min(table, key=lambda k: abs(k - count))
    if abs(nearest - count) <= nearest * _MACHINE_MATCH_TOLERANCE:
        return str(table[nearest])
    return ""


def _polmoni_segment(additional: str) -> str:
    """MED only: the 5th dash-separated segment of the additional
    description, kept as-is (e.g. '4 POLMONI') when it mentions Polmoni,
    matching the reference query's own POLMON output column -- this is
    text, not a number, and is what actually gets exported."""
    parts = [p.strip() for p in _clean(additional).split("-")]
    seg = parts[4] if len(parts) > 4 else ""
    return seg if "POLMON" in seg.upper() else ""


def _polmoni_extra(additional: str) -> int:
    """MED only: 'N POLMONI' adds N*4 Rocche worth of machine capacity
    (1 polmone stands in for 4 coni) -- used only for the M/C calculation,
    never exported directly."""
    seg = _polmoni_segment(additional)
    m = re.search(r"(\d+)\s*POLMON", seg, re.I) if seg else None
    return int(m.group(1)) * 4 if m else 0


def _commento_for(customer: str, additional: str) -> str:
    """Per-customer Commento, mirroring the reference query's dash-split
    on ALL '-' occurrences:
    ELVY keeps everything after the 2nd '-' (e.g. 'PG-157474-PO-1466-2026'
    -> 'PO-1466-2026'); MED keeps the 6th segment (.6), which is usually
    absent -> blank, confirmed against real MED data."""
    parts = [p.strip() for p in _clean(additional).split("-")]
    if customer == "ELVY":
        return "-".join(parts[2:]).strip() if len(parts) > 2 else ""
    return parts[5].strip() if len(parts) > 5 else ""


def _segment(additional: str, index: int) -> str:
    """1-based dash-segment of the additional description, 'X' treated as
    unassigned/blank (matches the reference query's own X->null rule)."""
    parts = [p.strip() for p in _clean(additional).split("-")]
    seg = parts[index - 1].strip() if len(parts) >= index else ""
    return "" if seg.upper() == "X" else seg


def _partita_med(additional: str) -> str:
    """Descrizione aggiuntiva ordine.4 -> 'Partita MED' in the reference
    query. Not a PO number for MED orders (that pattern is ELVY-only)."""
    return _segment(additional, 4)


def _cliente_med(additional: str) -> str:
    """Descrizione aggiuntiva ordine.5 -> 'Cliente MED'. The raw segment
    (e.g. 'ORDINE CAM') is expanded through _CLIENTE_MED_EXPANSIONS when
    recognised -- 'ORDINE CAM' -> 'ORDINE CAMPIONARIO' was confirmed
    against 3 real pre-existing tickets in the file the user sent, but
    this is the only code seen so far; ask the user for the rest of the
    mapping if other order-type abbreviations show up blank/unexpanded."""
    seg = _segment(additional, 5)
    return _CLIENTE_MED_EXPANSIONS.get(seg.upper(), seg)


_CLIENTE_MED_EXPANSIONS: dict[str, str] = {
    "ORDINE CAM": "ORDINE CAMPIONARIO",
}


def _po_token(additional: str) -> str:
    """Everything from the word 'PO' onward, e.g. 'PO-1466-2026'.
    ELVY-only naming convention -- MED orders use PM/dispo, not PO."""
    m = re.search(r"\bPO\b[-\s]*[0-9][0-9\-]*", additional, re.I)
    return m.group(0).strip() if m else ""


def _dispo_number(dispo: str) -> str:
    """Numeric part of a Dispo value like 'D-00505450-001' -> '505450'."""
    m = re.search(r"D-0*([0-9]+)", _clean(dispo), re.I)
    if m:
        return m.group(1)
    return _clean(dispo).lstrip("0") or _clean(dispo)


def _dispo_suffix(dispo: str) -> str:
    """The per-row sequence suffix of a Dispo value, e.g.
    'D-00505449-001' -> '1' -- this is what Sheet1's own 'Riga' column
    matches against (confirmed against real data: Riga=1,2,3,4 lines up
    with Dispo-Bagno rows -001,-002,-003,-004 in a different physical
    order than Sheet1 itself). Also accepts a bare number so it can be
    used directly on a Riga value too."""
    s = _clean(dispo)
    m = re.search(r"-0*([0-9]+)\s*$", s)
    if m:
        return str(int(m.group(1)))
    n = _number(s)
    return str(int(n)) if n is not None else ""


def build_output_stem(records: list["OrderRecord"], customer: str) -> str:
    """The identifier used both as the output filename and, for ELVY, as
    the ticket's own Commento -- e.g. 'PO-1466-2026' or 'MED-D-505450-2026'.
    """
    if customer == "ELVY":
        for r in records:
            token = _po_token(r.additional_raw)
            if token:
                return token
        return "ELVY"
    dispo_num = _dispo_number(records[0].dispo) if records else ""
    year = str(datetime.now().year)
    return f"MED-D-{dispo_num}-{year}" if dispo_num else f"MED-{year}"


# ---------------------------------------------------------------------------
# Titolo lookup -- shared with the Situazione tab's "Articoli" upload
# (Articolo Filato -> TITOLO), persisted in situazione_db so uploading the
# file from either tab makes it available to both.
# ---------------------------------------------------------------------------

def load_articoli_titolo_map() -> dict[str, str]:
    try:
        import utility.situazione_db as situazione_db
        return situazione_db.load_codes()  # {articolo_filato: titolo}
    except Exception:
        return {}


def load_articoli_marca_map(path: Path) -> tuple[dict[str, str], list[str]]:
    """Articolo Filato -> Marca (the fuller descriptive text, e.g.
    '100/2 - COTTON 100%') -- confirmed against real data as the actual
    source for the ticket's Titolo, NOT the shorter 'TITOLO' column that
    Situazione's own Articoli upload uses."""
    errors: list[str] = []
    out: dict[str, str] = {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        target = wb["Sheet2"] if "Sheet2" in wb.sheetnames else wb.active
        for row in _read_sheet_rows(target):
            art = _clean(_get(row, "Articolo Filato")).upper()
            marca = _clean(_get(row, "Marca"))
            if art and marca:
                out.setdefault(art, marca)
    finally:
        wb.close()
    if not out:
        errors.append("Nessuna riga Articolo Filato/Marca trovata nel file Articoli.")
    return out, errors


_ARTICOLI_MARCA_CACHE: tuple[str, int, int, dict[str, str]] | None = None


def load_articoli_marca_lookup() -> dict[str, str]:
    """Convenience: re-read whatever Articoli.xlsx was last uploaded via
    the Biglietti tab's own Articoli button (path cached in
    articoli_cache.py), returning {} if none has been uploaded yet."""
    try:
        import utility.articoli_cache as articoli_cache
    except Exception:
        return {}
    cache = articoli_cache.load_articoli_cache()
    path = cache.get("source_path")
    if not path or not Path(path).is_file():
        return {}
    source = Path(path)
    stat = source.stat()
    cache_key = (str(source), stat.st_mtime_ns, stat.st_size)
    global _ARTICOLI_MARCA_CACHE
    if _ARTICOLI_MARCA_CACHE and _ARTICOLI_MARCA_CACHE[:3] == cache_key:
        return _ARTICOLI_MARCA_CACHE[3]
    marca_map, _errors = load_articoli_marca_map(source)
    _ARTICOLI_MARCA_CACHE = (*cache_key, marca_map)
    return marca_map


def _titolo_lookup(articolo: str, codes_map: dict[str, str]) -> str:
    art = _clean(articolo).upper()
    if not art or not codes_map:
        return ""
    if art in codes_map:
        return codes_map[art]
    if art[:1] in ("C", "G"):
        swapped = ("G" if art[:1] == "C" else "C") + art[1:]
        if swapped in codes_map:
            return codes_map[swapped]
    return ""


# ---------------------------------------------------------------------------
# Prezzo lookup -- reuses the existing Listini/Prezzi infrastructure
# (prezzi_cache remembers the last uploaded file; no separate button needed
# here since the Prezzi tab already provides one).
# ---------------------------------------------------------------------------

_PREZZO_LOOKUP_CACHE: tuple[str, int, int, dict[tuple, tuple], str] | None = None


def load_prezzo_lookup() -> tuple[dict[tuple, tuple], str]:
    """Returns (lookup, source_file_name). Empty lookup + '' if nothing has
    been uploaded to the Prezzi tab yet."""
    global _PREZZO_LOOKUP_CACHE
    try:
        import utility.prezzi_cache as prezzi_cache
        import calculate.prezzi as prezzi_logic
    except Exception:
        return {}, ""
    cache = prezzi_cache.load_prezzi_cache()
    path = cache.get("source_path")
    if not path or not Path(path).is_file():
        return {}, ""
    source = Path(path)
    stat = source.stat()
    cache_key = (str(source), stat.st_mtime_ns, stat.st_size)
    if _PREZZO_LOOKUP_CACHE and _PREZZO_LOOKUP_CACHE[:3] == cache_key:
        return _PREZZO_LOOKUP_CACHE[3], _PREZZO_LOOKUP_CACHE[4]
    df, _errors = prezzi_logic.load_prezzi(path)
    if df is None or df.empty:
        return {}, ""
    lookup = prezzi_logic.build_price_lookup(df)
    source_file = cache.get("source_file", "")
    _PREZZO_LOOKUP_CACHE = (*cache_key, lookup, source_file)
    return lookup, source_file


# ---------------------------------------------------------------------------
# Delivery Date (ELVY only) -- how many working days out an order is quoted,
# based on whether the raw batch/yarn is already assigned (Partita GG) and
# whether it's already priced (Prezzo), plus which machine tier it's on.
# Working days skip Friday (pushed to Saturday if a computed date lands on
# one) -- matches the 'AddDaysSkipFriday' Power Query logic exactly.
# ---------------------------------------------------------------------------
_DELIVERY_MACHINES_LONG = {7, 8, 9, 10, 11, 12}
_DELIVERY_MACHINES_MED = {3, 4, 5, 6}


def _add_days_skip_friday(start_date, days_to_add: int):
    target = start_date + timedelta(days=days_to_add)
    if target.weekday() == 4:  # Friday
        target += timedelta(days=1)
    return target


def _compute_delivery_date(raw_batch: Any, prezzo: Any, machine: Any, today=None) -> Any:
    if not _clean(raw_batch):
        return "Bending for yarn"
    today = today or datetime.now().date()
    try:
        mc = int(_number(machine))
    except (TypeError, ValueError):
        mc = None
    has_prezzo = prezzo not in (None, "")
    if not has_prezzo:
        if mc in _DELIVERY_MACHINES_LONG:
            days = 24
        elif mc in _DELIVERY_MACHINES_MED:
            days = 16
        else:
            days = 8
    else:
        if mc in _DELIVERY_MACHINES_LONG:
            days = 16
        elif mc in _DELIVERY_MACHINES_MED:
            days = 8
        else:
            days = 0
    return _add_days_skip_friday(today, days)


def compute_delivery_date(records: list["OrderRecord"]) -> None:
    today = datetime.now().date()
    for r in records:
        r.delivery_date = _compute_delivery_date(r.raw_batch, r.prezzo, r.machine, today)


def _prezzo_for(articolo: str, codice: str, lookup: dict[tuple, tuple]) -> Any:
    if not lookup:
        return ""
    for key in ((articolo, codice), (_clean(articolo), _clean(codice))):
        if key in lookup:
            return lookup[key][1] or ""
    return ""


# Public alias -- other tabs (Situazione) reuse this same matching logic
# for their own Prezzo column instead of re-implementing it.
prezzo_for = _prezzo_for


# Machines (by their Rocche-based M/C total, e.g. Situazione's own 'mc'
# field or Ordine MED's) that get a $2 surcharge on top of the Listini
# price. Shared here since both Ordine MED's 'PREZZO + 2$' column and
# Situazione's own Prezzo column apply the exact same rule.
# The same machines may appear as capacity (24/32/56) or as the operator's
# machine number (12/9/10) in Situazione and the editing dialogs.
PREZZO_SURCHARGE_MACHINES = {9, 10, 12, 24, 32, 56}


def apply_machine_surcharge(price: Any, machine: Any) -> Any:
    """price + 2 when machine is 24/32/56 Rocche, otherwise price
    unchanged. machine may be an int, numeric string, or None/blank --
    anything that doesn't cleanly parse just skips the surcharge rather
    than raising."""
    if not isinstance(price, (int, float)):
        return price
    try:
        mc = int(_number(machine))
    except (TypeError, ValueError):
        return price
    return round(price + 2, 2) if mc in PREZZO_SURCHARGE_MACHINES else price


# ---------------------------------------------------------------------------
# Densita' Query workbook -- KG (PESO ROCCHE query) + Densita`(360-390),
# both keyed by Partita (raw_batch / "Partita GG").
# ---------------------------------------------------------------------------

def load_densita_query(path: Path) -> tuple[dict[int, dict[str, Any]], list[str]]:
    errors: list[str] = []
    out: dict[int, dict[str, Any]] = {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if "Entry" in wb.sheetnames:
            for row in _read_sheet_rows(wb["Entry"]):
                partita = _number(_get(row, "PARTITA", "Partita"))
                peso_lord = _number(_get(row, "Peso Lord", "Peso Lordo"))
                if partita is None:
                    continue
                peso_net = (peso_lord - 45) / 1000 if peso_lord is not None else None
                out.setdefault(int(partita), {})["peso_net"] = peso_net
                color_tube = _clean(_get(row, "Colore Tube", "Color Tube"))
                if color_tube:
                    out[int(partita)]["color_tube"] = color_tube
        else:
            errors.append('Nel file Densita\' Query.xlsx manca il foglio "Entry" (serve per KG).')
        if "Densita" in wb.sheetnames:
            for row in _read_sheet_rows(wb["Densita"]):
                partita = _number(_get(row, "PARTITA", "Partita"))
                if partita is None:
                    continue
                out.setdefault(int(partita), {})["densita"] = _get(row, "Densita` (360-390)", "Densita")
        else:
            errors.append('Nel file Densita\' Query.xlsx manca il foglio "Densita" (serve per Densita`(360-390)).')
    finally:
        wb.close()
    if not out and not errors:
        errors.append("Nessuna riga valida trovata nel file Densita' Query.")
    return out, errors


# ---------------------------------------------------------------------------
# Magazino (Color Tube) workbook -- Color Tube + kg-per-cone by Partita,
# used for the Biglietti "Color Tube" and "VMM22" columns. This is a
# different file from the raw-yarn Magazino used elsewhere in the app.
# ---------------------------------------------------------------------------

def load_vmm22_ratio_from_magazino(path: Path) -> tuple[dict[int, float], list[str]]:
    """{PARTITA: kg-per-cone} for VMM22, from the same raw Magazino Filato
    export already uploaded elsewhere in the app (Magazino Filato /
    Ordine Kamal tabs, cached path in magazino_cache.py) -- no separate
    upload needed. Per Partita: sum ESISTENZA (kg) and COLLI (cones)
    across MAGAZZINO 900910, 900160 and 900923, ratio = kg per cone.
    VMM22 for an order line = that ratio * the line's own Rocche."""
    errors: list[str] = []
    totals: dict[int, list[float]] = {}  # partita -> [esistenza_sum, colli_sum]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for row in _read_sheet_rows(wb.active):
            magazzino = _number(_get(row, "MAGAZZINO"))
            if magazzino not in (900910, 900160, 900923):
                continue
            partita = _number(_get(row, "PARTITA"))
            if partita is None:
                continue
            esistenza = _number(_get(row, "ESISTENZA")) or 0
            colli = _number(_get(row, "COLLI")) or 0
            bucket = totals.setdefault(int(partita), [0.0, 0.0])
            bucket[0] += esistenza
            bucket[1] += colli
    finally:
        wb.close()
    out = {p: (e / c) for p, (e, c) in totals.items() if c}
    if not out:
        errors.append("Nessuna riga valida trovata nel file Magazino Filato per VMM22.")
    return out, errors


@dataclass
class OrderRecord:
    customer_code: str
    customer_name: str
    article: str
    description: str
    additional_raw: str
    color_code: str
    color_name: str
    order_no: str
    order_row: str
    colored_batch: str
    raw_batch: str
    quantity_cones: float | int | None
    raw_weight: float | int | None
    dispo: str
    bagno: str
    delivery: Any = None
    machine: str = ""
    title: str = ""
    formato: str = "7777"
    commento: str = ""
    partita_med: str = ""
    cliente_med: str = ""
    kg: float | int | None = None
    densita: Any = ""
    color_tube: str = ""
    vmm22: float | int | None = None
    prezzo: Any = ""
    delivery_date: Any = ""


def _read_sheet_rows(ws) -> list[dict[str, Any]]:
    """Read a sheet using its most likely header row, even after export shifts."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    known_headers = {
        "cliente", "clienti", "articolo", "codice", "colore", "ordine", "riga",
        "consegna", "ordinata", "assegnata", "descrizioneaggiuntivaordine",
        "dispo", "bagno", "partita", "partitagg", "titolo", "mc", "rocche",
        "qta", "q ta", "dataconsegna", "data consegna", "commento",
        "disposizione", "numbagno",
    }
    candidates = []
    for idx in range(min(12, len(rows))):
        if idx < len(rows):
            headers = [_key(x) or f"__col{i + 1}" for i, x in enumerate(rows[idx])]
            known_score = sum(header in known_headers for header in headers)
            nonempty_score = sum(bool(x) for x in headers)
            candidates.append((known_score, nonempty_score, idx, headers))
    best_known = max((item[0] for item in candidates), default=0)
    if best_known >= 2:
        _, _, header_idx, headers = max(
            candidates, key=lambda item: (item[0], item[1], -item[2])
        )
    else:
        _, _, header_idx, headers = max(
            candidates, key=lambda item: (item[1], -item[2]), default=(0, 0, 0, [])
        )
    out = []
    for values in rows[header_idx + 1:]:
        if not any(_clean(x) for x in values):
            continue
        # Many ERP exports repeat the header row twice.  Do not turn that
        # repeated header into a real data record (especially important for
        # the first Dispo-Bagno row).
        comparable = [(i, _key(values[i]), headers[i]) for i in range(min(len(values), len(headers))) if _clean(values[i]) and headers[i].startswith("__") is False]
        if comparable and sum(a == b for _, a, b in comparable) >= max(2, len(comparable) // 2):
            continue
        rec = {}
        for i, value in enumerate(values):
            if i < len(headers):
                # Keep the first duplicate header; this avoids row-2 blank
                # labels replacing the actual row-1 label.
                rec.setdefault(headers[i], value)
                rec.setdefault(f"__col{i + 1}", value)
        out.append(rec)
    return out


def _get(row: dict[str, Any], *names: str) -> Any:
    aliases = {
        "articolo": ("codice articolo", "articolo codice"),
        "colore": ("codice colore", "colore codice", "descrizione colore"),
        "ordinata": ("qta", "q ta", "q.t.a", "quantita", "quantità", "rocche"),
        "riga": ("numero riga", "nr riga", "n riga"),
        "consegna": ("data consegna", "delivery date"),
        "descrizione aggiuntiva ordine": ("descrizione ordine", "commento ordine", "commento"),
        "dispo": ("disposizione", "sigla disposizione"),
        "bagno": ("n bagno", "numero bagno"),
        "partita": ("partita col", "partita colore"),
        "partita.gg": ("partita gg", "partita grezzo", "partita filato"),
    }
    expanded = list(names)
    for name in names:
        expanded.extend(aliases.get(_key(name), ()))
    for name in expanded:
        if _key(name) in row and _clean(row[_key(name)]):
            return row[_key(name)]
    return ""


def detect_order_format(data_path: Path) -> str:
    """'ELVY_MED' or 'EL_KAMAL', by peeking at Sheet1's own header row --
    lets a single Data Ordine picker serve every customer instead of a
    separate one per source shape. EL KAMAL's Sheet1 already carries
    computed columns (CODICE/TITOLO/M-C/Clienti) that ELVY/MED's raw
    export never has; ELVY/MED's Sheet1 has 'Descrizione aggiuntiva
    ordine' and a numeric 'Cliente' code that EL KAMAL's never has.
    Defaults to 'ELVY_MED' if neither signature is conclusive."""
    wb = openpyxl.load_workbook(data_path, read_only=True)
    try:
        if "Sheet1" not in wb.sheetnames:
            return "ELVY_MED"
        rows = list(wb["Sheet1"].iter_rows(values_only=True, max_row=1))
    finally:
        wb.close()
    if not rows:
        return "ELVY_MED"
    headers = {_key(v) for v in rows[0] if v}
    el_kamal_signature = {_key("CODICE"), _key("Clienti"), _key("M/C")}
    elvy_med_signature = {_key("Descrizione aggiuntiva ordine"), _key("Cliente")}
    if el_kamal_signature.issubset(headers):
        return "EL_KAMAL"
    if elvy_med_signature.issubset(headers):
        return "ELVY_MED"
    return "EL_KAMAL" if "Doispo-Bagno" not in wb.sheetnames else "ELVY_MED"


def load_dispo_bagno_rows(path: Path) -> list[dict[str, Any]]:
    """Reads a Dispo-Bagno source regardless of shape: a .csv (EL KAMAL's
    own export), or an .xlsx with the data on its first sheet."""
    if path.suffix.lower() == ".csv":
        return _read_dispo_csv(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = next(
            (ws for ws in wb.worksheets
             if "dispo" in _key(ws.title) and "bagno" in _key(ws.title)),
            wb.active,
        )
        return _read_sheet_rows(sheet)
    finally:
        wb.close()


def load_order(input_path: Path, dispo_path: Path | None = None) -> tuple[list[OrderRecord], list[dict[str, Any]]]:
    wb = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
    order_sheet = next(
        (name for name in wb.sheetnames if _key(name) in {"ordine", "sheet1"}),
        None,
    )
    if order_sheet is None:
        raise ValueError('Il file Data Ordine deve contenere un foglio Ordine o Sheet1.')
    data = _read_sheet_rows(wb[order_sheet])
    dispo_sheet = next(
        (name for name in wb.sheetnames
         if "dispo" in _key(name) and "bagno" in _key(name)),
        None,
    )
    dispo_rows = _read_sheet_rows(wb[dispo_sheet]) if dispo_sheet else []
    raw_sheet = next(
        (name for name in wb.sheetnames
         if "filato" in _key(name) or "tinturia" in _key(name)),
        None,
    )
    raw_rows = _read_sheet_rows(wb[raw_sheet]) if raw_sheet else []
    if dispo_path and dispo_path != input_path:
        dispo_rows = load_dispo_bagno_rows(dispo_path)
    if not data:
        raise ValueError("Sheet1 non contiene righe d'ordine.")
    dispo_by_riga = {
        _dispo_suffix(_get(r, "Dispo")): r
        for r in dispo_rows if _clean(_get(r, "Dispo"))
    }
    raw_by_article = {_clean(_get(r, "Articolo")).upper(): r for r in raw_rows}
    result: list[OrderRecord] = []
    for row in data:
        code = _clean(_get(row, "Cliente"))
        if code not in {"3009", "3004"}:
            continue
        article = _clean(_get(row, "Articolo"))
        raw = raw_by_article.get(("G" + article[1:]).upper(), {}) if article else {}
        additional = _clean(_get(row, "Descrizione aggiuntiva ordine"))
        raw_batch = _clean(_get(raw, "Partita.GG", "Partita")) or _first_match(additional, r"PG[- ]*([0-9]+)")
        riga = _clean(_get(row, "Riga"))
        dispo = dispo_by_riga.get(_dispo_suffix(riga), {})
        result.append(OrderRecord(
            customer_code=code,
            customer_name=_clean(_get(row, "__col2")) or ("MED" if code == "3004" else "ELVY"),
            article=article,
            description=additional or article,
            additional_raw=additional,
            color_code=_clean(_get(row, "Colore")),
            color_name=_clean(_get(row, "__col7")) or _clean(_get(row, "Colore")),
            order_no=_order_number(_get(row, "Ordine")),
            order_row=riga,
            colored_batch=_clean(_get(row, "Partita")),
            raw_batch=raw_batch,
            quantity_cones=_number(_get(row, "Ordinata", "Assegnata")) or _number(_get(raw, "عدد")),
            raw_weight=_number(_get(raw, "وزن", "Peso")),
            dispo=_clean(_get(dispo, "Dispo")),
            bagno=_clean(_get(dispo, "Field2", "Bagno")),
            delivery=_get(row, "Consegna"),
            title=_clean(_get(raw, "Titolo")) or _clean(_get(row, "Descrizione")),
        ))
    wb.close()
    if not result:
        raise ValueError("Non sono state trovate righe con Cliente 3004 o 3009.")
    return result, raw_rows


# ---------------------------------------------------------------------------
# EL KAMAL -- third customer, different source shape entirely: their own
# order export (Sheet1) already arrives with Titolo/KG/M-C/Bagno computed,
# instead of the raw ERP dump ELVY/MED use. The companion Dispo-Bagno file
# is a semicolon CSV (same columns as the Doispo-Bagno sheet) rather than a
# second sheet in the same workbook. Sheet1's own "Dispo" text embeds the
# row's sequence number as 'RIGA N', which is what joins to the CSV's
# 'D-...-00N' suffix -- same join pattern as MED's Riga column, just
# extracted from a different place.
# ---------------------------------------------------------------------------

def _read_dispo_csv(path: Path) -> list[dict[str, Any]]:
    import csv
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";"
        reader = csv.reader(f, dialect)
        rows = list(reader)
    if not rows:
        return []
    headers = [_key(x) or f"__col{i + 1}" for i, x in enumerate(rows[0])]
    out = []
    for values in rows[1:]:
        if not any(_clean(x) for x in values):
            continue
        rec: dict[str, Any] = {}
        for i, value in enumerate(values):
            if i < len(headers):
                rec.setdefault(headers[i], value)
                rec.setdefault(f"__col{i + 1}", value)
        out.append(rec)
    return out


def load_el_kamal_order(data_path: Path, dispo_path: Path | None = None) -> tuple[list[OrderRecord], list[dict[str, Any]]]:
    wb = openpyxl.load_workbook(data_path, data_only=True, read_only=True)
    if "Sheet1" not in wb.sheetnames:
        raise ValueError('Il file Data Ordine EL KAMAL deve contenere il foglio "Sheet1".')
    data = _read_sheet_rows(wb["Sheet1"])
    wb.close()
    if not data:
        raise ValueError("Sheet1 non contiene righe d'ordine.")

    dispo_by_riga: dict[str, dict[str, Any]] = {}
    if dispo_path and dispo_path.is_file():
        dispo_rows = load_dispo_bagno_rows(dispo_path)
        dispo_by_riga = {
            _dispo_suffix(_get(r, "Dispo")): r
            for r in dispo_rows if _clean(_get(r, "Dispo"))
        }

    result: list[OrderRecord] = []
    for row in data:
        dispo_text = _clean(_get(row, "Dispo"))
        riga = _first_match(dispo_text, r"RIGA\s*([0-9]+)")
        dispo_row = dispo_by_riga.get(_dispo_suffix(riga), {}) if riga else {}
        bagno = _clean(_get(row, "Bagno")) or _clean(_get(dispo_row, "Field2", "Bagno"))
        result.append(OrderRecord(
            customer_code="3019",
            customer_name=_clean(_get(row, "Clienti")) or "EL KAMAL",
            article=_clean(_get(row, "CODICE", "Articolo")),
            description=_clean(_get(row, "COMMENTO RIGO ORDINE")),
            additional_raw=_clean(_get(row, "COMMENTO RIGO ORDINE")),
            color_code=_clean(_get(row, "COLORE", "Codice")),
            color_name=_clean(_get(row, "DESCR COL", "Colore")),
            order_no=_order_number(_get(row, "ORDINE", "Ordine")),
            order_row=riga,
            colored_batch=_clean(_get(row, "Partita Col", "Partita")),
            raw_batch=_clean(_get(row, "Partita GG")),
            quantity_cones=_number(_get(row, "ROCCHE", "Rocche")),
            raw_weight=_number(_get(row, "KG")),
            dispo=dispo_text or _clean(_get(dispo_row, "Dispo")),
            bagno=bagno,
            delivery=_get(row, "CONSEGNA", "Consegna"),
            title=_clean(_get(row, "TITOLO", "Titolo")),
            formato=_clean(_get(row, "Formmato", "Formato")) or "7777",
            machine=_clean(_get(row, "M/C")),
            commento=_clean(_get(row, "COMMENTO RIGO ORDINE")),
            kg=_number(_get(row, "KG")),
        ))
    if not result:
        raise ValueError("Non sono state trovate righe d'ordine EL KAMAL in Sheet1.")
    return result, []


def build_el_kamal_stem(records: list["OrderRecord"]) -> str:
    """e.g. 'D-505444-EL_KAMAL' -- confirmed against a real reference
    ticket filename the user sent (no year in it, unlike MED)."""
    dispo_num = _dispo_number(records[0].dispo) if records else ""
    return f"D-{dispo_num}-EL_KAMAL" if dispo_num else "EL_KAMAL"


def enrich_records(
    records: list[OrderRecord],
    customer: str,
    codes_map: dict[str, str] | None = None,
    densita_map: dict[int, dict[str, Any]] | None = None,
    vmm_ratio_map: dict[int, float] | None = None,
    price_lookup: dict[tuple, tuple] | None = None,
) -> None:
    """Fills in Titolo, Commento, M/C, KG, Densita, Color Tube, VMM22,
    Prezzo and the MED "Partita MED"/"Cliente MED" pair, in place.
    Anything whose source file was not provided is left blank, same as
    before, instead of raising.
    """
    codes_map = codes_map or {}
    densita_map = densita_map or {}
    vmm_ratio_map = vmm_ratio_map or {}
    price_lookup = price_lookup or {}

    # --- M/C: group by Bagno, sum Rocche (+ MED Polmoni bonus). EL KAMAL's
    # own source already gives a correct M/C per row, so it's left alone.
    if customer != "EL_KAMAL":
        groups: dict[str, list[OrderRecord]] = {}
        for r in records:
            groups.setdefault(r.bagno or "", []).append(r)
        for bagno, members in groups.items():
            total = sum(_number(m.quantity_cones) or 0 for m in members)
            if customer == "MED":
                total += sum(_polmoni_extra(m.additional_raw) for m in members)
            extra = MACHINE_SIZE_TABLE_MED_EXTRA if customer == "MED" else None
            machine = _machine_for_count(total, extra=extra)
            for m in members:
                m.machine = machine

    for r in records:
        # Titolo: Articoli.xlsx lookup takes priority (fuller description);
        # the raw yarn sheet's shorter Titolo is the fallback.
        r.title = _titolo_lookup(r.article, codes_map) or r.title

        # EL KAMAL's Commento (COMMENTO RIGO ORDINE) already arrives correct
        # from the source and isn't a dash-encoded comment like ELVY/MED's.
        if customer != "EL_KAMAL":
            r.commento = _commento_for(customer, r.additional_raw)

        if customer == "MED":
            r.partita_med = _partita_med(r.additional_raw)
            r.cliente_med = _cliente_med(r.additional_raw)

        try:
            batch_key = int(_number(r.raw_batch)) if _number(r.raw_batch) is not None else None
        except (TypeError, ValueError):
            batch_key = None

        if batch_key is not None and batch_key in densita_map:
            entry = densita_map[batch_key]
            peso_net = entry.get("peso_net")
            r.kg = int(round(peso_net * (_number(r.quantity_cones) or 0))) if peso_net is not None else r.raw_weight
            r.densita = entry.get("densita", "")
            r.color_tube = entry.get("color_tube", "")
        else:
            r.kg = r.raw_weight

        if batch_key is not None and batch_key in vmm_ratio_map:
            kg_per_cone = vmm_ratio_map[batch_key]
            r.vmm22 = round(kg_per_cone * (_number(r.quantity_cones) or 0), 2)

        r.prezzo = _prezzo_for(r.article, r.color_code, price_lookup)

    if customer == "ELVY":
        compute_delivery_date(records)


def _filato_rows(
    records: list["OrderRecord"],
    raw_rows: list[dict[str, Any]],
    magazino_summary=None,
) -> list[dict[str, Any]]:
    """One row per raw batch, for the post-ERP-entry (Create EXCEL+Biglietti)
    stage: stock availability was already checked earlier, on the Ordine
    page, before the order went into the system -- so Rocche here is simply
    the order's own total cone count for that batch, not a fresh warehouse
    lookup. Peso is still derived from Magazino, but scaled to that same
    order quantity: (warehouse weight / warehouse cones) for the batch,
    times how many cones THIS order actually needs, i.e. the total weight
    for the cones being pulled -- not the warehouse's full stock weight.
    """
    raw_by_article = {_clean(_get(r, "Articolo")).upper(): r for r in raw_rows}

    totals: dict[str, float] = {}
    for rec in records:
        key = _clean(rec.raw_batch)
        if not key:
            continue
        totals[key] = totals.get(key, 0) + (_number(rec.quantity_cones) or 0)

    warehouse_rate = {}
    rate_by_partita = {}

    def stock_key(value: Any) -> str:
        text = _clean(value)
        number = _number(text)
        return str(int(number)) if number is not None and float(number).is_integer() else text

    if magazino_summary is not None:
        for row in magazino_summary.itertuples(index=False):
            article = _clean(getattr(row, "articolo", "")).upper()
            partita = stock_key(getattr(row, "partita", ""))
            if not article or not partita:
                continue
            try:
                mag_rocche = float(getattr(row, "mag_rocche", 0) or 0)
                mag_peso = float(getattr(row, "mag_peso", 0) or 0)
            except (TypeError, ValueError):
                continue
            if mag_rocche <= 0:
                continue
            rate = mag_peso / mag_rocche
            warehouse_rate[(article, partita)] = rate
            rate_by_partita.setdefault(partita, set()).add((article, rate))

    out = []
    seen: set[str] = set()
    for rec in records:
        key = _clean(rec.raw_batch)
        if not key or key in seen:
            continue
        seen.add(key)
        article_g = ("G" + rec.article[1:]) if rec.article[:1].upper() == "C" else rec.article
        raw = raw_by_article.get(article_g.upper(), {})
        stock_partita = stock_key(key)
        order_rocche = totals[key]
        per_cone_rate = warehouse_rate.get((article_g.upper(), stock_partita))
        if per_cone_rate is None:
            # Some ERP exports contain a visually similar article code, e.g.
            # G1300275 in the order and G130027S in Magazino.  A Partita is
            # unique in the stock export, so use it only when it has one
            # unambiguous warehouse article.
            partita_matches = rate_by_partita.get(stock_partita, set())
            if len(partita_matches) == 1:
                per_cone_rate = next(iter(partita_matches))[1]
        peso = round(per_cone_rate * order_rocche, 2) if per_cone_rate is not None else _number(_get(raw, "وزن", "Peso"))
        out.append({
            "Articolo": article_g,
            "Titolo": rec.title or _clean(_get(raw, "Titolo")),
            "Partita": key,
            "Rocche": order_rocche,
            "Peso": peso,
            "تحضير خام": _clean(_get(raw, "Custom", "تحضير خام")) or "تحضير خام",
        })
    return out


def export_workbook(
    path: Path,
    records: list[OrderRecord],
    raw_rows: list[dict[str, Any]],
    include_filato: bool = True,
    stem: str = "",
    customer: str = "",
    magazino_summary=None,
) -> None:
    customer = customer or ("ELVY" if records[0].customer_code == "3009" else "MED")
    # The Create (EXCEL+Biglietti) extract is consumed in Partita Col order.
    # Sort numerically (not lexicographically, so 20 comes after 3) and keep
    # rows with an empty/non-numeric Partita Col at the end.
    records = sorted(
        records,
        key=lambda record: (
            _number(record.colored_batch) is None,
            _number(record.colored_batch) if _number(record.colored_batch) is not None else 0,
        ),
    )
    wb = Workbook()
    ws = wb.active
    ws.title = customer
    common = ["Dispo/Riga", "Articolo", "Titolo", "Formato", "Ordine", "Codice", "Colore", "Rocche", "KG", "M/C", "Partita Col", "Consegna", "Commento", "Bagno", "Cliente", "Partita GG"]
    extra_cols = ["Color Tube", "VMM22", "Prezzo", "Densita` (360-390)"]
    if customer == "MED":
        headers = common + ["Partita MED", "Cliente MED", "POLMON"] + extra_cols
    elif customer == "EL_KAMAL":
        # EL KAMAL's source doesn't carry these -- dropped per request, not
        # just left blank.
        headers = common
    else:
        headers = common + ["Delivery Date"] + extra_cols
    ws.append(headers)
    for r in records:
        row = [r.dispo, r.article, r.title, r.formato, r.order_no, r.color_code, r.color_name,
               r.quantity_cones, r.kg, r.machine, r.colored_batch, r.delivery,
               r.commento, r.bagno, r.customer_name, r.raw_batch]
        if customer == "MED":
            polmon = _polmoni_segment(r.additional_raw)
            row += [r.partita_med, r.cliente_med, polmon, r.color_tube, r.vmm22, r.prezzo, r.densita]
        elif customer == "EL_KAMAL":
            pass
        else:
            row += [r.delivery_date, r.color_tube, r.vmm22, r.prezzo, r.densita]
        ws.append(row)
    if include_filato:
        fws = wb.create_sheet("Filato x Tinturia")
        fheaders = ["Articolo", "Titolo", "Partita", "Rocche", "Peso", "تحضير خام"]
        fws.append(fheaders)
        for r in _filato_rows(records, raw_rows, magazino_summary):
            fws.append([r[h] for h in fheaders])
        _style_sheet(fws)
    _style_sheet(
        ws,
        date_columns=("Consegna", "Delivery Date"),
        duplicate_highlight_columns=("Bagno",),
        range_highlight_columns={"Densita` (360-390)": (360, 390)},
        days_until_highlight_columns={"Delivery Date": 4},
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


# ---------------------------------------------------------------------------
# Shared Create (Excel + Biglietti) workbook
# ---------------------------------------------------------------------------

CREATE_EXCEL_HEADERS = [
    "Dispo/Riga", "Cliente", "Articolo", "Titolo", "Formato", "Ordine", "Codice",
    "Colore", "Rocche", "KG", "M/C", "Partita Col", "Consegna",
    "Commento", "Bagno", "Partita GG", "Delivery Date", "Partita MED",
    "Cliente MED", "POLMON", "Color Tube", "VMM22", "Prezzo",
    "Densita` (360-390)",
]


def _create_excel_row(record: OrderRecord, customer: str) -> list[Any]:
    """Return the stable superset row used by the shared order workbook."""
    polmon = _polmoni_segment(record.additional_raw) if customer == "MED" else ""
    return [
        record.dispo, record.customer_name or customer,
        record.article, record.title, record.formato, record.order_no,
        record.color_code, record.color_name, record.quantity_cones,
        record.kg, record.machine, record.colored_batch, record.delivery,
        record.commento, record.bagno, record.raw_batch, record.delivery_date,
        record.partita_med, record.cliente_med, polmon, record.color_tube,
        record.vmm22, record.prezzo, record.densita,
    ]


def _partita_key(value: Any) -> str:
    value = _clean(value)
    number = _number(value)
    if number is not None:
        return str(int(number)) if float(number).is_integer() else str(number)
    return value.casefold()


def _is_pg_x(value: Any) -> bool:
    text = _clean(value).upper().replace(" ", "")
    return not text or text in {"X", "PG-X", "PGX"}


def append_create_excel(path: Path, records: list[OrderRecord], customer: str) -> dict[str, Any]:
    """Create or append the shared workbook used by the dyeing review flow.

    ``Partita Col`` is the business key: rows whose key already exists in the
    workbook are skipped.  The complete data set is then sorted numerically by
    that column, with blank/non-numeric values at the end.
    """
    from openpyxl import Workbook, load_workbook

    path.parent.mkdir(parents=True, exist_ok=True)
    if path.is_file():
        wb = load_workbook(path)
        ws = wb["Orders"] if "Orders" in wb.sheetnames else wb.active
        headers = [_clean(c.value) for c in ws[1]]
        if "Partita Col" not in headers:
            wb.close()
            raise ValueError("The selected Excel file has no 'Partita Col' column.")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Orders"
        headers = list(CREATE_EXCEL_HEADERS)
        ws.append(headers)

    pg_ws = wb["PG-X"] if "PG-X" in wb.sheetnames else wb.create_sheet("PG-X")
    if pg_ws.max_row == 1 and all(c.value is None for c in pg_ws[1]):
        pg_ws.delete_rows(1)
    if pg_ws.max_row == 0:
        pg_ws.append(headers)
    pg_headers = [_clean(c.value) for c in pg_ws[1]]
    if not pg_headers or "Partita Col" not in pg_headers:
        if pg_ws.max_row > 0:
            pg_ws.delete_rows(1, pg_ws.max_row)
        pg_headers = list(headers)
        pg_ws.append(pg_headers)

    # Migrate legacy shared files: rows without Partita GG belong on PG-X.
    raw_col = next((i + 1 for i, h in enumerate(headers) if _key(h) == _key("Partita GG")), None)
    if raw_col and ws.max_row > 1:
        move_rows = []
        for row_idx in range(2, ws.max_row + 1):
            if _is_pg_x(ws.cell(row=row_idx, column=raw_col).value):
                move_rows.append([ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, ws.max_column + 1)])
        for row in move_rows:
            pg_ws.append([dict(zip(headers, row)).get(header, "") for header in pg_headers])
        for row_idx in range(ws.max_row, 1, -1):
            if _is_pg_x(ws.cell(row=row_idx, column=raw_col).value):
                ws.delete_rows(row_idx, 1)

    def col(name: str) -> int | None:
        wanted = _key(name)
        for idx, header in enumerate(headers, start=1):
            if _key(header) == wanted:
                return idx
        return None

    existing = set()
    for candidate_ws, candidate_headers in ((ws, headers), (pg_ws, pg_headers)):
        candidate_col = next((i + 1 for i, h in enumerate(candidate_headers) if _key(h) == _key("Partita Col")), None)
        if candidate_col:
            existing.update(
                _partita_key(candidate_ws.cell(row=row, column=candidate_col).value)
                for row in range(2, candidate_ws.max_row + 1)
                if _partita_key(candidate_ws.cell(row=row, column=candidate_col).value)
            )
    added = 0
    skipped = 0
    for record in records:
        key = _partita_key(record.colored_batch)
        if not key or key in existing:
            skipped += 1
            continue
        values = dict(zip(CREATE_EXCEL_HEADERS, _create_excel_row(record, customer)))
        target_ws = pg_ws if _is_pg_x(record.raw_batch) else ws
        target_headers = pg_headers if target_ws is pg_ws else headers
        target_ws.append([values.get(header, "") for header in target_headers])
        existing.add(key)
        added += 1

    def sort_sheet(target_ws, target_headers):
        target_col = next(i + 1 for i, h in enumerate(target_headers) if _key(h) == _key("Partita Col"))
        rows = list(target_ws.iter_rows(min_row=2, values_only=True))
        rows.sort(key=lambda row: (
            _number(row[target_col - 1]) is None,
            _number(row[target_col - 1]) if _number(row[target_col - 1]) is not None else 0,
        ))
        if target_ws.max_row > 1:
            target_ws.delete_rows(2, target_ws.max_row - 1)
        for row in rows:
            target_ws.append(list(row))
        _style_sheet(target_ws, date_columns=("Consegna", "Delivery Date"))

    sort_sheet(ws, headers)
    sort_sheet(pg_ws, pg_headers)
    wb.save(path)
    wb.close()
    return {"added": added, "skipped": skipped, "total": ws.max_row - 1 + pg_ws.max_row - 1}


def load_create_excel_records(path: Path, partita_gg: str = "", sheet_name: str = "Orders") -> list[OrderRecord]:
    """Read shared-workbook rows, optionally filtered by ``Partita GG``."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        elif sheet_name == "Orders":
            ws = wb.active
        else:
            return []
        rows = _read_sheet_rows(ws)
    finally:
        wb.close()
    wanted = _partita_key(partita_gg) if _clean(partita_gg) else ""
    out: list[OrderRecord] = []
    for row in rows:
        if wanted and _partita_key(_get(row, "Partita GG")) != wanted:
            continue
        customer = _clean(_get(row, "Cliente")) or "ELVY"
        code = "3004" if customer.upper().startswith("MED") else "3009"
        out.append(OrderRecord(
            customer_code=code, customer_name=customer,
            article=_clean(_get(row, "Articolo")),
            description=_clean(_get(row, "Titolo")), additional_raw=_clean(_get(row, "Commento")),
            color_code=_clean(_get(row, "Codice")), color_name=_clean(_get(row, "Colore")),
            order_no=_clean(_get(row, "Ordine")), order_row="",
            colored_batch=_clean(_get(row, "Partita Col")), raw_batch=_clean(_get(row, "Partita GG")),
            quantity_cones=_number(_get(row, "Rocche")), raw_weight=_number(_get(row, "KG")),
            dispo=_clean(_get(row, "Dispo/Riga")), bagno=_clean(_get(row, "Bagno")),
            delivery=_get(row, "Consegna"), machine=_clean(_get(row, "M/C")),
            title=_clean(_get(row, "Titolo")), commento=_clean(_get(row, "Commento")),
            partita_med=_clean(_get(row, "Partita MED")), cliente_med=_clean(_get(row, "Cliente MED")),
            kg=_number(_get(row, "KG")), color_tube=_clean(_get(row, "Color Tube")),
            vmm22=_number(_get(row, "VMM22")), prezzo=_get(row, "Prezzo"),
            densita=_get(row, "Densita` (360-390)"), delivery_date=_get(row, "Delivery Date"),
        ))
    if not out and not _clean(partita_gg):
        return []
    if not out:
        raise ValueError(f"No rows were found for Partita GG '{partita_gg}'.")
    return out


def update_pg_x_row(path: Path, partita_col: str, updates: dict[str, Any]) -> int:
    """Update editable fields for a PG-X color without assigning it yet."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        if "PG-X" not in wb.sheetnames:
            raise ValueError("The shared Excel has no PG-X sheet.")
        ws = wb["PG-X"]
        headers = [_clean(c.value) for c in ws[1]]

        def header_col(name):
            wanted = _key(name)
            return next((i + 1 for i, value in enumerate(headers) if _key(value) == wanted), None)

        partita_col_idx = header_col("Partita Col")
        if not partita_col_idx:
            raise ValueError("Partita Col column is missing.")
        wanted = _partita_key(partita_col)
        matches = [
            row_idx for row_idx in range(2, ws.max_row + 1)
            if _partita_key(ws.cell(row=row_idx, column=partita_col_idx).value) == wanted
        ]
        if not matches:
            raise ValueError(f"Partita Col '{partita_col}' was not found in PG-X.")

        columns = {name: header_col(name) for name in updates}
        for name, value in updates.items():
            column = columns[name]
            if column:
                for row_idx in matches:
                    ws.cell(row=row_idx, column=column).value = value
        _style_sheet(ws, date_columns=("Consegna", "Delivery Date"))
        wb.save(path)
        return len(matches)
    finally:
        wb.close()


def update_order_row(
    path: Path,
    partita_col: str,
    updates: dict[str, Any],
    magazino_summary=None,
    allow_article_mismatch: bool = False,
) -> dict[str, Any]:
    """Edit an Orders row; clearing Partita GG moves it back to PG-X."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        if "Orders" not in wb.sheetnames or "PG-X" not in wb.sheetnames:
            raise ValueError("The shared Excel must contain Orders and PG-X sheets.")
        orders_ws, pg_ws = wb["Orders"], wb["PG-X"]
        order_headers = [_clean(c.value) for c in orders_ws[1]]
        pg_headers = [_clean(c.value) for c in pg_ws[1]]

        def header_col(headers, name):
            wanted = _key(name)
            return next((i + 1 for i, value in enumerate(headers) if _key(value) == wanted), None)

        part_col_idx = header_col(order_headers, "Partita Col")
        gg_col_idx = header_col(order_headers, "Partita GG")
        if not part_col_idx or not gg_col_idx:
            raise ValueError("Orders is missing Partita Col or Partita GG columns.")
        wanted = _partita_key(partita_col)
        matches = [
            row_idx for row_idx in range(2, orders_ws.max_row + 1)
            if _partita_key(orders_ws.cell(row=row_idx, column=part_col_idx).value) == wanted
        ]
        if not matches:
            raise ValueError(f"Partita Col '{partita_col}' was not found in Orders.")

        # "Partita GG" may not be in *updates* at all (e.g. a caller only
        # changing Articolo) -- that must never be read as "cleared" (which
        # would wrongly move the row to PG-X) nor let a changed Articolo
        # skip validation against whatever Partita GG the row already has.
        # Only an explicit, empty "Partita GG" in updates means "clear it".
        partita_gg_submitted = "Partita GG" in updates
        if partita_gg_submitted:
            new_gg = _clean(updates.get("Partita GG", ""))
        else:
            new_gg = _clean(orders_ws.cell(row=matches[0], column=gg_col_idx).value)
        if new_gg and magazino_summary is not None:
            article_col_idx = header_col(order_headers, "Articolo")
            color_article = _clean(updates.get("Articolo"))
            if not color_article and article_col_idx:
                color_article = _clean(orders_ws.cell(row=matches[0], column=article_col_idx).value)
            color_article = color_article.upper()
            expected_raw = "G" + color_article[1:] if color_article.startswith("C") else color_article
            batch_key = _partita_key(new_gg)
            matching_stock = magazino_summary[
                (magazino_summary["articolo"].astype(str).str.strip().str.upper() == expected_raw)
                & (magazino_summary["partita"].map(_partita_key) == batch_key)
            ]
            if matching_stock.empty and not allow_article_mismatch:
                raise ValueError(
                    f"Partita GG {new_gg} does not belong to article {expected_raw} "
                    f"required by color article {color_article}."
                )

        updates_by_header = {header_col(order_headers, name): value for name, value in updates.items()}
        for row_idx in matches:
            for column, value in updates_by_header.items():
                if column:
                    orders_ws.cell(row=row_idx, column=column).value = value

        moved_to_pgx = partita_gg_submitted and not new_gg
        if moved_to_pgx:
            for row_idx in matches:
                pg_values = {header: orders_ws.cell(row=row_idx, column=idx + 1).value for idx, header in enumerate(order_headers)}
                pg_ws.append([pg_values.get(header, "") for header in pg_headers])
            for row_idx in reversed(matches):
                orders_ws.delete_rows(row_idx, 1)

        _style_sheet(orders_ws, date_columns=("Consegna", "Delivery Date"))
        _style_sheet(pg_ws, date_columns=("Consegna", "Delivery Date"))
        wb.save(path)
        return {"updated": len(matches), "moved_to_pgx": moved_to_pgx}
    finally:
        wb.close()


def save_pg_x_partita(
    path: Path,
    partita_col: str,
    partita_gg: str,
    densita_map: dict[int, dict[str, Any]] | None = None,
    vmm_ratio_map: dict[int, float] | None = None,
    magazino_summary=None,
    allow_article_mismatch: bool = False,
) -> dict[str, Any]:
    """Assign raw yarn to a PG-X color, enrich it, and move it to Orders."""
    from openpyxl import load_workbook

    densita_map = densita_map or {}
    vmm_ratio_map = vmm_ratio_map or {}
    wb = load_workbook(path)
    if "PG-X" not in wb.sheetnames:
        wb.close()
        raise ValueError("The shared Excel has no PG-X sheet.")
    pg_ws = wb["PG-X"]
    orders_ws = wb["Orders"] if "Orders" in wb.sheetnames else wb.active
    pg_headers = [_clean(c.value) for c in pg_ws[1]]
    order_headers = [_clean(c.value) for c in orders_ws[1]]

    def header_col(headers, name):
        wanted = _key(name)
        return next((i + 1 for i, value in enumerate(headers) if _key(value) == wanted), None)

    col_partita_col = header_col(pg_headers, "Partita Col")
    col_partita_gg = header_col(pg_headers, "Partita GG")
    if not col_partita_col or not col_partita_gg:
        wb.close()
        raise ValueError("PG-X is missing Partita Col or Partita GG columns.")
    wanted_col = _partita_key(partita_col)
    matching_rows = [
        row_idx for row_idx in range(2, pg_ws.max_row + 1)
        if _partita_key(pg_ws.cell(row=row_idx, column=col_partita_col).value) == wanted_col
    ]
    if not matching_rows:
        wb.close()
        raise ValueError(f"Partita Col '{partita_col}' was not found in PG-X.")

    source_map = {header: index + 1 for index, header in enumerate(pg_headers)}
    updates = {"Partita GG": partita_gg}
    batch_number = _number(partita_gg)
    density_entry = densita_map.get(int(batch_number)) if batch_number is not None else None
    if density_entry:
        updates["KG"] = [
            _number(pg_ws.cell(row=row_idx, column=header_col(pg_headers, "Rocche")).value) or 0
            for row_idx in matching_rows
        ]
    moved_rows = []
    available_total = 0.0
    for row_idx in matching_rows:
        rocche_col = header_col(pg_headers, "Rocche")
        rocche = _number(pg_ws.cell(row=row_idx, column=rocche_col).value) if rocche_col else 0
        values = {header: pg_ws.cell(row=row_idx, column=col_idx).value for header, col_idx in source_map.items()}
        values["Partita GG"] = partita_gg
        if magazino_summary is not None:
            color_article = _clean(values.get("Articolo")).upper()
            expected_raw_article = "G" + color_article[1:] if color_article.startswith("C") else color_article
            batch_key = str(int(batch_number)) if batch_number is not None else _clean(partita_gg)
            matching_stock = magazino_summary[
                (magazino_summary["articolo"].astype(str).str.strip().str.upper() == expected_raw_article)
                & (magazino_summary["partita"].map(_partita_key) == batch_key)
            ]
            if matching_stock.empty:
                if not allow_article_mismatch:
                    wb.close()
                    raise ValueError(
                        f"Partita GG {partita_gg} does not belong to article {expected_raw_article} "
                        f"required by color article {color_article}."
                    )
            if not matching_stock.empty:
                available_total += float(matching_stock.iloc[0].get("mag_rocche", 0) or 0)
        if density_entry:
            if density_entry.get("peso_net") is not None:
                values["KG"] = int(round(density_entry["peso_net"] * (rocche or 0)))
            values["Densita` (360-390)"] = density_entry.get("densita", "")
            values["Color Tube"] = density_entry.get("color_tube", "")
        if batch_number is not None and int(batch_number) in vmm_ratio_map:
            values["VMM22"] = round(vmm_ratio_map[int(batch_number)] * (rocche or 0), 2)
        moved_rows.append([values.get(header, "") for header in order_headers])

    for row_idx, row in zip(matching_rows, moved_rows):
        values = dict(zip(order_headers, row))
        for header, col_idx in ((header, index + 1) for index, header in enumerate(pg_headers)):
            pg_ws.cell(row=row_idx, column=col_idx).value = values.get(header, "")

    def sort_sheet(ws, headers):
        col_idx = header_col(headers, "Partita Col")
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        rows.sort(key=lambda row: (_number(row[col_idx - 1]) is None, _number(row[col_idx - 1]) or 0))
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for row in rows:
            ws.append(list(row))
        _style_sheet(ws, date_columns=("Consegna", "Delivery Date"))

    sort_sheet(pg_ws, pg_headers)
    sort_sheet(orders_ws, order_headers)
    wb.save(path)
    wb.close()
    return {"updated": len(matching_rows), "partita_gg": partita_gg, "available": available_total}


def move_pg_x_to_orders(path: Path, partita_col: str) -> int:
    """Move the selected PG-X color row to Orders."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    if "PG-X" not in wb.sheetnames or "Orders" not in wb.sheetnames:
        wb.close()
        raise ValueError("The shared Excel must contain Orders and PG-X sheets.")
    pg_ws, orders_ws = wb["PG-X"], wb["Orders"]
    pg_headers = [_clean(c.value) for c in pg_ws[1]]
    order_headers = [_clean(c.value) for c in orders_ws[1]]
    pg_col = next((i + 1 for i, h in enumerate(pg_headers) if _key(h) == _key("Partita Col")), None)
    order_col = next((i + 1 for i, h in enumerate(order_headers) if _key(h) == _key("Partita Col")), None)
    if not pg_col or not order_col:
        wb.close()
        raise ValueError("Partita Col column is missing.")
    wanted = _partita_key(partita_col)
    matches = [r for r in range(2, pg_ws.max_row + 1) if _partita_key(pg_ws.cell(r, pg_col).value) == wanted]
    if not matches:
        wb.close()
        raise ValueError(f"Partita Col '{partita_col}' was not found in PG-X.")
    existing = {_partita_key(orders_ws.cell(r, order_col).value) for r in range(2, orders_ws.max_row + 1)}
    for row_idx in matches:
        values = {header: pg_ws.cell(row_idx, idx + 1).value for idx, header in enumerate(pg_headers)}
        row = [values.get(header, "") for header in order_headers]
        if _partita_key(row[order_col - 1]) not in existing:
            orders_ws.append(row)
            existing.add(_partita_key(row[order_col - 1]))
    for row_idx in reversed(matches):
        pg_ws.delete_rows(row_idx, 1)
    wb.save(path)
    wb.close()
    return len(matches)


def delete_pg_x_row(path: Path, partita_col: str) -> int:
    """Delete the selected PG-X color row."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    if "PG-X" not in wb.sheetnames:
        wb.close()
        raise ValueError("The shared Excel has no PG-X sheet.")
    ws = wb["PG-X"]
    headers = [_clean(c.value) for c in ws[1]]
    col_idx = next((i + 1 for i, h in enumerate(headers) if _key(h) == _key("Partita Col")), None)
    if not col_idx:
        wb.close()
        raise ValueError("Partita Col column is missing.")
    wanted = _partita_key(partita_col)
    matches = [r for r in range(2, ws.max_row + 1) if _partita_key(ws.cell(r, col_idx).value) == wanted]
    if not matches:
        wb.close()
        raise ValueError(f"Partita Col '{partita_col}' was not found in PG-X.")
    for row_idx in reversed(matches):
        ws.delete_rows(row_idx, 1)
    wb.save(path)
    wb.close()
    return len(matches)


def delete_shipped_shared_rows(path: Path, partita_cols, sheet_name: str) -> int:
    """Delete rows in one shared-workbook sheet whose Partita Col shipped."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"The shared Excel has no '{sheet_name}' sheet.")
        ws = wb[sheet_name]
        headers = [_clean(c.value) for c in ws[1]]
        col_idx = next((i + 1 for i, h in enumerate(headers) if _key(h) == _key("Partita Col")), None)
        if not col_idx:
            raise ValueError("Partita Col column is missing.")
        shipped = {_partita_key(value) for value in partita_cols if _partita_key(value)}
        matches = [
            row_idx for row_idx in range(2, ws.max_row + 1)
            if _partita_key(ws.cell(row=row_idx, column=col_idx).value) in shipped
        ]
        for row_idx in reversed(matches):
            ws.delete_rows(row_idx, 1)
        _style_sheet(ws, date_columns=("Consegna", "Delivery Date"))
        wb.save(path)
        return len(matches)
    finally:
        wb.close()


def export_filato_workbook(
    path: Path,
    records: list["OrderRecord"],
    raw_rows: list[dict[str, Any]],
    magazino_summary=None,
) -> None:
    """Export only the optional ``Filato x Tinturia`` workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Filato x Tinturia"
    headers = ["Articolo", "Titolo", "Partita", "Rocche", "Peso", "تحضير خام"]
    ws.append(headers)
    for r in _filato_rows(records, raw_rows, magazino_summary):
        ws.append([r[h] for h in headers])
    _style_sheet(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def _style_sheet(
    ws,
    date_columns: tuple[str, ...] = (),
    duplicate_highlight_columns: tuple[str, ...] = (),
    range_highlight_columns: dict[str, tuple[float, float]] | None = None,
    days_until_highlight_columns: dict[str, int] | None = None,
) -> None:
    """Bold/blue header, every cell (not just the header) centered, optional
    date-only number format for the given columns, optional red highlight
    for duplicate values in the given columns (e.g. Bagno), optional red
    highlight for numbers outside a (min, max) range in the given columns
    (e.g. Densita`(360-390)), and optional red highlight for a date column
    where fewer than N days remain until that date (e.g. Delivery Date),
    recalculated live against today's date every time the file is opened."""
    fill = PatternFill("solid", fgColor="FF1F4E78")
    header = [c.value for c in ws[1]]
    last_row = ws.max_row

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    band_fill = PatternFill("solid", fgColor="FFDCE6F1")
    for row in ws.iter_rows():
        banded = row[0].row > 1 and (row[0].row % 2 == 0)
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(cell.row == 1))
            cell.border = border
            if banded:
                cell.fill = band_fill
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill

    for col_name in date_columns:
        if col_name in header:
            col_idx = header.index(col_name) + 1
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=col_idx).number_format = "dd/mm/yyyy"

    if last_row > 1:
        from openpyxl.formatting.rule import FormulaRule
        # Differential-format (conditional-formatting) fills render from
        # bgColor, not fgColor -- a plain PatternFill("solid", fgColor=...)
        # writes only fgColor into the dxf and Excel shows no highlight at
        # all even though the rule fires. Setting both covers it either way.
        red_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")

        for col_name in duplicate_highlight_columns:
            if col_name in header:
                col_letter = ws.cell(row=1, column=header.index(col_name) + 1).column_letter
                rng = f"{col_letter}2:{col_letter}{last_row}"
                ws.conditional_formatting.add(
                    rng,
                    FormulaRule(
                        formula=[f'AND(${col_letter}2<>"",COUNTIF(${col_letter}$2:${col_letter}${last_row},{col_letter}2)>1)'],
                        fill=red_fill,
                        stopIfTrue=False,
                    ),
                )

        for col_name, (lo, hi) in (range_highlight_columns or {}).items():
            if col_name in header:
                col_letter = ws.cell(row=1, column=header.index(col_name) + 1).column_letter
                rng = f"{col_letter}2:{col_letter}{last_row}"
                first = f"{col_letter}2"
                formula = f'AND({first}<>"",OR({first}<{lo},{first}>{hi}))'
                ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=red_fill, stopIfTrue=False))

        for col_name, days_threshold in (days_until_highlight_columns or {}).items():
            if col_name in header:
                col_letter = ws.cell(row=1, column=header.index(col_name) + 1).column_letter
                rng = f"{col_letter}2:{col_letter}{last_row}"
                first = f"{col_letter}2"
                # ISNUMBER guards against text placeholders (e.g. "Bending
                # for yarn" when the delivery date can't be computed yet)
                # -- those aren't a real date, so they must never be
                # subtracted from TODAY() or Excel shows a #VALUE! error.
                formula = f'AND(ISNUMBER({first}),({first}-TODAY())<{days_threshold})'
                ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=red_fill, stopIfTrue=False))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        letter = col[0].column_letter
        content_width = max((len(_clean(c.value)) for c in col), default=8)
        ws.column_dimensions[letter].width = max(10, content_width + 2)


def export_word(path: Path, template_path: Path, records: list[OrderRecord], stem: str = "") -> None:
    """Create one Word file, one Biglietto page per order color.

    Uses python-docx (lxml-backed) instead of the stdlib ElementTree, and
    deep-copies the template's own table element rather than re-parsing a
    serialized copy. Both changes avoid the namespace-prefix rewriting that
    xml.etree.ElementTree does on round-trip, which is the most likely
    cause of Word's "unreadable content" repair prompt on the old output.
    """
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches

    # Keep the ticket pages in the same numeric Partita Col order as the
    # extracted Excel sheet.  MED source rows are not always delivered in
    # that order, so sorting only the workbook is not enough.
    records = sorted(
        records,
        key=lambda record: (
            _number(record.colored_batch) is None,
            _number(record.colored_batch) if _number(record.colored_batch) is not None else 0,
        ),
    )

    doc = Document(str(template_path))
    if not doc.tables:
        raise ValueError("Il modello Biglietti non contiene la tabella prevista.")

    # The template's own page margins (1in top/bottom) leave the ticket
    # table only ~0.05in of headroom before it spills its last 2-3 rows
    # onto a second page whenever a bilingual label wraps to 2 lines --
    # that's the "biglietto finishes outside its own page" issue. Trimming
    # the margins buys back room without touching row heights or fonts.
    for sect in doc.sections:
        sect.top_margin = Inches(0.5)
        sect.bottom_margin = Inches(0.5)

    template_tbl = doc.tables[0]._tbl
    body = doc.element.body
    template_tbl.getparent().remove(template_tbl)
    sect_pr = body.find(qn("w:sectPr"))

    def _insert(el) -> None:
        if sect_pr is not None:
            sect_pr.addprevious(el)
        else:
            body.append(el)

    for idx, record in enumerate(records):
        new_tbl = copy.deepcopy(template_tbl)
        _fill_ticket_table(new_tbl, record, qn, OxmlElement, stem)
        for row in new_tbl.findall(qn("w:tr")):
            cant_split = OxmlElement("w:cantSplit")
            row.get_or_add_trPr().append(cant_split)
        _insert(new_tbl)
        if idx != len(records) - 1:
            p = OxmlElement("w:p")
            r = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            r.append(br)
            p.append(r)
            _insert(p)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _fill_ticket_table(tbl, r: OrderRecord, qn, OxmlElement, stem: str) -> None:
    rows = tbl.findall(qn("w:tr"))

    def put(row: int, cell: int, text: Any) -> None:
        c = rows[row - 1].findall(qn("w:tc"))[cell - 1]
        for p in c.findall(".//" + qn("w:p")):
            for t in p.findall(".//" + qn("w:t")):
                t.text = ""
            first = p.find(".//" + qn("w:t"))
            if first is None:
                run = OxmlElement("w:r")
                first = OxmlElement("w:t")
                run.append(first)
                p.append(run)
            first.text = _clean(text)
            for run in p.findall(".//" + qn("w:r")):
                r_pr = run.find(qn("w:rPr"))
                if r_pr is None:
                    r_pr = OxmlElement("w:rPr")
                    run.insert(0, r_pr)
                size = r_pr.find(qn("w:sz"))
                if size is None:
                    size = OxmlElement("w:sz")
                    r_pr.append(size)
                size.set(qn("w:val"), "28")
            break

    put(1, 2, _ticket_header(r, stem))
    put(2, 2, r.title); put(2, 3, r.formato); put(2, 4, r.article)
    put(3, 2, r.raw_batch); put(4, 2, r.color_name); put(5, 2, r.colored_batch)
    put(6, 2, r.dispo); put(6, 3, f"Ordine.  {r.order_no}"); put(7, 3, r.bagno)
    put(8, 2, r.quantity_cones); put(9, 2, r.kg); put(10, 2, r.machine)
    vmm = r.vmm22 if r.vmm22 is not None else r.raw_weight
    put(11, 3, f"VMM22( {float(vmm):.2f} )Kg" if vmm is not None else "VMM22(       )Kg")


def _format_date(value: Any) -> str:
    try:
        return value.strftime("%d/%m/%Y")
    except AttributeError:
        return _clean(value)


def _ticket_header(r: OrderRecord, stem: str) -> str:
    """Top line of the ticket. Differs per customer:
    - MED: first 3 letters of Cliente - Commento - Polmon - Cliente MED
      (blank segments dropped rather than left as bare dashes).
    - EL KAMAL: Cliente - Consegna.
    - Everything else (ELVY): Cliente - <order identifier>, unchanged."""
    if r.customer_code == "3004":  # MED
        polmon = _polmoni_segment(r.additional_raw)
        parts = [r.customer_name[:3], r.commento, polmon, r.cliente_med]
        return "-".join(p for p in parts if p)
    if r.customer_code == "3019":  # EL KAMAL
        return f"{r.customer_name} - {_format_date(r.delivery)}"
    return f"{r.customer_name}-{stem}"
