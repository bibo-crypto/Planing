"""Extract order data into the Italian dyeing-ticket workbook and Word form.

The source workbook is the ERP export (``Sheet1``) plus the manually completed
``Doispo-Bagno`` and optional raw-yarn sheet.  The code deliberately matches
headers by normalized text and accepts either one or two header rows.
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


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


@dataclass
class OrderRecord:
    customer_code: str
    customer_name: str
    article: str
    description: str
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


def _read_sheet_rows(ws) -> list[dict[str, Any]]:
    """Read a sheet using the best header row (row 1 or row 2)."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    candidates = []
    for idx in (0, 1):
        if idx < len(rows):
            headers = [_key(x) or f"__col{i + 1}" for i, x in enumerate(rows[idx])]
            score = sum(bool(x) for x in headers)
            candidates.append((score, idx, headers))
    # Prefer the first row when scores tie; a data row can contain as many
    # non-empty cells as a header row in ERP exports.
    _, header_idx, headers = max(candidates, key=lambda item: (item[0], -item[1]), default=(0, 0, []))
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
    for name in names:
        if _key(name) in row and _clean(row[_key(name)]):
            return row[_key(name)]
    return ""


def load_order(input_path: Path, dispo_path: Path | None = None) -> tuple[list[OrderRecord], list[dict[str, Any]]]:
    wb = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
    if "Sheet1" not in wb.sheetnames:
        raise ValueError('Il file Data Ordine deve contenere il foglio "Sheet1".')
    data = _read_sheet_rows(wb["Sheet1"])
    dispo_rows = _read_sheet_rows(wb["Doispo-Bagno"]) if "Doispo-Bagno" in wb.sheetnames else []
    if dispo_path and dispo_path != input_path:
        dwb = openpyxl.load_workbook(dispo_path, data_only=True, read_only=True)
        dispo_rows = _read_sheet_rows(dwb.active)
        dwb.close()
    raw_rows = _read_sheet_rows(wb["تحضير خيط خام"]) if "تحضير خيط خام" in wb.sheetnames else []
    if not data:
        raise ValueError("Sheet1 non contiene righe d'ordine.")
    dispo = next((r for r in dispo_rows if _clean(_get(r, "Dispo"))), {})
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
        result.append(OrderRecord(
            customer_code=code,
            customer_name=_clean(_get(row, "__col2")) or ("MED" if code == "3004" else "ELVY"),
            article=article,
            description=_clean(_get(row, "Descrizione aggiuntiva ordine")) or _clean(_get(row, "Articolo")),
            color_code=_clean(_get(row, "Colore")),
            color_name=_clean(_get(row, "Colore")) + ("  " + _clean(_get(row, "__col7")) if _clean(_get(row, "__col7")) else ""),
            order_no=_order_number(_get(row, "Ordine")),
            order_row=_clean(_get(row, "Riga")),
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


def _filato_rows(raw_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out = []
    for r in raw_rows:
        if not _clean(_get(r, "Articolo")):
            continue
        out.append({
            "Articolo": _clean(_get(r, "Articolo")),
            "Titolo": _clean(_get(r, "Titolo")),
            "Partita": _clean(_get(r, "Partita.GG", "Partita")),
            "Rocche": _number(_get(r, "عدد", "Rocche")),
            "Peso": _number(_get(r, "وزن", "Peso")),
            "تحضير خام": _clean(_get(r, "Custom", "تحضير خام")) or "تحضير خام",
        })
    return out


def export_workbook(path: Path, records: list[OrderRecord], raw_rows: list[dict[str, Any]], include_filato: bool = True) -> None:
    customer = "ELVY" if records[0].customer_code == "3009" else "MED"
    wb = Workbook()
    ws = wb.active
    ws.title = customer
    common = ["Dispo/Riga", "Articolo", "Titolo", "Formato", "Ordine", "Codice", "Colore", "Rocche", "KG", "M/C", "Partita Col", "Consegna", "Commento", "Bagno", "Cliente", "Partita GG"]
    if customer == "MED":
        headers = common + ["Partita MED", "Cliente MED", "Color Tube", "VMM22", "Prezzo", "Densita` (360-390)"]
    else:
        headers = common + ["Color Tube", "VMM22", "Prezzo", "Densita` (360-390)"]
    ws.append(headers)
    for r in records:
        row = [r.dispo, r.article, r.title, r.formato, r.order_no, r.color_code, r.color_name,
               r.quantity_cones, r.raw_weight, r.machine, r.colored_batch, r.delivery,
               "", r.bagno, r.customer_name, r.raw_batch]
        if customer == "MED":
            row += [_first_match(path.stem, r"PO[- ]*([0-9]+)") or "", _first_match(path.stem, r"-(20[0-9]{2})$"), "", r.raw_weight, "", ""]
        else:
            row += ["", r.raw_weight, "", ""]
        ws.append(row)
    if include_filato:
        fws = wb.create_sheet("Filato x Tinturia")
        fheaders = ["Articolo", "Titolo", "Partita", "Rocche", "Peso", "تحضير خام"]
        fws.append(fheaders)
        for r in _filato_rows(raw_rows):
            fws.append([r[h] for h in fheaders])
        _style_sheet(fws)
    _style_sheet(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def export_filato_workbook(path: Path, raw_rows: list[dict[str, Any]]) -> None:
    """Export only the optional ``Filato x Tinturia`` workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Filato x Tinturia"
    headers = ["Articolo", "Titolo", "Partita", "Rocche", "Peso", "تحضير خام"]
    ws.append(headers)
    for r in _filato_rows(raw_rows):
        ws.append([r[h] for h in headers])
    _style_sheet(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def _style_sheet(ws) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = min(28, max(12, max(len(_clean(c.value)) for c in col) + 2))


def export_word(path: Path, template_path: Path, records: list[OrderRecord]) -> None:
    """Create one Word file, one Biglietto page per order color."""
    with ZipFile(template_path) as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tbl = root.find(".//w:tbl", ns)
        if tbl is None:
            raise ValueError("Il modello Biglietti non contiene la tabella prevista.")
        template_table = ET.tostring(tbl, encoding="utf-8")
        body = root.find("w:body", ns)
        sect = body.find("w:sectPr", ns)
        body.remove(tbl)
        for idx, record in enumerate(records):
            page_tbl = ET.fromstring(template_table)
            _fill_ticket_table(page_tbl, record, ns, template_path.stem)
            body.insert(len(body) - (1 if sect is not None else 0), page_tbl)
            if idx != len(records) - 1:
                p = ET.Element("{%s}p" % ns["w"])
                r = ET.SubElement(p, "{%s}r" % ns["w"])
                br = ET.SubElement(r, "{%s}br" % ns["w"]); br.set("{%s}type" % ns["w"], "page")
                body.insert(len(body) - (1 if sect is not None else 0), p)
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        path.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(path, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, new_xml if item.filename == "word/document.xml" else zin.read(item.filename))


def _fill_ticket_table(tbl, r: OrderRecord, ns: dict[str, str], stem: str) -> None:
    rows = tbl.findall("./w:tr", ns)
    def put(row: int, cell: int, text: Any) -> None:
        c = rows[row - 1].findall("./w:tc", ns)[cell - 1]
        for p in c.findall(".//w:p", ns):
            for t in p.findall(".//w:t", ns):
                t.text = ""
            first = p.find(".//w:t", ns)
            if first is None:
                run = ET.SubElement(p, "{%s}r" % ns["w"]); first = ET.SubElement(run, "{%s}t" % ns["w"])
            first.text = _clean(text)
            break
    put(1, 2, f"{r.customer_name}-{stem}")
    put(2, 2, r.title); put(2, 3, r.formato); put(2, 4, r.article)
    put(3, 2, r.raw_batch); put(4, 2, r.color_name); put(5, 2, r.colored_batch)
    put(6, 2, r.dispo); put(6, 3, f"Ordine.  {r.order_no}"); put(7, 3, r.bagno)
    put(8, 2, r.quantity_cones); put(9, 2, r.raw_weight); put(10, 2, r.machine)
    put(11, 3, f"VMM22( {_clean(r.raw_weight)} )Kg" if r.raw_weight else "VMM22(       )Kg")
