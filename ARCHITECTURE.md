# Architecture

A map of what lives where. The app is a single Tkinter desktop app; most
files sit flat at the repo root (Python doesn't need folders to keep
modules organized, and a flat layout keeps every import a plain
`import module_name` with no package-path churn if a file moves between
folders later).

## UI (`ui/`)
- `ui/gui.py` — the main window: builds the notebook of tabs, wires
  cross-tab callbacks (e.g. `_on_shared_cache_changed`), owns
  `self._prefs`/`self._save_prefs` (the persisted-settings dict every tab
  reads/writes through).
- `ui/tabs/*.py` — one file per tab. Each tab owns its own widgets and
  its own upload/export handlers; business logic it needs lives in a
  same-named `*_logic.py` at the root, not inside the tab file.

## Business logic (`*_logic.py`, root)
Pure(ish) computation, no Tkinter: `situazione_logic.py` (Copertura,
machine-queue scheduling, `compute_situation`), `magazino_logic.py`,
`prezzi_logic.py`, `lotti_logic.py`. These take dataframes/paths in,
return dataframes/values out — safe to unit-test without a display.

## Loaders / parsers (root)
Read a specific file shape into a dataframe or a list of dataclass rows:
`situazione_loaders.py` (DFM/Copertura/Produzione/Wincoint/Uscita/
Qualita/Articoli), `pdf_parser.py` (the Elvy PO PDF), `dfm_lookup.py`,
`prod_lookup.py`.

## Customer-specific pipelines (root)
Each customer's "raw order -> ticket/workbook" flow is one module:
- `biglietti_exporter.py` — ELVY / MED / EL KAMAL dyeing tickets
  (Biglietti) + their Excel workbook. `load_order`/`load_el_kamal_order`
  parse, `enrich_records` fills in Titolo/M-C/KG/Prezzo/etc.,
  `export_workbook`/`export_word` write the output.
- `ordine_med.py` — the "Ordine da creare" / Filato-availability
  extraction (separate from Biglietti; different source shape).
- `ordini_elvy.py` — the ERP-import "Ordini ELVY" sheet built from PDF
  orders.
- `kamal_parser.py` / `kamal_excel_exporter.py` — the older
  Kamal-specific PDF pipeline (predates `biglietti_exporter.py`'s own
  EL KAMAL support; kept for its own tab, not merged in).

Each pipeline module owns its own field-normalization helpers
(`_clean`/`_key`/`_number`/`_read_sheet_rows`) rather than sharing one
"utils" grab-bag — they're one-liners, and duplicating them keeps each
module's logic self-contained.

## Centralized source paths (`path_manager.py` + `file_cache.py`)
`path_manager.py` is the canonical registry for every shared source name used
by the pages and by Overview's bulk import. It normalizes aliases such as
`data_prod`/`produzione` and `listini`/`prezzi`, while `file_cache.py` remains
the single JSON persistence implementation. Individual `*_cache.py` modules
are compatibility wrappers, so uploads from any page resolve to the same
stored path and survive application restarts. Data Ordine and Dispo-Bagno are
also recorded centrally for Biglietti and Ordine MED.

## "Where was that file last uploaded" caches (root, `*_cache.py`)
One JSON file per source (`settings/<key>_file_cache.json`) remembering
the last path used, so re-opening the app or switching tabs doesn't
require re-browsing. All of them are thin wrappers around
`file_cache.py`'s `save_file_cache(key, path)` / `load_file_cache(key)`
— the actual read/write logic lives in exactly one place. Add a new
cached source by adding a new few-line wrapper in the same shape, not by
extending `file_cache.py` itself. Note: Ordine MED's raw-yarn
availability check reuses `magazino_cache.py` (the same shared Magazino
Filato source as Ordine Elvy/Ordine Kamal/Situazione/Magazino Filato) —
there is no separate "Filato Disponibile" cache; that file shape is the
same as Magazino's own export.

## Cross-tab / bulk import (root)
- `master_import.py` — "load everything from one folder/file" for
  Overview: matches files by name (falling back to content-sniffing for
  Data Ordine/Dispo-Bagno, which have no stable filename), routes each to
  the same handler its own tab's upload button uses.
- `situazione_db.py` — the one piece of actual persistence beyond
  per-source path caches: SQLite-backed upload log + the shared
  Articolo->TITOLO codes table.

`Ordine MED` selects an output directory (rather than a file) and writes
`Ordine_MED.xlsx`. Its ERP panel follows the Ordine Kamal pattern: separate
ERP and Filato folder rows, each with its own extraction checkbox. The ERP
checkbox writes `Ordine_MED_ERP.xlsx`; the Filato checkbox writes a standalone
`Filato X Tinturia.xlsx`; the combined workbook remains unchanged.

`Situazione Generale` also runs a background synchronization pass for every
shared source (Copertura, WINCOINT, Uscita, Qualita, Articoli, Listini, DFM,
and Produzione). A successful upload from another page invokes the same pass,
so the file is read into the correct page state instead of merely displaying a
saved filename.

## Packaging
`main.spec` (PyInstaller) + `requirements.txt` + `installer.iss` (Inno
Setup) + `build.bat`. When a module gains a new third-party import, add
it to `requirements.txt` and, if PyInstaller's static analysis won't
find it on its own (dynamic imports, C extensions), to
`extra_hiddenimports` in `main.spec` too.
