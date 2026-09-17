# Architecture

A map of what lives where. The app is a single Tkinter desktop app. Modules
are grouped into folders by role, so a new contributor (or future-you) can
tell what a file does from its path alone, and every import is a plain
`from <folder>.<module> import ...` with no root-level clutter.

## Layout

```
main.py                 entry point

gui/                     Tkinter UI: windows, tabs, widgets
  gui.py                 main window, notebook, cross-tab wiring
  modern_widgets.py       shared custom widgets (RoundedButton, ...)
  tabs/                   one file per notebook tab

calculate/               pure business logic, no Tkinter
  situazione.py           Copertura / machine-queue / compute_situation
  situazione_settimana.py
  magazino.py, lotti.py, prezzi.py
  abbina_calculator.py, abbina_suggestions.py
  constants.py            shared machine capacity/code tables

parsers/                 read a specific file shape into a dataframe/rows
  pdf_parser.py, bolla_parser.py, elvy_invoice_parser.py, kamal_parser.py
  dfm_lookup.py, prod_lookup.py, situazione_loaders.py, elvy_mapping.py

exporters/               write Excel/Word output
  excel_exporter.py, bolla_exporter.py, elvy_invoice_exporter.py
  kamal_excel_exporter.py, biglietti_exporter.py

pipelines/               customer-specific "raw order -> ticket/workbook" flows
  ordine_kamal.py, ordine_med.py, ordini_elvy.py
  master_import.py        "load everything from one folder" for Overview

utility/                 shared infra: generic helpers, path/file caches, db
  utils.py, path_manager.py, file_cache.py
  articoli_cache.py, densita_cache.py, lotti_cache.py, magazino_cache.py,
  prezzi_cache.py
  situazione_db.py        SQLite-backed upload log + Articolo->TITOLO table
  notifications.py        persisted, deduplicated application notifications
  master_data.py          local JSON reference data (temporary, not central DB)

tests/                   regression tests (pytest/unittest)
```

## Notifications

The main window has a global `Notifications` button. Notices are stored in
`settings/notifications.json`, deduplicated by key, and remain visible until
the user marks them resolved. The notification window shows severity, title,
page, message, and creation time.

`Ordine > Ordine Med` creates a high-severity notice after conversion when an
article required by `Filato X Tinturia` is missing from `Magazino Filato`.
The notice lists the missing article codes so the operator can upload or fix
the Magazino source. `Situazione Generale` also compares the displayed color
price against `Prezzi` using `Articolo + Codice`; missing matches and existing
price differences become notifications. The machine surcharge rule is applied
for both capacity values `24/32/56` and machine numbers `12/9/10`: add $2 to
the Listini price.

## Master Data (local phase)

The `Master Data` tab intentionally contains only the two manually maintained
tables that are useful globally: `Customers` and `Machines`. Machines use the
fields `Code`, `Number`, and `Rocche`; for example `3307 - 7 - 72`. The seeded
machine registry is loaded by `calculate/constants.py` and supplies the
machine-code, machine-number, and capacity mappings used by planning logic.
Article files remain managed through their dedicated upload flow, and the
G/C article-prefix rule remains business logic rather than a separate table.
Changes are saved in `settings/master_data.json` and are picked up on the
next application start.

## Central database (postponed)

A central business database is deliberately not part of this phase. Existing
source-path caches, the Situazione upload log, and the new local JSON stores
remain in place. A future database phase can add migrations, permissions,
history, and shared records after the operating rules and backup strategy are
approved.

Packaging/build tooling stays at the repo root since it isn't application
code: `main.spec`, `build.bat`, `installer.iss`, `requirements.txt`,
`runtime_hook.py`, `sync_venv_packages.py`, `icon.ico`, plus the two
standalone smoke-test scripts `test_startup.py` / `test_update.py` (these
are manual scripts, not part of the `tests/` pytest package).

### Import rule
New code imports from these packages directly — `from calculate.situazione
import ...`, `from utility.utils import ...`, etc. There are no more
root-level `*_tab.py` / `*_logic.py` compatibility shims: the folder reorg
finished the migration that this file previously described as in-progress,
so every caller now points straight at the canonical module. If you're
looking at an old note/branch that still says `import situazione_logic` or
`import ui.tabs...`, update it to `import calculate.situazione` /
`import gui.tabs...`.

### Maintenance map
- `gui/tabs/` — Tkinter widgets, event handlers, and rendering only.
- `calculate/*.py` — pure dataframe/business rules and calculations.
- `parsers/*.py` — input normalization and validation.
- `exporters/*.py` — Excel, Word, and PDF output.
- `utility/*_cache.py`, `utility/file_cache.py`, `utility/path_manager.py` —
  persisted source paths and shared cache state.
- `pipelines/master_import.py` — orchestration boundary for folder/master-file
  imports; it may call tab adapters but should not contain business
  calculations.

## Business logic (`calculate/`)
Pure(ish) computation, no Tkinter: `calculate/situazione.py` (Copertura,
machine-queue scheduling, `compute_situation`), `calculate/magazino.py`,
`calculate/prezzi.py`, `calculate/lotti.py`. These take dataframes/paths in,
return dataframes/values out — safe to unit-test without a display.

## Loaders / parsers (`parsers/`)
Read a specific file shape into a dataframe or a list of dataclass rows:
`situazione_loaders.py` (DFM/Copertura/Produzione/Wincoint/Uscita/
Qualita/Articoli), `pdf_parser.py` (the Elvy PO PDF), `dfm_lookup.py`,
`prod_lookup.py`.

## Customer-specific pipelines (`pipelines/`)
Each customer's "raw order -> ticket/workbook" flow is one module:
- `exporters/biglietti_exporter.py` — ELVY / MED / EL KAMAL dyeing tickets
  (Biglietti) + their Excel workbook. `load_order`/`load_el_kamal_order`
  parse, `enrich_records` fills in Titolo/M-C/KG/Prezzo/etc.,
  `export_workbook`/`export_word` write the output.
- `pipelines/ordine_med.py` — the "Ordine da creare" / Filato-availability
  extraction (separate from Biglietti; different source shape).
- `pipelines/ordini_elvy.py` — the ERP-import "Ordini ELVY" sheet built from
  PDF orders.
- `pipelines/ordine_kamal.py` / `exporters/kamal_excel_exporter.py` — the
  older Kamal-specific PDF pipeline (predates `biglietti_exporter.py`'s own
  EL KAMAL support; kept for its own tab, not merged in).

Each pipeline module owns its own field-normalization helpers
(`_clean`/`_key`/`_number`/`_read_sheet_rows`) rather than sharing one
"utils" grab-bag — they're one-liners, and duplicating them keeps each
module's logic self-contained.

## Centralized source paths (`utility/path_manager.py` + `utility/file_cache.py`)
`path_manager.py` is the canonical registry for every shared source name used
by the pages and by Overview's bulk import. It normalizes aliases such as
`data_prod`/`produzione` and `listini`/`prezzi`, while `file_cache.py` remains
the single JSON persistence implementation. Individual `*_cache.py` modules
are thin wrappers, so uploads from any page resolve to the same stored path
and survive application restarts. Data Ordine and Dispo-Bagno are also
recorded centrally for Biglietti and Ordine MED.

## "Where was that file last uploaded" caches (`utility/*_cache.py`)
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

## Cross-tab / bulk import
- `pipelines/master_import.py` — "load everything from one folder/file" for
  Overview: matches files by name (falling back to content-sniffing for
  Data Ordine/Dispo-Bagno, which have no stable filename), routes each to
  the same handler its own tab's upload button uses.
- `utility/situazione_db.py` — the one piece of actual persistence beyond
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

## Dev-mode data location (`utility/utils.py`)
`APP_DATA_DIR` decides where `logs/`, `settings/`, and the Situazione SQLite
file live. When frozen (PyInstaller build) it uses the per-user AppData
folder; when running from source it uses "the folder next to this file" —
which, since `utils.py` sits one level under the project root now
(`utility/utils.py`), is computed as `Path(__file__).resolve().parent.parent`
so dev-mode data keeps landing at the repo root, exactly as it did before the
folder reorg. If `utility/utils.py` ever moves again, update that parent
count to match its new depth.

## Packaging
`main.spec` (PyInstaller) + `requirements.txt` + `installer.iss` (Inno
Setup) + `build.bat`. `pathex=['.']` plus PyInstaller's import-following
analysis means the `gui/`, `calculate/`, `parsers/`, `exporters/`,
`pipelines/`, `utility/` packages need no extra spec changes — PyInstaller
discovers them the same way it always discovered `ui/`/`logic/`. When a
module gains a new third-party import, add it to `requirements.txt` and, if
PyInstaller's static analysis won't find it on its own (dynamic imports, C
extensions), to `extra_hiddenimports` in `main.spec` too. `build.bat` reuses
an existing `venv` in place (via `sync_venv_packages.py`, which removes
anything no longer in `requirements.txt`) instead of deleting and recreating
it on every build.

## Number/text parsing (deliberately NOT consolidated)
`biglietti_exporter._clean`/`_number`, `utils.clean_text`/`parse_number`,
`abbina_suggestions._number`, and `elvy_invoice_parser._parse_number`
(which already just wraps `utils.parse_number`) look like duplicates at
a glance, but they aren't behaviorally identical: `utils.parse_number`
disambiguates thousands-vs-decimal separators for the general case
("1,234.56" vs "1.234,56"), while `biglietti_exporter._number` and
`abbina_suggestions._number` take the simpler, deliberate assumption
that a comma is always a decimal separator (correct for the real
Italian/Egyptian ERP exports Biglietti/Ordine MED parse, which use
comma-decimals without thousands separators) and differ from each other
in their None-vs-0.0 failure return. Merging any of these would risk
silently changing parsed values across pipelines that are already
verified against real data -- if you're looking at this thinking "these
should be one function," they were considered and kept apart on purpose.
