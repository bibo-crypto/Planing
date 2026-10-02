# Performance Report — Planing v1.1.1 production update

Generated: 2026-10-01 18:27:33

## Verification benchmark

| Check | Result | Duration |
|---|---:|---:|
| Regression suite | PASS (69 passed in 2.38s) | 3.17s |

## Changes measured / reasoned about

- **SQLite startup snapshot:** `situazione_db.save_frame_cache()` now writes a versioned `source_snapshot` metadata row and row-level `source_snapshot_row` records. Reads use the normalized snapshot first and retain the old `frame_cache` JSON table as a migration fallback.
- **SQLite concurrency:** connections use WAL mode, `synchronous=NORMAL`, foreign keys, and row-level snapshot writes are transactional.
- **Prezzi export:** normal Excel export was already off the Tk thread; Price Changes calculation and its Excel export are now also background operations.
- **DFM startup:** DFM is no longer restored by shared-cache or Situation startup paths. It is loaded only after explicit upload.
- **ESC lifecycle:** the main application now has one global Escape handler that closes ordinary child windows and removes stale child registrations. The update window is protected while a download is in flight.
- **Category rule:** Prezzi detects `CUSTOMER/CLIENTE` and `MARCA/BRAND` aliases and constructs `Customer - Marca`; article/category fallback remains for legacy exports.

## Existing profile hotspots

### `startup.prof`

| Cumulative seconds | Internal seconds | Function |
|---:|---:|---|
| 58.352 | 0.000 | `E:\Planing\ui\gui.py:130 — __init__` |
| 57.813 | 0.001 | `E:\Planing\ui\gui.py:215 — _build_ui` |
| 49.637 | 0.194 | `E:\Planing\venv\Lib\site-packages\openpyxl\worksheet\_read_only.py:60 — _cells_by_row` |
| 48.048 | 0.010 | `E:\Planing\biglietti_exporter.py:278 — load_prezzo_lookup` |
| 27.482 | 0.946 | `E:\Planing\logic\prezzi.py:66 — build_price_lookup` |
| 27.439 | 0.008 | `E:\Planing\ui\tabs\biglietti_tab.py:33 — __init__` |
| 27.139 | 0.003 | `E:\Planing\ui\tabs\biglietti_tab.py:116 — _restore` |
| 26.059 | 0.000 | `E:\Planing\ui\tabs\situazione_tab.py:143 — __init__` |
| 25.891 | 0.007 | `E:\Planing\ui\tabs\situazione_tab.py:1184 — _load_table_from_db` |
| 25.497 | 0.004 | `E:\Planing\ui\tabs\situazione_tab.py:1144 — _recompute_prezzo_densita` |
| 20.553 | 0.023 | `E:\Planing\logic\prezzi.py:39 — load_prezzi` |
| 20.308 | 0.598 | `E:\Planing\venv\Lib\site-packages\openpyxl\worksheet\_reader.py:125 — parse` |

### `startup_after.prof`

| Cumulative seconds | Internal seconds | Function |
|---:|---:|---|
| 43.119 | 0.201 | `E:\Planing\venv\Lib\site-packages\openpyxl\worksheet\_read_only.py:60 — _cells_by_row` |
| 15.957 | 0.618 | `E:\Planing\venv\Lib\site-packages\openpyxl\worksheet\_reader.py:125 — parse` |
| 15.553 | 0.000 | `E:\Planing\ui\gui.py:130 — __init__` |
| 15.152 | 0.001 | `E:\Planing\ui\gui.py:215 — _build_ui` |
| 12.491 | 0.350 | `C:\Users\Mega Store\AppData\Local\Programs\Python\Python314\Lib\xml\etree\ElementTree.py:1243 — iterator` |
| 9.899 | 0.000 | `E:\Planing\ui\tabs\biglietti_tab.py:33 — __init__` |
| 9.657 | 0.003 | `E:\Planing\ui\tabs\biglietti_tab.py:116 — _restore` |
| 9.331 | 0.000 | `test_startup.py:7 — main` |
| 9.331 | 0.002 | `C:\Users\Mega Store\AppData\Local\Programs\Python\Python314\Lib\tkinter\__init__.py:1611 — mainloop` |
| 8.029 | 0.003 | `E:\Planing\biglietti_exporter.py:281 — load_prezzo_lookup` |
| 7.948 | 0.001 | `E:\Planing\venv\Lib\site-packages\pandas\io\excel\_base.py:451 — read_excel` |
| 7.814 | 0.009 | `E:\Planing\logic\prezzi.py:39 — load_prezzi` |

## Acceptance criteria status

| Requirement | Status | Notes |
|---|---|---|
| Category = Customer-Marca | **Implemented** | Robust aliases plus legacy fallback |
| Price Changes / Export do not freeze UI | **Implemented** | Worker threads; Tk updates stay on main thread |
| DFM explicit upload only | **Implemented** | Startup/shared sync skips DFM |
| Global ESC closes child windows | **Implemented** | Update dialog is intentionally protected |
| SQLite snapshots | **Implemented** | Versioned normalized row-level snapshot plus compatibility fallback |
| Lazy heavy tabs | **Partial / existing staged loading** | Current app defers source restore and renders in chunks; full constructor-level tab factory remains a follow-up because Overview and cross-tab wiring currently require live tab objects |

## Reproduction

```bash
cd /home/ubuntu/work/Planing-final
python3 -m pytest -q
python3 -m compileall -q .
```
