# Performance Report — Planing v1.1.4 startup hardening

Generated: 2026-10-07

## Verification

| Check | Result |
|---|---:|
| Regression suite | **PASS — 191 passed, 5 skipped** |
| Python compile check | **PASS** |
| GUI startup build benchmark | **~0.48 s** in the Linux/Xvfb test environment |
| Lazy-page smoke test | **PASS — 21 pages opened sequentially, 0 callback errors** |

## What was causing startup delay

The previous startup path constructed every major tab inside `_build_ui()`. That
meant heavy parser/export dependency trees (Pandas, OpenPyXL, PDF tooling) were
imported and several hidden tabs restored cached state even though the operator
had not opened those pages yet. `Situazione` also restored its SQLite state during
construction, which could make a large local database block the first paint.

## Changes in v1.1.4

- Heavy pages are now **lazy-built** when first selected: Create/Biglietti,
  Ordine Kamal, Ordine MED, Situazione Settimanale, and Master Data.
- Their heavy imports are also deferred until first real use.
- `Situazione` SQLite startup restore now runs after first paint in a worker;
  SQLite remains the source of truth and no Excel workbook is read for this
  startup restore.
- The existing Magazino/Prezzi cache restores remain asynchronous.
- Bulk synchronization can materialize the heavy consumers it actually needs.
- Existing business logic, exports, calculations, and shared-cache behavior are
  preserved; the change is focused on UI lifecycle/startup scheduling.

## Measured result

The same local startup probe previously measured approximately **1.05 s** for
`_build_ui()`. After the change it measured approximately **0.48 s**, a reduction
of about **54%** before considering machine-specific Windows differences. The
probe is a development benchmark, not a Windows performance guarantee.

## Remaining performance opportunities

1. Replace repeated per-row JSON reconstruction in large SQLite snapshots with
   a more compact normalized representation or a compressed bulk payload where
   appropriate.
2. Move the remaining path-only caches (Articoli, Densita, Magazino, LOTTI,
   Listini/Prezzi, Produzione) toward the same SQLite-first source lifecycle so
   startup and source recovery have one consistent authority.
3. Replace broad shared-refresh fan-out with source-specific invalidation events
   (`dfm_changed`, `magazino_changed`, `prezzi_changed`, etc.).
4. Profile real Windows startup on the production machine after packaging;
   especially measure PyInstaller cold start separately from Python UI build time.
5. Continue replacing DataFrame `iterrows()` in calculation hot paths with
   vectorized operations/groupby/map or `itertuples()` where iteration is only
   needed for rendering/export.

## Acceptance criteria

- First window is painted without constructing every heavy page. **Implemented.**
- Heavy pages still construct successfully when selected. **Implemented.**
- SQLite restore does not block first paint. **Implemented.**
- Existing regression suite remains green. **Implemented.**
- No Excel source is required merely to restore the SQLite-backed Situazione
  state after restart. **Implemented.**
