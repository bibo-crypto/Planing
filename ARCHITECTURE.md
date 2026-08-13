# Planing (Delta Dyeing Converter) – Project Structure

The project is organized by responsibility. The small root-level facade files
are intentionally kept for backward compatibility with older imports and the
existing PyInstaller entry point.

## Packages

- `ui/` – main application window and shared UI widgets.
- `ui/tabs/` – one module per notebook tab. Each tab owns its controls,
  events, and presentation logic.
- `parsers/` – PDF and source-file parsing logic.
- `exporters/` – Excel workbook generation.
- `domain/` – business rules and row-building/matching logic.
- `data/` – caches, SQLite persistence, and source-file loaders.

## Maintenance rules

1. Keep UI event handlers inside the relevant tab module.
2. Keep parsing free of Tkinter code.
3. Keep business calculations free of file dialogs and widgets.
4. Keep Excel formatting in exporter modules.
5. Add a short English module docstring and comments around non-obvious rules.
6. Run `python -m compileall .` before building the executable.

The root facades can be removed only after all external scripts and deployment
configurations have migrated to the package imports.

