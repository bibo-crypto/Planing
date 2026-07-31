"""
gui.py — Tkinter user interface for the Delta Dyeing PDF-to-Excel converter.

Layout
------
┌─────────────────────────────────────────────────┐
│  [Purchase Orders] [Bolla]   <- tabs             │
├─────────────────────────────────────────────────│
│   (tab content — see _build_po_tab /             │
│    _build_bolla_tab)                             │
├─────────────────────────────────────────────────│
│  Log window (scrollable, shared by both tabs)    │
└─────────────────────────────────────────────────┘

Purchase Orders tab
    Input (PDF/Folder/Output) -> Abbina Mode -> Options+Convert -> Progress

Bolla tab
    Input (PDF/Folder/Output) -> Options+Convert -> Progress
    Parses Italian delivery-note ("Bolla") PDFs via bolla_parser.py and
    exports an "Items" + "Totals" workbook via bolla_exporter.py.
"""

from __future__ import annotations

import logging
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from abbina_calculator import AbbinaCalculator
from pdf_parser import PDFParser, OrderRow
from excel_exporter import ExcelExporter
from ordini_elvy import build_ordini_elvy_rows, update_existing_ordini_file
from bolla_parser import BollaParser, BollaRow, BollaTotals
from bolla_exporter import BollaExporter
from elvy_invoice_parser import ElvyInvoiceParser, ElvyInvoiceRow
from elvy_invoice_exporter import ElvyInvoiceExporter
from elvy_mapping import (
    add_elvy_mapping,
    delete_elvy_mapping,
    load_elvy_mapping,
    lookup_articolo_delta,
)
from dfm_lookup import (
    build_dfm_lookup,
    load_dfm_cache,
    lookup_dfm_color,
    save_dfm_cache,
)
from utils import find_pdfs, load_settings, logger, make_output_path, save_settings


def _resource_path(filename: str) -> Path:
    """
    Resolve the path to a bundled resource (e.g. icon.ico) so it works both
    when running from source and when frozen by PyInstaller.

    A frozen --onedir build extracts/keeps bundled data files next to the
    .exe under ``sys._MEIPASS`` — a plain ``Path(__file__).parent`` lookup
    (which works fine in dev mode) resolves to the wrong place once frozen,
    since gui.py itself no longer lives on disk as a real file at runtime.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base / filename


# ---------------------------------------------------------------------------
# Log handler that forwards records to the GUI log window via a thread-safe queue
# ---------------------------------------------------------------------------

class _QueueHandler(logging.Handler):
    """Push log records into a :class:`queue.Queue` for GUI consumption."""

    def __init__(self, log_queue: "queue.Queue[str]") -> None:
        super().__init__()
        self._queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self._queue.put_nowait(msg)
        except Exception:  # noqa: BLE001
            self.handleError(record)


# ---------------------------------------------------------------------------
# Main Application Window
# ---------------------------------------------------------------------------

class ConverterApp(tk.Tk):
    """
    Root Tk window.  All UI widgets live here.
    Business logic is dispatched to background threads to keep the UI responsive.
    """

    WINDOW_TITLE = "Delta Dyeing PDF to Excel Converter"
    WINDOW_MIN_W = 1000
    WINDOW_MIN_H = 760

    def __init__(self) -> None:
        super().__init__()
        self.title(self.WINDOW_TITLE)
        self.minsize(self.WINDOW_MIN_W, self.WINDOW_MIN_H)
        self.resizable(True, True)
        self._set_window_icon()

        # ── Persisted state ───────────────────────────────────────────
        self._prefs = load_settings()

        # ── Purchase Orders tab state ─────────────────────────────────
        self._po_pdf_path: Path | None = None
        self._po_folder_path: Path | None = None
        self._po_output_dir: Path | None = None
        self._po_one_per_file = tk.BooleanVar(value=False)
        self._po_update_erp_file = tk.BooleanVar(value=False)
        self._po_update_erp_file.trace_add(
            "write",
            lambda *_a: self._save_prefs(po_update_erp_file=self._po_update_erp_file.get()),
        )
        self._po_erp_file_path: Path | None = None

        # ── Bolla tab state ────────────────────────────────────────────
        self._bolla_pdf_path: Path | None = None
        self._bolla_folder_path: Path | None = None
        self._bolla_output_dir: Path | None = None
        self._bolla_one_per_file = tk.BooleanVar(value=False)

        # ── Elvy Invoice tab state ──────────────────────────────────────
        self._einv_pdf_path: Path | None = None
        self._einv_folder_path: Path | None = None
        self._einv_output_dir: Path | None = None
        self._einv_one_per_file = tk.BooleanVar(value=False)

        # Thread-safe log queue (shared by both tabs)
        self._log_queue: "queue.Queue[str]" = queue.Queue()

        self._build_ui()
        self._attach_log_handler()
        self._poll_log_queue()

    # ------------------------------------------------------------------
    # Window / taskbar icon
    # ------------------------------------------------------------------

    def _set_window_icon(self) -> None:
        """
        Set the titlebar/taskbar icon from icon.ico.

        This is separate from the .exe's own icon (set via main.spec) —
        that only controls how the .exe file looks in Explorer/shortcuts.
        Without this call, the running window falls back to Tk's default
        feather icon regardless of what icon the .exe file has.
        """
        icon_path = _resource_path("icon.ico")
        if not icon_path.exists():
            logger.warning("icon.ico not found at %s — using default window icon", icon_path)
            return
        try:
            self.iconbitmap(str(icon_path))
        except tk.TclError as exc:
            logger.warning("Could not set window icon: %s", exc)

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Build all widgets."""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=3)   # all application pages
        self.rowconfigure(1, weight=1)   # log area

        style = ttk.Style(self)
        # Distinct background for whichever tab is currently selected, so
        # it's obvious at a glance which tab is open. Only the background
        # is overridden — leaving foreground at its theme default avoids a
        # bug where forcing white text made the label invisible against
        # some themes' selected-tab rendering.
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#1976D2"), ("!selected", "#D9E4EC")],
            foreground=[("selected", "#000000"), ("!selected", "#1F2937")],
        )

        # ── All pages use one normal tab bar; Data Elvy is not persistent ──
        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 4))

        elvy_tab = ttk.Frame(notebook)
        po_tab = ttk.Frame(notebook)
        bolla_tab = ttk.Frame(notebook)
        elvy_invoice_tab = ttk.Frame(notebook)
        notebook.add(elvy_tab, text="Data Elvy")
        notebook.add(po_tab, text="Ordini ELVY")
        notebook.add(bolla_tab, text="Med Bolla")
        notebook.add(elvy_invoice_tab, text="Elvy Invoice")

        self._build_elvy_tab(elvy_tab)
        self._build_po_tab(po_tab)
        self._build_bolla_tab(bolla_tab)
        self._build_elvy_invoice_tab(elvy_invoice_tab)
        self._build_log_area()
        self._restore_saved_paths()

    # ------------------------------------------------------------------
    # Purchase Orders tab
    # ------------------------------------------------------------------

    def _build_po_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=8)
        sel_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 4))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._po_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_po_select_pdf, width=20
        )
        self._po_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._po_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_po_select_folder, width=20
        )
        self._po_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._po_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_po_select_output, width=20
        )
        self._po_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._po_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._po_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)

        self._po_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)

        self._po_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4)

        abbina_info = ttk.Label(
            parent,
            text="Abbina is taken from the PDF's own Machine annotation when present; "
                 "any row left without one is filled in automatically (smallest fitting "
                 "machine for its same-colour group).",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        abbina_info.grid(row=1, column=0, sticky="ew", padx=4, pady=(6, 0))

        # ── Update existing ERP file ─────────────────────────────────
        erp_frame = ttk.LabelFrame(parent, text="Also Update Existing ERP File", padding=8)
        erp_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(6, 4))
        erp_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            erp_frame,
            text="After converting, write the Ordini ELVY rows into this existing file "
                 "too — every row below the header is cleared first, then replaced",
            variable=self._po_update_erp_file,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Button(
            erp_frame, text="📄 Select ERP File…", command=self._on_po_select_erp_file, width=20
        ).grid(row=1, column=0, padx=(0, 8), pady=(4, 0), sticky="w")

        self._po_lbl_erp_file = ttk.Label(
            erp_frame, text="No file selected", foreground="grey", anchor="w"
        )
        self._po_lbl_erp_file.grid(row=1, column=1, sticky="ew", pady=(4, 0))

        erp_warning = ttk.Label(
            erp_frame,
            text="⚠ Close this file in Excel before converting, or saving will fail.",
            foreground="#8a6d00", anchor="w",
        )
        erp_warning.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(8, 4))
        opt_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=2)
        opt_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            opt_frame,
            text="One Excel file per PDF  (otherwise merge all into one workbook)",
            variable=self._po_one_per_file,
        ).grid(row=0, column=0, sticky="w")

        self._po_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_po_convert,
            style="Accent.TButton",
            width=14,
        )
        self._po_btn_convert.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(8, 4))
        prog_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=2)
        prog_frame.columnconfigure(0, weight=1)

        self._po_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._po_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._po_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._po_lbl_status.grid(row=1, column=0, sticky="ew")

    # ------------------------------------------------------------------
    # Bolla tab
    # ------------------------------------------------------------------

    def _build_bolla_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        info = ttk.Label(
            parent,
            text="Extracts Bolla No / Del, and every Pallet row (Scatola, Disposizione, "
                 "Articolo, Descrizione, Colore, Partita, Rocche, KgNetto, KgLordo, "
                 "Famiglia) plus the Totale summary line.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        info.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=8)
        sel_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 4))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._bolla_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_bolla_select_pdf, width=20
        )
        self._bolla_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._bolla_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_bolla_select_folder, width=20
        )
        self._bolla_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._bolla_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_bolla_select_output, width=20
        )
        self._bolla_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._bolla_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)

        self._bolla_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)

        self._bolla_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4)

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(8, 4))
        opt_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=2)
        opt_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            opt_frame,
            text="One Excel file per Bolla  (otherwise merge all into one workbook)",
            variable=self._bolla_one_per_file,
        ).grid(row=0, column=0, sticky="w")

        self._bolla_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_bolla_convert,
            style="Accent.TButton",
            width=14,
        )
        self._bolla_btn_convert.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(8, 4))
        prog_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=2)
        prog_frame.columnconfigure(0, weight=1)

        self._bolla_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._bolla_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._bolla_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._bolla_lbl_status.grid(row=1, column=0, sticky="ew")

    # ------------------------------------------------------------------
    # Elvy tab — Article No (Elvy) -> Articolo Delta mapping
    # ------------------------------------------------------------------

    def _build_elvy_tab(self, parent: ttk.Frame) -> None:
        """
        This tab manages data specific to Elvy orders only: a lookup table
        matching Elvy's own "Article No" (the same value already extracted
        as Article No on the Purchase Orders tab) to the equivalent
        internal Delta article code. It does not affect Bolla data at all.

        Every time Purchase Orders are converted, each row's Article No is
        looked up here and the result is written into a new "Articolo
        Delta" column — like a merge/VLOOKUP by Article No. Rows with no
        matching entry are left blank in that column.
        """
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        info = ttk.Label(
            parent,
            text="Elvy-specific data. Map each Elvy Article No to its equivalent "
                 "Delta article code here. When converting Purchase Orders, every "
                 "row's Article No is looked up in this table and the match is "
                 "written into a new \"Articolo Delta\" column — rows with no "
                 "match are left blank.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        info.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))

        # ── Add / update entry ──────────────────────────────────────────
        entry_frame = ttk.LabelFrame(parent, text="Add / Update Mapping", padding=8)
        entry_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 4))
        entry_frame.columnconfigure(1, weight=1)
        entry_frame.columnconfigure(3, weight=1)

        ttk.Label(entry_frame, text="Article No (Articolo Elvy):").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self._elvy_entry_elvy = ttk.Entry(entry_frame)
        self._elvy_entry_elvy.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=4)

        ttk.Label(entry_frame, text="Articolo Delta:").grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=4
        )
        self._elvy_entry_delta = ttk.Entry(entry_frame)
        self._elvy_entry_delta.grid(row=0, column=3, sticky="ew", padx=(0, 12), pady=4)

        ttk.Button(
            entry_frame, text="💾  Save", command=self._on_elvy_save, width=10
        ).grid(row=0, column=4, sticky="e")

        # ── DFM colour reference ─────────────────────────────────────────
        dfm_frame = ttk.LabelFrame(parent, text="DFM Color Reference", padding=8)
        dfm_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 4))
        dfm_frame.columnconfigure(1, weight=1)

        dfm_info = ttk.Label(
            dfm_frame,
            text="Load the DFM.xlsx export (filtered to Elvy / C130 articles) so "
                 "Purchase Order conversions can look up each row's colour: its "
                 "Article No + Colore is matched against this data to fill in two "
                 "new columns, COLOREDFM and CLDESCR (Delta's colour code and "
                 "name). Reload this whenever you get an updated DFM export.",
            foreground="grey", anchor="w", wraplength=700, justify="left",
        )
        dfm_info.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        ttk.Button(
            dfm_frame,
            text="📂  Select Update File (DFM.xlsx)…",
            command=self._on_dfm_load,
        ).grid(row=1, column=0, sticky="w")

        self._dfm_lbl_status = ttk.Label(dfm_frame, text="", foreground="grey", anchor="w")
        self._dfm_lbl_status.grid(row=1, column=1, sticky="ew", padx=(12, 0))

        # ── Saved mappings table ─────────────────────────────────────────
        table_frame = ttk.LabelFrame(parent, text="Saved Mappings", padding=8)
        table_frame.grid(row=3, column=0, sticky="nsew", padx=4, pady=(4, 4))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self._elvy_tree = ttk.Treeview(
            table_frame,
            columns=("elvy", "delta"),
            show="headings",
            selectmode="browse",
        )
        self._elvy_tree.heading("elvy", text="Article No (Articolo Elvy)",
                                 command=lambda: self._on_elvy_sort("elvy"))
        self._elvy_tree.heading("delta", text="Articolo Delta",
                                 command=lambda: self._on_elvy_sort("delta"))
        self._elvy_tree.column("elvy", width=260, anchor="w")
        self._elvy_tree.column("delta", width=260, anchor="w")
        self._elvy_tree.tag_configure("oddrow", background="#FFFFFF")
        self._elvy_tree.tag_configure("evenrow", background="#EAF1FB")
        self._elvy_tree.grid(row=0, column=0, sticky="nsew")

        # Which column the table is currently sorted by, and whether
        # ascending or descending — toggled by _on_elvy_sort on each click.
        self._elvy_sort_column: str = "elvy"
        self._elvy_sort_reverse: bool = False

        tree_scroll = ttk.Scrollbar(table_frame, command=self._elvy_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self._elvy_tree.configure(yscrollcommand=tree_scroll.set)

        btn_row = ttk.Frame(table_frame)
        btn_row.grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Button(
            btn_row, text="✏  Edit Selected", command=self._on_elvy_edit
        ).grid(row=0, column=0, padx=(0, 6))

        ttk.Button(
            btn_row, text="🗑  Delete Selected", command=self._on_elvy_delete
        ).grid(row=0, column=1)

        # Tracks the original Article No of the row being edited (if any),
        # so Save can detect a renamed key and remove the old entry instead
        # of leaving a stale duplicate behind.
        self._elvy_editing_key: str | None = None

        self._refresh_elvy_tree()
        self._refresh_dfm_status()

    def _refresh_dfm_status(self) -> None:
        """Show what DFM reference (if any) is currently cached, and when."""
        cache = load_dfm_cache()
        entries = cache.get("entries", [])
        if entries:
            self._dfm_lbl_status.config(
                text=f"Loaded: {cache.get('source_file', '?')} "
                     f"({len(entries)} colour entries, {cache.get('loaded_at', '?')})",
                foreground="grey",
            )
        else:
            self._dfm_lbl_status.config(
                text="No DFM reference loaded yet — COLOREDFM/CLDESCR columns will stay blank.",
                foreground="grey",
            )

    def _on_dfm_load(self) -> None:
        path = filedialog.askopenfilename(
            title="Select the DFM.xlsx export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            lookup = build_dfm_lookup(Path(path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Could Not Load File", str(exc))
            return

        if not lookup:
            messagebox.showwarning(
                "No Elvy Rows Found",
                "This file was read fine, but no ARTICOLODFM rows starting with "
                "\"C130\" (Elvy) were found in it.",
            )
            return

        save_dfm_cache(lookup, Path(path).name)
        self._refresh_dfm_status()
        messagebox.showinfo(
            "DFM Reference Loaded",
            f"Loaded {len(lookup)} Elvy colour entries from {Path(path).name}.",
        )

    # ------------------------------------------------------------------
    # Elvy mapping table handlers
    # ------------------------------------------------------------------

    def _refresh_elvy_tree(self) -> None:
        """Reload the mapping from disk and repopulate the table, sorted by
        whichever column/direction was last clicked, with zebra striping."""
        for item in self._elvy_tree.get_children():
            self._elvy_tree.delete(item)
        mapping = load_elvy_mapping()

        key_index = 0 if self._elvy_sort_column == "elvy" else 1
        items = sorted(
            mapping.items(),
            key=lambda kv: kv[key_index].lower(),
            reverse=self._elvy_sort_reverse,
        )
        for i, (elvy_code, delta_code) in enumerate(items):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self._elvy_tree.insert("", "end", values=(elvy_code, delta_code), tags=(tag,))

    def _on_elvy_sort(self, column: str) -> None:
        """Column header clicked: sort by it ascending, or flip direction
        if it's already the active sort column."""
        if self._elvy_sort_column == column:
            self._elvy_sort_reverse = not self._elvy_sort_reverse
        else:
            self._elvy_sort_column = column
            self._elvy_sort_reverse = False
        self._refresh_elvy_tree()

    def _on_elvy_save(self) -> None:
        elvy_code = self._elvy_entry_elvy.get().strip()
        delta_code = self._elvy_entry_delta.get().strip()

        if not elvy_code:
            messagebox.showwarning(
                "Missing Article No", "Please enter an Article No (Articolo Elvy)."
            )
            return

        # If editing an existing row and the Article No (key) was changed,
        # remove the old entry first so it isn't left behind as a duplicate.
        if self._elvy_editing_key and self._elvy_editing_key != elvy_code:
            delete_elvy_mapping(self._elvy_editing_key)

        add_elvy_mapping(elvy_code, delta_code)
        logger.info("Elvy mapping saved: %s -> %s", elvy_code, delta_code)
        self._elvy_editing_key = None
        self._elvy_entry_elvy.delete(0, "end")
        self._elvy_entry_delta.delete(0, "end")
        self._refresh_elvy_tree()

    def _on_elvy_edit(self) -> None:
        """Load the selected row's values into the entry fields for editing.
        Pressing Save afterwards updates that same row (even if the Article
        No itself is changed — see _on_elvy_save)."""
        selection = self._elvy_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Select a row in the table first.")
            return
        elvy_code, delta_code = self._elvy_tree.item(selection[0], "values")
        self._elvy_editing_key = elvy_code
        self._elvy_entry_elvy.delete(0, "end")
        self._elvy_entry_elvy.insert(0, elvy_code)
        self._elvy_entry_delta.delete(0, "end")
        self._elvy_entry_delta.insert(0, delta_code)

    def _on_elvy_delete(self) -> None:
        selection = self._elvy_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Select a row in the table first.")
            return
        elvy_code = self._elvy_tree.item(selection[0], "values")[0]
        delete_elvy_mapping(elvy_code)
        logger.info("Elvy mapping deleted: %s", elvy_code)
        self._refresh_elvy_tree()

    # ------------------------------------------------------------------
    # Shared log area
    # ------------------------------------------------------------------

    def _build_log_area(self) -> None:
        log_frame = ttk.LabelFrame(self, text="Log", padding=6)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self._log_text = tk.Text(
            log_frame,
            state="disabled",
            wrap="word",
            font=("Courier New", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            relief="flat",
            height=18,
        )
        self._log_text.grid(row=0, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self._log_text.configure(yscrollcommand=log_scroll.set)

        self._log_text.tag_configure("INFO", foreground="#4FC1FF")
        self._log_text.tag_configure("WARNING", foreground="#FFD700")
        self._log_text.tag_configure("ERROR", foreground="#F44747")
        self._log_text.tag_configure("DEBUG", foreground="#858585")

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------

    def _save_prefs(self, **kwargs: object) -> None:
        """
        Update one or more keys in self._prefs and persist immediately.
        A value of None removes that key (so a stale remembered path
        doesn't linger once the person switches from folder to single-PDF
        mode or vice versa).
        """
        for key, value in kwargs.items():
            if value is None:
                self._prefs.pop(key, None)
            else:
                self._prefs[key] = value
        save_settings(self._prefs)

    def _restore_saved_paths(self) -> None:
        """
        Re-populate the PDF/folder/output selections on all three
        conversion tabs from what was last used, so the person doesn't
        have to reselect the same paths every session. Silently skips any
        remembered path that no longer exists on disk (moved/deleted/on
        an unavailable drive) rather than pointing at a dead path.
        """
        tabs = (
            ("po", self._po_lbl_pdf_path, self._po_lbl_folder_path, self._po_lbl_output_path),
            ("bolla", self._bolla_lbl_pdf_path, self._bolla_lbl_folder_path, self._bolla_lbl_output_path),
            ("einv", self._einv_lbl_pdf_path, self._einv_lbl_folder_path, self._einv_lbl_output_path),
        )
        for prefix, lbl_pdf, lbl_folder, lbl_output in tabs:
            pdf_str = self._prefs.get(f"{prefix}_pdf_path")
            if pdf_str and Path(pdf_str).is_file():
                setattr(self, f"_{prefix}_pdf_path", Path(pdf_str))
                lbl_pdf.config(text=pdf_str, foreground="black")

            folder_str = self._prefs.get(f"{prefix}_folder_path")
            if folder_str and Path(folder_str).is_dir():
                setattr(self, f"_{prefix}_folder_path", Path(folder_str))
                lbl_folder.config(text=folder_str, foreground="black")

            output_str = self._prefs.get(f"{prefix}_output_dir")
            if output_str and Path(output_str).is_dir():
                setattr(self, f"_{prefix}_output_dir", Path(output_str))
                lbl_output.config(text=output_str, foreground="black")

        erp_str = self._prefs.get("po_erp_file_path")
        if erp_str and Path(erp_str).is_file():
            self._po_erp_file_path = Path(erp_str)
            self._po_lbl_erp_file.config(text=erp_str, foreground="black")
        if self._prefs.get("po_update_erp_file"):
            self._po_update_erp_file.set(True)

    # ------------------------------------------------------------------
    # Purchase Orders — file/folder selection callbacks
    # ------------------------------------------------------------------

    def _on_po_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("po_last_dir") or None,
        )
        if path:
            self._po_pdf_path = Path(path)
            self._po_folder_path = None
            self._po_lbl_pdf_path.config(text=str(self._po_pdf_path), foreground="black")
            self._po_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(po_pdf_path=str(self._po_pdf_path), po_folder_path=None,
                              po_last_dir=str(self._po_pdf_path.parent))

    def _on_po_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of PDFs",
            initialdir=self._prefs.get("po_last_dir") or None,
        )
        if path:
            self._po_folder_path = Path(path)
            self._po_pdf_path = None
            self._po_lbl_folder_path.config(text=str(self._po_folder_path), foreground="black")
            self._po_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(po_folder_path=str(self._po_folder_path), po_pdf_path=None,
                              po_last_dir=str(self._po_folder_path))

    def _on_po_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("po_output_dir") or None,
        )
        if path:
            self._po_output_dir = Path(path)
            self._po_lbl_output_path.config(text=str(self._po_output_dir), foreground="black")
            self._save_prefs(po_output_dir=str(self._po_output_dir))

    def _on_po_select_erp_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select the existing ERP import Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self._prefs.get("po_erp_file_path") or None,
        )
        if path:
            self._po_erp_file_path = Path(path)
            self._po_lbl_erp_file.config(text=str(self._po_erp_file_path), foreground="black")
            self._save_prefs(po_erp_file_path=str(self._po_erp_file_path))

    # ------------------------------------------------------------------
    # Purchase Orders — convert callback
    # ------------------------------------------------------------------

    def _on_po_convert(self) -> None:
        """Validate inputs then kick off conversion in a background thread."""
        input_source: Path | None = self._po_pdf_path or self._po_folder_path
        if input_source is None:
            messagebox.showwarning(
                "No Input", "Please select a PDF file or a folder first."
            )
            return

        if self._po_output_dir is None:
            messagebox.showwarning(
                "No Output Folder", "Please select an output folder first."
            )
            return

        pdf_list = find_pdfs(input_source)
        if not pdf_list:
            messagebox.showwarning(
                "No PDFs Found", f"No PDF files found in:\n{input_source}"
            )
            return

        self._set_po_ui_enabled(False)
        self._po_progress["value"] = 0
        self._po_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_po_conversion,
            args=(
                pdf_list,
                self._po_output_dir,
                self._po_one_per_file.get(),
                self._po_update_erp_file.get(),
                self._po_erp_file_path,
            ),
            daemon=True,
        )
        thread.start()

    # ------------------------------------------------------------------
    # Purchase Orders — background conversion logic
    # ------------------------------------------------------------------

    def _run_po_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
        update_erp_file: bool,
        erp_file_path: Path | None,
    ) -> None:
        """
        Run in a background thread.

        Abbina is mandatory and hybrid, not a user choice: a PDF's own
        Machine annotation is kept exactly as extracted wherever present;
        any same-colour group left without one is filled in automatically
        (smallest machine that fits the group's combined Quantity/Cones).
        """
        errors: list[str] = []
        merged_rows: list[OrderRow] = []
        all_rows: list[OrderRow] = []
        total = len(pdf_list)
        calculator = AbbinaCalculator()
        elvy_mapping = load_elvy_mapping()
        dfm_entries = load_dfm_cache().get("entries", [])

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_po_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows = PDFParser(pdf_path).parse()

                for row in rows:
                    row.articolo_delta = lookup_articolo_delta(row.article_no, elvy_mapping)
                    row.coloredfm, row.cldescr = lookup_dfm_color(
                        row.articolo_delta, row.colour, row.ne, row.dye_type, row.yarn,
                        dfm_entries,
                    )

                calculator.calculate(rows, only_if_missing=True)
                all_rows.extend(rows)

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    ExcelExporter(rows, out_path).export()
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_po_progress()

        # Final export for merged mode
        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "merged_purchase_orders.xlsx"
                ExcelExporter(merged_rows, out_path).export()
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        # Also push every row processed in this run into the existing ERP
        # file, if configured — this clears its old data rows first.
        if update_erp_file and all_rows:
            if erp_file_path is None:
                msg = "ERP file update was enabled but no file is selected."
                logger.error(msg)
                errors.append(msg)
            else:
                try:
                    ordini_rows = build_ordini_elvy_rows(all_rows)
                    n = update_existing_ordini_file(erp_file_path, ordini_rows)
                    logger.info("Updated ERP file: %s (%d rows)", erp_file_path.name, n)
                except Exception as exc:  # noqa: BLE001
                    msg = f"Error updating ERP file {erp_file_path.name}: {exc}"
                    logger.error(msg)
                    errors.append(msg)

        self.after(0, self._on_po_conversion_done, errors, total)

    def _on_po_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_po_ui_enabled(True)

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n"
                f"{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_po_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} PDF file(s).\n\n"
                f"Output saved to:\n{self._po_output_dir}",
            )
            self._set_po_status(f"Done — {total} file(s) converted successfully.")

    def _set_po_status(self, text: str) -> None:
        self.after(0, self._po_lbl_status.config, {"text": text})

    def _advance_po_progress(self) -> None:
        def _inc() -> None:
            self._po_progress["value"] += 1
        self.after(0, _inc)

    def _set_po_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._po_btn_pdf,
                self._po_btn_folder,
                self._po_btn_output,
                self._po_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)

    # ------------------------------------------------------------------
    # Bolla — file/folder selection callbacks
    # ------------------------------------------------------------------

    def _on_bolla_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a Bolla PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("bolla_last_dir") or None,
        )
        if path:
            self._bolla_pdf_path = Path(path)
            self._bolla_folder_path = None
            self._bolla_lbl_pdf_path.config(text=str(self._bolla_pdf_path), foreground="black")
            self._bolla_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(bolla_pdf_path=str(self._bolla_pdf_path), bolla_folder_path=None,
                              bolla_last_dir=str(self._bolla_pdf_path.parent))

    def _on_bolla_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of Bolla PDFs",
            initialdir=self._prefs.get("bolla_last_dir") or None,
        )
        if path:
            self._bolla_folder_path = Path(path)
            self._bolla_pdf_path = None
            self._bolla_lbl_folder_path.config(text=str(self._bolla_folder_path), foreground="black")
            self._bolla_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(bolla_folder_path=str(self._bolla_folder_path), bolla_pdf_path=None,
                              bolla_last_dir=str(self._bolla_folder_path))

    def _on_bolla_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("bolla_output_dir") or None,
        )
        if path:
            self._bolla_output_dir = Path(path)
            self._bolla_lbl_output_path.config(text=str(self._bolla_output_dir), foreground="black")
            self._save_prefs(bolla_output_dir=str(self._bolla_output_dir))

    # ------------------------------------------------------------------
    # Bolla — convert callback
    # ------------------------------------------------------------------

    def _on_bolla_convert(self) -> None:
        input_source: Path | None = self._bolla_pdf_path or self._bolla_folder_path
        if input_source is None:
            messagebox.showwarning(
                "No Input", "Please select a PDF file or a folder first."
            )
            return

        if self._bolla_output_dir is None:
            messagebox.showwarning(
                "No Output Folder", "Please select an output folder first."
            )
            return

        pdf_list = find_pdfs(input_source)
        if not pdf_list:
            messagebox.showwarning(
                "No PDFs Found", f"No PDF files found in:\n{input_source}"
            )
            return

        self._set_bolla_ui_enabled(False)
        self._bolla_progress["value"] = 0
        self._bolla_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_bolla_conversion,
            args=(pdf_list, self._bolla_output_dir, self._bolla_one_per_file.get()),
            daemon=True,
        )
        thread.start()

    # ------------------------------------------------------------------
    # Bolla — background conversion logic
    # ------------------------------------------------------------------

    def _run_bolla_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
    ) -> None:
        errors: list[str] = []
        merged_rows: list[BollaRow] = []
        merged_totals: list[BollaTotals] = []
        total = len(pdf_list)

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_bolla_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows, totals = BollaParser(pdf_path).parse()

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    BollaExporter(rows, [totals] if totals else [], out_path).export()
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)
                    if totals:
                        merged_totals.append(totals)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_bolla_progress()

        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "merged_bolle.xlsx"
                BollaExporter(merged_rows, merged_totals, out_path).export()
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        self.after(0, self._on_bolla_conversion_done, errors, total)

    def _on_bolla_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_bolla_ui_enabled(True)

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n"
                f"{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_bolla_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} Bolla PDF file(s).\n\n"
                f"Output saved to:\n{self._bolla_output_dir}",
            )
            self._set_bolla_status(f"Done — {total} file(s) converted successfully.")

    def _set_bolla_status(self, text: str) -> None:
        self.after(0, self._bolla_lbl_status.config, {"text": text})

    def _advance_bolla_progress(self) -> None:
        def _inc() -> None:
            self._bolla_progress["value"] += 1
        self.after(0, _inc)

    def _set_bolla_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._bolla_btn_pdf,
                self._bolla_btn_folder,
                self._bolla_btn_output,
                self._bolla_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)

    # ------------------------------------------------------------------
    # Elvy Invoice tab
    # ------------------------------------------------------------------

    def _build_elvy_invoice_tab(self, parent: ttk.Frame) -> None:
        """
        Converts Elvy's raw-yarn "Invoice to Delta" PDFs (Proforma invoice +
        Packing list, 2 pages). Extracts Inv No / Date, one row per Pos
        (Yarn Code, NM, Ne, Yarn Type, LOT, Price/Net/Gross/Value), looks
        up each Yarn Code's Articolo Delta from the same Elvy mapping used
        on the Purchase Orders tab, and reads each Pos's total No. of
        cones from the packing-list page into a Rocche column.
        """
        parent.columnconfigure(0, weight=1)

        info = ttk.Label(
            parent,
            text="Elvy-specific data. Converts Elvy's raw-yarn \"Invoice to Delta\" "
                 "PDFs (Proforma invoice + Packing list). Each row's Yarn Code is "
                 "looked up against the same Elvy mapping from the Elvy tab to fill "
                 "in Articolo Delta, and Rocche is read from the packing-list page's "
                 "total No. of cones per Pos.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        info.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=8)
        sel_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 4))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._einv_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_einv_select_pdf, width=20
        )
        self._einv_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._einv_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_einv_select_folder, width=20
        )
        self._einv_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._einv_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_einv_select_output, width=20
        )
        self._einv_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._einv_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)

        self._einv_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)

        self._einv_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4)

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(8, 4))
        opt_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=2)
        opt_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            opt_frame,
            text="One Excel file per invoice  (otherwise merge all into one workbook)",
            variable=self._einv_one_per_file,
        ).grid(row=0, column=0, sticky="w")

        self._einv_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_einv_convert,
            style="Accent.TButton",
            width=14,
        )
        self._einv_btn_convert.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(8, 4))
        prog_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=2)
        prog_frame.columnconfigure(0, weight=1)

        self._einv_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._einv_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._einv_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._einv_lbl_status.grid(row=1, column=0, sticky="ew")

    def _on_einv_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select an Elvy Invoice PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("einv_last_dir") or None,
        )
        if path:
            self._einv_pdf_path = Path(path)
            self._einv_folder_path = None
            self._einv_lbl_pdf_path.config(text=str(self._einv_pdf_path), foreground="black")
            self._einv_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(einv_pdf_path=str(self._einv_pdf_path), einv_folder_path=None,
                              einv_last_dir=str(self._einv_pdf_path.parent))

    def _on_einv_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of Elvy Invoice PDFs",
            initialdir=self._prefs.get("einv_last_dir") or None,
        )
        if path:
            self._einv_folder_path = Path(path)
            self._einv_pdf_path = None
            self._einv_lbl_folder_path.config(text=str(self._einv_folder_path), foreground="black")
            self._einv_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(einv_folder_path=str(self._einv_folder_path), einv_pdf_path=None,
                              einv_last_dir=str(self._einv_folder_path))

    def _on_einv_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("einv_output_dir") or None,
        )
        if path:
            self._einv_output_dir = Path(path)
            self._einv_lbl_output_path.config(text=str(self._einv_output_dir), foreground="black")
            self._save_prefs(einv_output_dir=str(self._einv_output_dir))

    def _on_einv_convert(self) -> None:
        input_source: Path | None = self._einv_pdf_path or self._einv_folder_path
        if input_source is None:
            messagebox.showwarning("No Input", "Please select a PDF file or a folder first.")
            return

        if self._einv_output_dir is None:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return

        pdf_list = find_pdfs(input_source)
        if not pdf_list:
            messagebox.showwarning("No PDFs Found", f"No PDF files found in:\n{input_source}")
            return

        self._set_einv_ui_enabled(False)
        self._einv_progress["value"] = 0
        self._einv_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_einv_conversion,
            args=(pdf_list, self._einv_output_dir, self._einv_one_per_file.get()),
            daemon=True,
        )
        thread.start()

    def _run_einv_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
    ) -> None:
        errors: list[str] = []
        merged_rows: list[ElvyInvoiceRow] = []
        total = len(pdf_list)
        elvy_mapping = load_elvy_mapping()

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_einv_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows = ElvyInvoiceParser(pdf_path).parse()
                for row in rows:
                    row.articolo_delta = lookup_articolo_delta(row.yarn_code, elvy_mapping)

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    ElvyInvoiceExporter(rows, out_path).export()
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_einv_progress()

        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "merged_elvy_invoices.xlsx"
                ElvyInvoiceExporter(merged_rows, out_path).export()
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        self.after(0, self._on_einv_conversion_done, errors, total)

    def _on_einv_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_einv_ui_enabled(True)

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_einv_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} Elvy Invoice PDF file(s).\n\n"
                f"Output saved to:\n{self._einv_output_dir}",
            )
            self._set_einv_status(f"Done — {total} file(s) converted successfully.")

    def _set_einv_status(self, text: str) -> None:
        self.after(0, self._einv_lbl_status.config, {"text": text})

    def _advance_einv_progress(self) -> None:
        def _inc() -> None:
            self._einv_progress["value"] += 1
        self.after(0, _inc)

    def _set_einv_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._einv_btn_pdf,
                self._einv_btn_folder,
                self._einv_btn_output,
                self._einv_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)

    # ------------------------------------------------------------------
    # Log window integration (shared)
    # ------------------------------------------------------------------

    def _attach_log_handler(self) -> None:
        """Route logger records to the GUI log window."""
        handler = _QueueHandler(self._log_queue)
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s  %(levelname)-8s  %(message)s",
                datefmt="%H:%M:%S",
            )
        )
        logger.addHandler(handler)

    def _poll_log_queue(self) -> None:
        """Drain the log queue and append entries to the Text widget."""
        try:
            while True:
                record = self._log_queue.get_nowait()
                self._append_log(record)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_log_queue)

    def _append_log(self, message: str) -> None:
        """Append *message* to the log Text widget with level-based colour."""
        self._log_text.configure(state="normal")

        tag = "INFO"
        for level in ("DEBUG", "WARNING", "ERROR"):
            if level in message:
                tag = level
                break

        self._log_text.insert("end", message + "\n", tag)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")
