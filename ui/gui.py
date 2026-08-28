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
from ordini_elvy import (
    build_ordini_elvy_rows,
    export_filato_full,
    export_ordini_full,
    match_raw_yarn,
    read_filato_tinturia_sheet,
)
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
    is_first_time_dyeing,
    load_dfm_cache,
    lookup_dfm_color,
    raw_to_finished_articolo,
    save_dfm_cache,
)
import prezzi_logic
from magazino_cache import load_magazino_cache, save_magazino_cache
from utils import find_pdfs, load_settings, logger, make_output_path, save_settings
from modern_widgets import RoundedButton

# Keep the existing button call sites and their commands, but render them as
# rounded pill controls throughout the application.
ttk.Button = RoundedButton

from situazione_tab import SituazioneTab
from situazione_settimana_tab import SettimanaTab
from magazino_filato_tab import MagazinoFilatoTab
from kamal_tab import KamalTab
from ui.tabs.ordine_med_tab import OrdineMedTab
from ui.tabs.overview_tab import OverviewTab
from ui.tabs.prezzi_tab import PrezziTab
from biglietti_tab import BigliettiTab


def _resource_path(filename: str) -> Path:
    """
    Resolve the path to a bundled resource (e.g. icon.ico) so it works both
    when running from source and when frozen by PyInstaller.

    A frozen --onedir build extracts/keeps bundled data files next to the
    .exe under ``sys._MEIPASS`` — a plain ``Path(__file__).parent`` lookup
    (which works fine in dev mode) resolves to the wrong place once frozen,
    since gui.py itself no longer lives on disk as a real file at runtime.
    """
    # In source mode this file lives under ui/, while bundled resources are
    # placed beside the executable at the project/bundle root.
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent.parent))
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

    WINDOW_TITLE = "Planing"
    WINDOW_MIN_W = 1000
    WINDOW_MIN_H = 760

    def __init__(self) -> None:
        super().__init__()
        self.title(self.WINDOW_TITLE)
        self.minsize(self.WINDOW_MIN_W, self.WINDOW_MIN_H)
        self.resizable(True, True)
        # Give the application a useful initial size.  Without an explicit
        # geometry Tk can choose a size based on the currently selected page,
        # which makes the notebook tabs easy to miss on smaller displays or
        # with high-DPI scaling enabled.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        initial_w = min(1400, max(self.WINDOW_MIN_W, int(screen_w * 0.90)))
        initial_h = min(900, max(self.WINDOW_MIN_H, int(screen_h * 0.90)))
        self.geometry(f"{initial_w}x{initial_h}")
        # Start maximized while keeping the normal Windows title bar and its
        # minimize/restore/close buttons available.
        try:
            self.state("zoomed")
        except tk.TclError:
            # Some non-Windows window managers do not support the zoomed state.
            pass
        self._set_window_icon()

        # ── Persisted state ───────────────────────────────────────────
        self._prefs = load_settings()

        # ── Purchase Orders tab state ─────────────────────────────────
        self._po_pdf_path: Path | None = None
        self._po_folder_path: Path | None = None
        self._po_output_dir: Path | None = None
        self._po_last_export_path: Path | None = None
        self._po_erp_export_dir: Path | None = None
        self._po_one_per_file = tk.BooleanVar(value=False)
        self._po_update_erp_file = tk.BooleanVar(value=False)
        self._po_update_erp_file.trace_add(
            "write",
            lambda *_a: self._save_prefs(po_update_erp_file=self._po_update_erp_file.get()),
        )
        self._po_raw_yarn_path: Path | None = None

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
        self.rowconfigure(0, weight=4)   # all application pages
        self.rowconfigure(1, weight=1)   # log area

        style = ttk.Style(self)
        # Use the same renderer that gives Situazione its reliable colored
        # headers, then modernize controls without changing the application's
        # existing light palette or the dedicated Treeview colors.
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            "TButton",
            font=("Segoe UI", 9),
            padding=(12, 7),
            relief="flat",
            borderwidth=1,
            background="#f8fafc",
            foreground="#16324f",
            bordercolor="#b8c6d6",
            lightcolor="#b8c6d6",
            darkcolor="#b8c6d6",
        )
        style.map(
            "TButton",
            background=[
                ("pressed", "#cbd5e1"),
                ("active", "#e2e8f0"),
                ("disabled", "#eef2f6"),
                ("!active", "#f8fafc"),
            ],
            foreground=[("disabled", "#94a3b8"), ("!disabled", "#16324f")],
            relief=[("pressed", "sunken"), ("!pressed", "flat")],
        )
        style.configure(
            "TEntry",
            padding=(7, 5),
            relief="flat",
            borderwidth=1,
            fieldbackground="#ffffff",
            foreground="#1d2939",
            bordercolor="#b8c6d6",
            lightcolor="#b8c6d6",
            darkcolor="#b8c6d6",
        )
        style.map("TEntry", bordercolor=[("focus", "#5b9bd5")])
        # Keep the page tabs compact enough to remain visible on smaller
        # displays while preserving the normal Notebook tab appearance.
        style.configure("TNotebook.Tab", padding=(10, 5))
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
        notebook.add(elvy_tab, text="Data Elvy")
        self._build_elvy_tab(elvy_tab)

        # ── Biglietti: ERP order -> ELVY/MED workbook + Word tickets ──
        self._biglietti_tab = BigliettiTab(
            notebook, self._prefs, self._save_prefs, logger,
            on_shared_cache_changed=self._on_shared_cache_changed,
        )
        notebook.add(self._biglietti_tab, text="Create (EXCEL+Biglietti)")

        # ── Ordine: Ordine Elvy + Ordine Kamal, grouped under one parent tab ──
        ordine_parent = ttk.Frame(notebook)
        notebook.add(ordine_parent, text="Ordine")
        ordine_notebook = ttk.Notebook(ordine_parent)
        ordine_notebook.pack(fill="both", expand=True)

        po_tab = ttk.Frame(ordine_notebook)
        ordine_notebook.add(po_tab, text="Ordine Elvy")
        self._build_po_tab(po_tab)

        self._kamal_tab = KamalTab(ordine_notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        ordine_notebook.add(self._kamal_tab, text="Ordine Kamal")

        self._ordine_med_tab = OrdineMedTab(
            ordine_notebook, situazione_tab=None, prefs=self._prefs,
            save_prefs=self._save_prefs, logger=logger,
            on_shared_cache_changed=self._on_shared_cache_changed,
        )
        ordine_notebook.add(self._ordine_med_tab, text="Ordine Med")

        # ── Invoice: Bolla Med + Invoice Elvy, grouped under one parent tab ──
        invoice_parent = ttk.Frame(notebook)
        notebook.add(invoice_parent, text="Invoice")
        invoice_notebook = ttk.Notebook(invoice_parent)
        invoice_notebook.pack(fill="both", expand=True)

        bolla_tab = ttk.Frame(invoice_notebook)
        invoice_notebook.add(bolla_tab, text="Bolla Med")
        self._build_bolla_tab(bolla_tab)

        elvy_invoice_tab = ttk.Frame(invoice_notebook)
        invoice_notebook.add(elvy_invoice_tab, text="Invoice Elvy")
        self._build_elvy_invoice_tab(elvy_invoice_tab)

        # ── Situazione: Situazione Generale + Situazione Settimanale ──
        # Situazione tab is a self-contained module (situazione_tab.py) — it
        # manages its own uploads, SQLite state, and UI, so it's built by
        # instantiating it directly rather than through a _build_*_tab method.
        situazione_parent = ttk.Frame(notebook)
        notebook.add(situazione_parent, text="Situazione")
        situazione_notebook = ttk.Notebook(situazione_parent)
        situazione_notebook.pack(fill="both", expand=True)

        self._situazione_tab = SituazioneTab(situazione_notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        situazione_notebook.add(self._situazione_tab, text="Situazione Generale")

        self._settimana_tab = SettimanaTab(situazione_notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        situazione_notebook.add(self._settimana_tab, text="Situazione Settimanale")

        # Ordine Med's Consegna auto-scheduling needs Situazione's live
        # current_df + Copertura data, which doesn't exist until now.
        self._ordine_med_tab._situazione_tab = self._situazione_tab
        self._situazione_tab.ordine_med_tab = self._ordine_med_tab

        self._magazino_tab = MagazinoFilatoTab(notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        notebook.add(self._magazino_tab, text="Magazino Filato")

        self._prezzi_tab = PrezziTab(notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        notebook.add(self._prezzi_tab, text="Prezzi")

        # Now that Magazino Filato exists, let Situazione auto-fill its
        # "Filato Disponibile" column from it.
        self._situazione_tab.magazino_tab = self._magazino_tab
        # The cached table is already visible.  Recalculate the potentially
        # expensive yarn suggestions only after the first paint, in a worker.
        self.after(1200, self._situazione_tab.refresh_raw_yarn_match_async)

        # ── Overview: built last since it reads from the tabs above, but
        # inserted first so it's the landing page.
        self._overview_tab = OverviewTab(
            notebook,
            self._situazione_tab,
            self._magazino_tab,
            biglietti_tab=self._biglietti_tab,
            prezzi_tab=self._prezzi_tab,
            save_prefs=self._save_prefs,
            prefs=self._prefs,
        )
        notebook.insert(0, self._overview_tab, text="📊 Overview")
        notebook.select(0)

        self._build_log_area()

        def _on_any_tab_changed(_event=None) -> None:
            self._update_log_visibility(
                notebook, self._situazione_tab, self._settimana_tab, self._magazino_tab, self._kamal_tab,
                ordine_notebook, situazione_notebook,
            )
            try:
                if notebook.select() == str(self._overview_tab):
                    self._overview_tab.on_shown()
            except tk.TclError:
                pass

        notebook.bind("<<NotebookTabChanged>>", _on_any_tab_changed)
        ordine_notebook.bind("<<NotebookTabChanged>>", _on_any_tab_changed)
        situazione_notebook.bind("<<NotebookTabChanged>>", _on_any_tab_changed)
        _on_any_tab_changed()
        self._restore_saved_paths()

    def _on_shared_cache_changed(self) -> None:
        """Refresh every consumer after any shared source is uploaded."""
        self._refresh_dfm_status()
        self._refresh_magazino_status()
        if hasattr(self, "_situazione_tab"):
            self._situazione_tab.sync_shared_async()
            if hasattr(self._situazione_tab, "sync_remaining_shared_sources"):
                self._situazione_tab.sync_remaining_shared_sources()
        if getattr(self, "_settimana_tab", None):
            self._settimana_tab.sync_shared_async()
        if getattr(self, "_kamal_tab", None):
            self._kamal_tab.sync_shared_dfm()
            self._kamal_tab.sync_shared_magazino()
            self._kamal_tab.sync_shared_lotti()
        if getattr(self, "_ordine_med_tab", None):
            self._ordine_med_tab.sync_shared_magazino()
        if getattr(self, "_magazino_tab", None):
            self._magazino_tab.sync_shared_async()
            self._magazino_tab.sync_shared_lotti_async()
        if getattr(self, "_situazione_tab", None):
            # Best-effort: Magazino/LOTTI syncs above run in worker threads,
            # so this may still see slightly-stale data the first time --
            # it'll catch up on the next Situazione refresh regardless.
            self.after(500, self._situazione_tab.refresh_raw_yarn_match_async)

    def _refresh_magazino_status(self) -> None:
        cache = load_magazino_cache()
        source_path = cache.get("source_path", "")
        if source_path and Path(source_path).is_file():
            current_path = getattr(self, "_po_raw_yarn_path", None)
            if current_path is None or str(current_path) != str(source_path):
                self._po_raw_yarn_path = Path(source_path)
                self._po_lbl_raw_yarn.config(text=str(self._po_raw_yarn_path), foreground="black")
                self._save_prefs(
                    po_raw_yarn_path=str(self._po_raw_yarn_path),
                    po_last_dir=str(self._po_raw_yarn_path.parent),
                )

    # ------------------------------------------------------------------
    # Purchase Orders tab
    # ------------------------------------------------------------------

    def _build_po_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=6)
        sel_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._po_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_po_select_pdf, width=16
        )
        self._po_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._po_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_po_select_folder, width=16
        )
        self._po_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._po_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_po_select_output, width=16
        )
        self._po_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._po_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._po_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 2))

        self._po_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 2))

        self._po_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4, pady=(0, 2))

        abbina_info = ttk.Label(
            parent,
            text="Abbina uses the PDF Machine annotation when available. Rows without it are filled automatically.",
            foreground="grey", anchor="w", wraplength=620, justify="left",
        )
        abbina_info.grid(row=1, column=0, sticky="ew", padx=4, pady=(2, 0))

        # ── Raw yarn (Magazino) matching ────────────────────────────
        raw_yarn_frame = ttk.LabelFrame(parent, text="Raw Yarn Matching", padding=6)
        raw_yarn_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 2))
        raw_yarn_frame.columnconfigure(1, weight=1)

        ttk.Button(
            raw_yarn_frame, text="📦 Select Magazino File…", command=self._on_po_select_raw_yarn, width=20
        ).grid(row=0, column=0, padx=4, pady=3, sticky="w")

        self._po_lbl_raw_yarn = ttk.Label(
            raw_yarn_frame, text="No Magazino file selected", foreground="grey", anchor="w"
        )
        self._po_lbl_raw_yarn.grid(row=0, column=1, sticky="ew", padx=4)

        # ── Extract "EXCEL PER ORDINE VENDITA EGITTO" + "Filato x Tinturia" ──
        erp_frame = ttk.LabelFrame(parent, text="Also Extract ERP Files", padding=6)
        erp_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(3, 2))
        erp_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            erp_frame,
            text="After converting, extract \"EXCEL PER ORDINE VENDITA EGITTO\" and "
                 "\"Filato x Tinturia\" into the folder below — each file is (re)written "
                 "fresh, fully formatted, every Convert",
            variable=self._po_update_erp_file,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Button(
            erp_frame, text="📁 Select ERP Files Folder…", command=self._on_po_select_erp_folder, width=22
        ).grid(row=1, column=0, padx=(0, 6), pady=(3, 0), sticky="w")

        self._po_lbl_erp_dir = ttk.Label(
            erp_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_erp_dir.grid(row=1, column=1, sticky="ew", pady=(4, 0))

        erp_warning = ttk.Label(
            erp_frame,
            text="⚠ Close these files in Excel before converting, or saving will fail.",
            foreground="#8a6d00", anchor="w",
        )
        erp_warning.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(2, 0))

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(6, 3))
        opt_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=1)
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
            width=12,
        )
        self._po_btn_convert.grid(row=0, column=1, sticky="e", padx=(8, 0), pady=(0, 1))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(4, 2))
        prog_frame.grid(row=5, column=0, sticky="ew", padx=4, pady=1)
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

        # Keep the parsed reference and the original validated file location
        # in one shared cache used by both Data Elvy and Situazione.
        save_dfm_cache(lookup, Path(path).name, Path(path))
        self._refresh_dfm_status()
        if self._on_shared_cache_changed:
            self._on_shared_cache_changed()
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
        self._log_frame = log_frame
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self._log_text = tk.Text(
            log_frame,
            state="disabled",
            wrap="word",
            font=("Courier New", 8),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            relief="flat",
            height=12,
        )
        self._log_text.grid(row=0, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self._log_text.configure(yscrollcommand=log_scroll.set)

    def _update_log_visibility(self, notebook: ttk.Notebook, situazione_tab: ttk.Frame,
                                settimana_tab: ttk.Frame, magazino_tab: ttk.Frame,
                                kamal_tab: ttk.Frame, ordine_notebook: ttk.Notebook,
                                situazione_notebook: ttk.Notebook) -> None:
        """Hide the shared log where the page has its own full-screen workspace."""
        selected_top = notebook.select()
        top_text = notebook.tab(selected_top, "text")
        data_elvy_selected = top_text == "Data Elvy"
        magazino_selected = selected_top == str(magazino_tab)
        overview_selected = selected_top == str(self._overview_tab)
        prezzi_selected = selected_top == str(self._prezzi_tab)
        biglietti_selected = top_text == "Create (EXCEL+Biglietti)" or selected_top == str(self._biglietti_tab)

        situazione_selected = settimana_selected = False
        kamal_selected = ordini_selected = ordine_med_selected = False
        if top_text == "Situazione":
            inner = situazione_notebook.select()
            situazione_selected = inner == str(situazione_tab)
            settimana_selected = inner == str(settimana_tab)
        elif top_text == "Ordine":
            inner_text = ordine_notebook.tab(ordine_notebook.select(), "text")
            kamal_selected = inner_text == "Ordine Kamal"
            ordini_selected = inner_text == "Ordine Elvy"
            ordine_med_selected = inner_text == "Ordine Med"
        # Do not trigger heavy shared-file loading while switching tabs.
        # Keep the UI responsive; shared DFM/Produzione loads happen only when
        # the user explicitly refreshes or uploads on the target page.
        self._refresh_dfm_status()
        if (
            situazione_selected
            or settimana_selected
            or magazino_selected
            or kamal_selected
            or data_elvy_selected
            or ordini_selected
            or ordine_med_selected
            or overview_selected
            or prezzi_selected
            or biglietti_selected
        ):
            self._log_frame.grid_remove()
            self.rowconfigure(1, weight=0)
            self.rowconfigure(0, weight=5)
            notebook.configure(height=1)
        else:
            self._log_frame.grid()
            self.rowconfigure(0, weight=3)
            self.rowconfigure(1, weight=1)
            notebook.configure(height=220)

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

        # Magazino may be selected from Ordine ELVY, Ordine Kamal, or
        # Magazino Filato. Prefer the page preference, then use the shared
        # cache so a file selected in another page is restored here too.
        raw_yarn_str = self._prefs.get("po_raw_yarn_path", "")
        if not (raw_yarn_str and Path(raw_yarn_str).is_file()):
            raw_yarn_str = load_magazino_cache().get("source_path", "")
        if raw_yarn_str and Path(raw_yarn_str).is_file():
            self._po_raw_yarn_path = Path(raw_yarn_str)
            self._po_lbl_raw_yarn.config(text=str(self._po_raw_yarn_path), foreground="black")
            self._save_prefs(
                po_raw_yarn_path=str(self._po_raw_yarn_path),
                po_last_dir=str(self._po_raw_yarn_path.parent),
            )

        last_export_str = self._prefs.get("po_last_export_path")
        if last_export_str and Path(last_export_str).is_file():
            self._po_last_export_path = Path(last_export_str)

        erp_dir_str = self._prefs.get("po_erp_export_dir")
        if erp_dir_str and Path(erp_dir_str).is_dir():
            self._po_erp_export_dir = Path(erp_dir_str)
            self._po_lbl_erp_dir.config(text=erp_dir_str, foreground="black")
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

    def _on_po_select_erp_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select the folder for the ERP export files",
            initialdir=self._prefs.get("po_erp_export_dir") or None,
        )
        if path:
            self._po_erp_export_dir = Path(path)
            self._po_lbl_erp_dir.config(text=str(self._po_erp_export_dir), foreground="black")
            self._save_prefs(po_erp_export_dir=str(self._po_erp_export_dir))

    def _on_po_select_raw_yarn(self) -> None:
        path = filedialog.askopenfilename(
            title="Select the Magazino (raw yarn) Excel export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._po_raw_yarn_path = Path(path)
            self._po_lbl_raw_yarn.config(text=str(self._po_raw_yarn_path), foreground="black")
            save_magazino_cache(self._po_raw_yarn_path)
            self._save_prefs(po_raw_yarn_path=str(self._po_raw_yarn_path))
            self._on_shared_cache_changed()

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
                self._po_erp_export_dir,
                self._po_raw_yarn_path,
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
        erp_export_dir: Path | None,
        raw_yarn_path: Path | None = None,
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
        last_export_path: Path | None = None
        total = len(pdf_list)
        calculator = AbbinaCalculator()
        elvy_mapping = load_elvy_mapping()
        dfm_entries = load_dfm_cache().get("entries", [])
        prezzi_df = getattr(self._prezzi_tab, "prezzi_df", None)
        price_lookup = prezzi_logic.build_price_lookup(prezzi_df)

        magazino_summary = None
        codes_map = None
        if raw_yarn_path is not None:
            try:
                import magazino_logic
                magazino_df, magazino_errors = magazino_logic.load_magazino(str(raw_yarn_path))
                if magazino_errors or magazino_df is None or magazino_df.empty:
                    msg = f"Raw yarn file could not be read: {'; '.join(magazino_errors) if magazino_errors else 'empty after filtering'}"
                    logger.error(msg)
                    errors.append(msg)
                else:
                    magazino_summary = magazino_logic.summarize_by_partita(magazino_df)
                    logger.info("Raw yarn stock loaded: %d batches from %s",
                                len(magazino_summary), raw_yarn_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error loading raw yarn file: {exc}"
                logger.error(msg)
                errors.append(msg)
            try:
                # Titolo lookup for the Filato x Tinturia sheet comes from the
                # DFM reference already loaded on the Data Elvy tab (parsed
                # from DESCRIZARTICOLOLI), keyed by Articolo -- no separate
                # upload needed.
                codes_map = {e["articolo"]: e.get("titolo", "") for e in dfm_entries if e.get("titolo")}
            except Exception:  # noqa: BLE001
                codes_map = None

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
                    if is_first_time_dyeing(row.articolo_delta, row.coloredfm, dfm_entries):
                        row.check_articolo = "prima volta tint."
                    finished_articolo = raw_to_finished_articolo(row.articolo_delta)
                    if finished_articolo:
                        row.livello, row.prezzo = price_lookup.get(
                            (finished_articolo, row.coloredfm), (None, None)
                        )

                calculator.calculate(rows, only_if_missing=True)
                all_rows.extend(rows)

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    ExcelExporter(rows, out_path, magazino_summary, codes_map).export()
                    last_export_path = out_path
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
                out_path = output_dir / "Ordine_Elvy.xlsx"
                ExcelExporter(merged_rows, out_path, magazino_summary, codes_map).export()
                last_export_path = out_path
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        # Extract both ERP files fresh into the saved folder, if configured
        # -- each file is fully rebuilt (not edited in place) every Convert.
        if update_erp_file and all_rows:
            if erp_export_dir is None:
                msg = "ERP file extraction was enabled but no folder is selected."
                logger.error(msg)
                errors.append(msg)
            else:
                ordini_path = erp_export_dir / "EXCEL PER ORDINE VENDITA EGITTO.xlsx"
                filato_path = erp_export_dir / "Filato x Tinturia.xlsx"
                try:
                    ordini_rows = build_ordini_elvy_rows(all_rows)
                    if magazino_summary is not None and not magazino_summary.empty:
                        match_raw_yarn(ordini_rows, magazino_summary, codes_map)
                    n = export_ordini_full(ordini_path, ordini_rows)
                    logger.info("Extracted ERP file: %s (%d rows)", ordini_path.name, n)
                except Exception as exc:  # noqa: BLE001
                    msg = f"Error extracting {ordini_path.name}: {exc}"
                    logger.error(msg)
                    errors.append(msg)

                if last_export_path is not None:
                    try:
                        matches = read_filato_tinturia_sheet(last_export_path)
                        n2 = export_filato_full(filato_path, matches)
                        logger.info("Extracted Filato x Tinturia file: %s (%d rows)", filato_path.name, n2)
                    except Exception as exc:  # noqa: BLE001
                        msg = f"Error extracting {filato_path.name}: {exc}"
                        logger.error(msg)
                        errors.append(msg)

        self._po_last_export_path = last_export_path
        self.after(0, self._on_po_conversion_done, errors, total)

    def _on_po_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_po_ui_enabled(True)
        if self._po_last_export_path and self._po_last_export_path.is_file():
            self._save_prefs(po_last_export_path=str(self._po_last_export_path))

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
                out_path = output_dir / "Bolla_Med.xlsx"
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
                out_path = output_dir / "Invoice_Elvy.xlsx"
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
