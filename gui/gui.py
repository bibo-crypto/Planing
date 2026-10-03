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
    UI and conversion workflow: gui/workflows/purchase_orders.py

Bolla tab
    Input (PDF/Folder/Output) -> Options+Convert -> Progress
    Parses Italian delivery-note ("Bolla") PDFs via bolla_parser.py and
    exports an "Items" + "Totals" workbook via bolla_exporter.py.
    UI and conversion workflow: gui/workflows/bolla.py

Elvy Invoice workflow: gui/workflows/elvy_invoice.py
"""

from __future__ import annotations
from utility.excel_io import safe_save_workbook

import logging
import os
import queue
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utility.magazino_cache import load_magazino_cache
from utility.utils import load_settings, logger, save_settings
from utility.updater import APP_VERSION, ReleaseInfo, check_for_updates_async, download_installer, install_update
from utility import notifications
from utility.backup import create_backup, restore_backup
from gui.modern_widgets import RoundedButton
from utility.utils import keep_window_on_top, bind_escape_to_close

# Keep the existing button call sites and their commands, but render them as
# rounded pill controls throughout the application.
ttk.Button = RoundedButton

from gui.workflows import (
    BollaWorkflowMixin,
    ElvyInvoiceWorkflowMixin,
    PurchaseOrderWorkflowMixin,
)


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


def _first_text_input(widget):
    """Find the first editable text input on the currently selected page."""
    if widget.winfo_class() == "TNotebook":
        selected = widget.select()
        return _first_text_input(widget.nametowidget(selected)) if selected else None

    widget_class = widget.winfo_class()
    if widget_class in {"TEntry", "Entry", "TCombobox", "Combobox"}:
        if str(widget.cget("state")) not in {"disabled", "readonly"}:
            return widget

    for child in widget.winfo_children():
        target = _first_text_input(child)
        if target is not None:
            return target
    return None


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

class ConverterApp(PurchaseOrderWorkflowMixin, BollaWorkflowMixin, ElvyInvoiceWorkflowMixin, tk.Tk):
    """
    Root Tk window.  All UI widgets live here.
    Business logic is dispatched to background threads to keep the UI responsive.
    """

    WINDOW_TITLE = f"Planing v{APP_VERSION}"
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
        # Escape is intentionally global: operators frequently open several
        # lookup/edit windows and should never have to hunt for each X button.
        # The update dialog opts out because interrupting an in-flight update
        # can leave a partially downloaded installer.
        self.bind_all("<Escape>", self._close_child_windows, add="+")

        # ── Persisted state ───────────────────────────────────────────
        self._prefs = load_settings()

        # ── Purchase Orders tab state ─────────────────────────────────
        self._po_pdf_path: Path | None = None
        self._po_folder_path: Path | None = None
        self._po_output_dir: Path | None = None
        self._po_last_export_path: Path | None = None
        self._po_erp_export_dir: Path | None = None
        self._po_filato_export_dir: Path | None = None
        self._po_update_erp_file = tk.BooleanVar(value=True)
        self._po_update_filato_file = tk.BooleanVar(value=True)
        self._po_update_erp_file.trace_add(
            "write",
            lambda *_a: self._save_prefs(po_update_erp_file=self._po_update_erp_file.get()),
        )
        self._po_update_filato_file.trace_add(
            "write",
            lambda *_a: self._save_prefs(po_update_filato_file=self._po_update_filato_file.get()),
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

        # Let Tk paint the main window before constructing the tabs.  Several
        # tabs restore cached Excel data during construction; doing that before
        # mainloop starts makes Windows show an apparently frozen application.
        self._startup_label = ttk.Label(
            self,
            text="Planing is opening...",
            anchor="center",
            font=("Segoe UI", 14, "bold"),
        )
        # Use place so the loading message can stay visible while the real
        # notebook is built underneath it; mixing pack and grid on the root
        # would make Tk reject the layout.
        self._startup_label.place(relx=0.5, rely=0.5, anchor="center")
        self.after(50, self._finish_startup)

    def _close_child_windows(self, _event=None):
        """Close the active ordinary Toplevel, preserving its close handler."""
        update = getattr(self, "_update_window", None)
        try:
            child = _event.widget.winfo_toplevel() if _event is not None else None
            if isinstance(child, tk.Toplevel) and child is not update:
                close_handler = child.protocol("WM_DELETE_WINDOW")
                if close_handler:
                    child.tk.eval(close_handler)
                else:
                    child.destroy()
        except (AttributeError, tk.TclError):
            pass
        return "break"

    def _finish_startup(self) -> None:
        """Build the full UI after the initial window has been painted."""
        if not self.winfo_exists():
            return
        self._startup_label.config(text="Loading application pages...")
        self.update_idletasks()
        self._build_ui()
        self._attach_log_handler()
        self._poll_log_queue()
        self._startup_label.destroy()
        self._check_for_updates()

    def _check_for_updates(self) -> None:
        """Check GitHub without delaying startup, then ask before updating."""
        if not getattr(sys, "frozen", False):
            return
        check_for_updates_async(
            lambda release: self.after(0, self._show_update_prompt, release),
            logger=logger,
        )

    def _show_update_prompt(self, release: ReleaseInfo | None) -> None:
        if release is None or not self.winfo_exists():
            return
        notes = release.notes[:800] + ("…" if len(release.notes) > 800 else "")
        if not release.installer_url:
            return
        answer = messagebox.askyesno(
            "Planing update available",
            f"A new version of Planing is available.\n\n"
            f"Current version: {APP_VERSION}\nNew version: {release.version}\n\n"
            f"{notes}\n\nDownload and install it now?",
            parent=self,
        )
        if answer:
            self._start_update(release)

    def _start_update(self, release: ReleaseInfo) -> None:
        self._pending_update_version = release.version
        update_window = tk.Toplevel(self)
        self._update_window = update_window
        keep_window_on_top(update_window)
        update_window.title("Updating Planing")
        update_window.geometry("430x150")
        update_window.resizable(False, False)
        update_window.protocol("WM_DELETE_WINDOW", lambda: None)
        ttk.Label(update_window, text="Downloading Planing update…", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=18, pady=(18, 8)
        )
        self._update_status = ttk.Label(update_window, text="Downloaded: 0%")
        self._update_status.pack(anchor="w", padx=18)
        self._update_progress_bar = ttk.Progressbar(update_window, mode="determinate", maximum=100)
        self._update_progress_bar.pack(fill="x", padx=18, pady=(8, 18))

        def worker() -> None:
            try:
                installer = download_installer(
                    release,
                    progress=lambda percent: self.after(0, self._update_download_progress, percent),
                )
                self.after(0, self._finish_update, installer)
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda exc=exc: self._update_failed(str(exc)))

        import threading
        threading.Thread(target=worker, name="planing-update-download", daemon=True).start()

    def _update_download_progress(self, percent: int) -> None:
        if not getattr(self, "_update_window", None) or not self._update_window.winfo_exists():
            return
        percent = max(0, min(100, int(percent)))
        self._update_progress_bar["value"] = percent
        self._update_status.config(text=f"Downloaded: {percent}%")

    def _update_failed(self, error: str) -> None:
        if getattr(self, "_update_window", None) and self._update_window.winfo_exists():
            self._update_window.destroy()
        messagebox.showerror("Planing update", error, parent=self)

    def _finish_update(self, installer: Path) -> None:
        try:
            self._update_status.config(text="Installing update…")
            self._update_progress_bar["value"] = 100
            self.update_idletasks()
            install_update(
                installer,
                Path(sys.executable).resolve().parent,
                self._pending_update_version,
            )
            # self.destroy() alone stops mainloop() but does not guarantee
            # the OS process actually terminates -- if anything is still
            # holding a reference (a lingering thread, a live COM object
            # from the Outlook integration, etc.), Windows can leave this
            # process running in the background, still holding the exe/DLL
            # file locks the update script needs released. Force the process
            # to end so the helper can replace the executable immediately.
            self.destroy()
            os._exit(0)
        except Exception as exc:  # noqa: BLE001
            if getattr(self, "_update_window", None) and self._update_window.winfo_exists():
                self._update_window.destroy()
            messagebox.showerror("Planing update", str(exc), parent=self)

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
        from gui.tabs.biglietti_tab import BigliettiTab
        from gui.tabs.kamal_tab import KamalTab
        from gui.tabs.magazino_filato_tab import MagazinoFilatoTab
        from gui.tabs.ordine_med_tab import OrdineMedTab
        from gui.tabs.overview_tab import OverviewTab
        from gui.tabs.prezzi_tab import PrezziTab
        from gui.tabs.situazione_settimana_tab import SettimanaTab
        from gui.tabs.situazione_tab import SituazioneTab
        from gui.tabs.master_data_tab import MasterDataTab

        def startup_step(text: str) -> None:
            self._startup_label.config(text=text)
            self.update_idletasks()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)   # notification bar
        self.rowconfigure(1, weight=4)   # all application pages
        self.rowconfigure(2, weight=1)   # log area

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

        notification_bar = ttk.Frame(self, padding=(12, 5))
        notification_bar.grid(row=0, column=0, sticky="ew")
        notification_bar.columnconfigure(0, weight=1)
        self._notification_button = ttk.Button(notification_bar, command=self._open_notifications, danger=False)
        self._notification_button.grid(row=0, column=1, sticky="e", padx=3)
        ttk.Button(notification_bar, text="Backup", command=self._create_backup).grid(row=0, column=2, padx=3)
        ttk.Button(notification_bar, text="Restore", command=self._restore_backup).grid(row=0, column=3, padx=3)
        style.configure("Notification.TButton", background="#b91c1c", foreground="white")
        style.map("Notification.TButton", background=[("active", "#991b1b")])
        self._refresh_notification_badge()

        # ── All pages use one normal tab bar; Data Elvy is not persistent ──
        startup_step("Loading Data Elvy...")
        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 4))

        elvy_tab = ttk.Frame(notebook)
        notebook.add(elvy_tab, text="Data Elvy")
        self._build_elvy_tab(elvy_tab)

        # ── Biglietti: ERP order -> ELVY/MED workbook + Word tickets ──
        startup_step("Loading Create (EXCEL+Biglietti)...")
        self._biglietti_tab = BigliettiTab(
            notebook, self._prefs, self._save_prefs, logger,
            on_shared_cache_changed=self._on_shared_cache_changed,
            on_notification=self._add_notification,
        )
        notebook.add(self._biglietti_tab, text="Create (EXCEL+Biglietti)")

        # ── Ordine: Ordine Elvy + Ordine Kamal, grouped under one parent tab ──
        startup_step("Loading Order pages...")
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
            on_notification=self._add_notification,
        )
        ordine_notebook.add(self._ordine_med_tab, text="Ordine Med")

        # ── Invoice: Bolla Med + Invoice Elvy, grouped under one parent tab ──
        startup_step("Loading Invoice pages...")
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
        startup_step("Loading Situation pages...")
        # Situazione tab is a self-contained module (situazione_tab.py) — it
        # manages its own uploads, SQLite state, and UI, so it's built by
        # instantiating it directly rather than through a _build_*_tab method.
        situazione_parent = ttk.Frame(notebook)
        notebook.add(situazione_parent, text="Situazione")
        situazione_notebook = ttk.Notebook(situazione_parent)
        situazione_notebook.pack(fill="both", expand=True)

        self._situazione_tab = SituazioneTab(
            situazione_notebook, on_shared_cache_changed=self._on_shared_cache_changed,
            on_notification=self._add_notification,
        )
        situazione_notebook.add(self._situazione_tab, text="Situazione Generale")

        self._settimana_tab = SettimanaTab(situazione_notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        situazione_notebook.add(self._settimana_tab, text="Situazione Settimanale")

        # Ordine Med's Consegna auto-scheduling needs Situazione's live
        # current_df + Copertura data, which doesn't exist until now.
        self._ordine_med_tab._situazione_tab = self._situazione_tab
        self._situazione_tab.ordine_med_tab = self._ordine_med_tab

        self._magazino_tab = MagazinoFilatoTab(notebook, on_shared_cache_changed=self._on_shared_cache_changed)
        notebook.add(self._magazino_tab, text="Magazino Filato")

        startup_step("Loading Prices...")
        self._prezzi_tab = PrezziTab(
            notebook, on_shared_cache_changed=self._on_shared_cache_changed,
            on_notification=self._add_notification,
            on_notifications=self._add_notifications,
        )
        notebook.add(self._prezzi_tab, text="Prezzi")

        self._master_data_tab = MasterDataTab(notebook, on_data_changed=self._refresh_notification_badge)
        notebook.add(self._master_data_tab, text="Master Data")

        # Now that Magazino Filato exists, let Situazione auto-fill its
        # "Filato Disponibile" column from it.
        self._situazione_tab.magazino_tab = self._magazino_tab

        # ── Overview: built last since it reads from the tabs above, but
        # inserted first so it's the landing page.
        startup_step("Loading Overview...")
        self._overview_tab = OverviewTab(
            notebook,
            self._situazione_tab,
            self._magazino_tab,
            biglietti_tab=self._biglietti_tab,
            prezzi_tab=self._prezzi_tab,
            save_prefs=self._save_prefs,
            prefs=self._prefs,
            on_shared_cache_changed=self._on_shared_cache_changed,
            settimana_tab=self._settimana_tab,
        )
        notebook.insert(0, self._overview_tab, text="📊 Overview")
        notebook.select(0)

        self._build_log_area()

        def _on_any_tab_changed(_event=None) -> None:
            self._update_log_visibility(
                notebook, self._situazione_tab, self._settimana_tab, self._magazino_tab,
                ordine_notebook, situazione_notebook,
            )
            try:
                if notebook.select() == str(situazione_parent) and situazione_notebook.select() == str(self._situazione_tab):
                    self._situazione_tab.on_shown()
            except tk.TclError:
                pass
            try:
                if notebook.select() == str(situazione_parent) and situazione_notebook.select() == str(self._settimana_tab):
                    self._settimana_tab.on_shown()
            except tk.TclError:
                pass
            try:
                if notebook.select() == str(self._overview_tab):
                    self._overview_tab.on_shown()
            except tk.TclError:
                pass
            changed_notebook = getattr(_event, "widget", notebook)
            self.after_idle(self._focus_page_text_input, changed_notebook)

        def bind_notebook_events(parent) -> None:
            for child in parent.winfo_children():
                if isinstance(child, ttk.Notebook):
                    child.bind("<<NotebookTabChanged>>", _on_any_tab_changed, add="+")
                bind_notebook_events(child)

        bind_notebook_events(self)
        _on_any_tab_changed()
        self._restore_saved_paths()

    def _focus_page_text_input(self, notebook: ttk.Notebook) -> None:
        """Place the caret in the first editable input on the active page."""
        try:
            target = _first_text_input(notebook)
            if target is None or not target.winfo_exists() or not target.winfo_ismapped():
                return
            target.focus_set()
            if target.winfo_class() in {"TEntry", "Entry"}:
                target.icursor(tk.END)
        except tk.TclError:
            pass

    def _on_shared_cache_changed(self) -> None:
        """Refresh every consumer after any shared source is uploaded."""
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
            self._situazione_tab.refresh_prezzo_densita()
            # Best-effort: Magazino/LOTTI syncs above run in worker threads,
            # so this may still see slightly-stale data the first time --
            # it'll catch up on the next Situazione refresh regardless.
            self.after(500, self._situazione_tab.refresh_raw_yarn_match_async)

    def _add_notification(self, key: str, title: str, message: str, page: str, severity: str = "medium", partita_colore: str = "") -> None:
        def add_now():
            notifications.add(key, title, message, page, severity, partita_colore=partita_colore)
            self._refresh_notification_badge()
        self.after(0, add_now)

    def _add_notifications(self, entries: list[dict]) -> None:
        def add_now():
            notifications.add_many(entries)
            self._refresh_notification_badge()
        self.after(0, add_now)

    def _refresh_notification_badge(self) -> None:
        if not hasattr(self, "_notification_button"):
            return
        count = len(notifications.list_open())
        self._notification_button.configure(
            text=f"🔔 Notifications ({count})" if count else "🔔 Notifications",
            danger=bool(count),
        )

    def _open_notifications(self) -> None:
        win = tk.Toplevel(self)
        bind_escape_to_close(win)
        win.title("Notifications")
        win.geometry("860x380")
        win.minsize(680, 280)
        win.resizable(True, True)
        # Keep this as a normal overlapped Windows window so the native
        # Minimize, Restore/Maximize, and Close buttons are available.
        win.overrideredirect(False)
        try:
            win.wm_attributes("-toolwindow", False)
        except tk.TclError:
            pass
        win.lift()
        win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        filter_bar = ttk.Frame(win); filter_bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        ttk.Label(filter_bar, text="Category:").pack(side="left")
        category_var = tk.StringVar(value="All")
        category_combo = ttk.Combobox(filter_bar, textvariable=category_var, values=("All",) + notifications.CATEGORIES, state="readonly", width=14)
        category_combo.pack(side="left", padx=6)
        tree = ttk.Treeview(win, columns=("severity", "category", "title", "page", "partita_colore", "message", "created"), show="headings")
        for col, title, width in (("severity", "Severity", 80), ("category", "Category", 100), ("title", "Title", 160), ("page", "Page", 110), ("partita_colore", "Partita Colore", 110), ("message", "Message", 270), ("created", "Created", 130)):
            tree.heading(col, text=title); tree.column(col, width=width, anchor="center" if col in {"severity", "page", "partita_colore", "created"} else "w")
        tree.grid(row=1, column=0, columnspan=5, sticky="nsew", padx=8, pady=8)
        tree.tag_configure("evenrow", background="#ffffff")
        tree.tag_configure("oddrow", background="#eef4fb")
        def refresh():
            tree.delete(*tree.get_children())
            for index, item in enumerate(notifications.list_open(category_var.get())):
                tree.insert(
                    "", "end", iid=item["key"],
                    values=(item.get("severity", ""), item.get("category", "Other"), item.get("title", ""), item.get("page", ""), item.get("partita_colore", ""), item.get("message", ""), item.get("created_at", "")),
                    tags=("oddrow" if index % 2 else "evenrow",),
                )
            self._refresh_notification_badge()
        def resolve_selected():
            for iid in tree.selection(): notifications.resolve(iid)
            refresh()
        def snooze_selected():
            for iid in tree.selection(): notifications.snooze(iid, 24)
            refresh()
        def export_excel():
            items = notifications.list_open()
            if not items:
                messagebox.showinfo("Notifications", "There are no open notifications to export.", parent=win)
                return
            path = filedialog.asksaveasfilename(
                parent=win, title="Export Notifications", defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")], initialfile="notifications.xlsx",
            )
            if not path:
                return
            try:
                from openpyxl import Workbook
                workbook = Workbook(); sheet = workbook.active; sheet.title = "Notifications"
                headers = ("Severity", "Category", "Title", "Page", "Partita Colore", "Message", "Created")
                sheet.append(headers)
                for item in items:
                    sheet.append((item.get("severity", ""), item.get("category", "Other"), item.get("title", ""), item.get("page", ""), item.get("partita_colore", ""), item.get("message", ""), item.get("created_at", "")))
                sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions
                for cell in sheet[1]: cell.font = cell.font.copy(bold=True)
                safe_save_workbook(workbook, path); workbook.close()
                messagebox.showinfo("Notifications", f"Export completed:\n{path}", parent=win)
            except Exception as exc:
                messagebox.showerror("Export error", str(exc), parent=win)
        category_combo.bind("<<ComboboxSelected>>", lambda _event: refresh())
        ttk.Button(win, text="Mark Resolved", command=resolve_selected).grid(row=2, column=0, sticky="w", padx=8, pady=(0, 8))
        ttk.Button(win, text="Snooze 24h", command=snooze_selected).grid(row=2, column=1, padx=8, pady=(0, 8))
        ttk.Button(win, text="Extract to Excel", command=export_excel).grid(row=2, column=2, padx=8, pady=(0, 8))
        ttk.Button(win, text="Refresh", command=refresh).grid(row=2, column=3, padx=8, pady=(0, 8))
        ttk.Button(win, text="Close", command=win.destroy).grid(row=2, column=4, sticky="e", padx=8, pady=(0, 8))
        refresh()

    def _create_backup(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Create Planning Backup", defaultextension=".zip",
            filetypes=[("Planning backup", "*.zip")], initialfile="planning_backup.zip",
        )
        if not path:
            return
        try:
            create_backup(path)
            messagebox.showinfo("Backup", f"Backup created successfully:\n{path}", parent=self)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Backup failed", str(exc), parent=self)

    def _restore_backup(self) -> None:
        path = filedialog.askopenfilename(
            title="Restore Planning Backup", filetypes=[("Planning backup", "*.zip")],
        )
        if not path:
            return
        if not messagebox.askyesno("Confirm Restore", "Restore this backup? Current settings and cached data may be replaced.", parent=self):
            return
        try:
            count = restore_backup(path)
            notifications.add("system-backup-restored", "Backup restored", f"Restored {count} file(s). Restart the program to reload all data.", "System", "medium")
            self._refresh_notification_badge()
            messagebox.showinfo("Restore", f"Restored {count} file(s). Please restart the program.", parent=self)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Restore failed", str(exc), parent=self)

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
        parent.rowconfigure(2, weight=1)

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

        action_row = ttk.Frame(entry_frame)
        action_row.grid(row=1, column=0, columnspan=4, sticky="e", pady=(6, 0))
        ttk.Button(
            action_row, text="💾  Save", command=self._on_elvy_save, width=10
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            action_row, text="✏  Edit Selected", command=self._on_elvy_edit
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            action_row, text="🗑  Delete Selected", command=self._on_elvy_delete
        ).pack(side="left")

        # ── Saved mappings table ─────────────────────────────────────────
        table_frame = ttk.LabelFrame(parent, text="Saved Mappings", padding=8)
        table_frame.grid(row=2, column=0, sticky="nsew", padx=4, pady=(4, 4))
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

        # Tracks the original Article No of the row being edited (if any),
        # so Save can detect a renamed key and remove the old entry instead
        # of leaving a stale duplicate behind.
        self._elvy_editing_key: str | None = None

        self._refresh_elvy_tree()

    # ------------------------------------------------------------------
    # Elvy mapping table handlers
    # ------------------------------------------------------------------

    def _refresh_elvy_tree(self) -> None:
        """Reload the mapping from disk and repopulate the table, sorted by
        whichever column/direction was last clicked, with zebra striping."""
        from parsers.elvy_mapping import load_elvy_mapping
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
        from parsers.elvy_mapping import add_elvy_mapping, delete_elvy_mapping
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
        from parsers.elvy_mapping import delete_elvy_mapping
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
        log_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(4, 12))
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
                                ordine_notebook: ttk.Notebook,
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
            self.rowconfigure(2, weight=0)
            self.rowconfigure(1, weight=5)
            notebook.configure(height=1)
        else:
            self._log_frame.grid()
            self.rowconfigure(1, weight=3)
            self.rowconfigure(2, weight=1)
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

        # po_filato_export_dir is new (Filato used to share the ERP folder)
        # -- fall back to the old shared folder on first run after
        # upgrading, so existing users keep working without reconfiguring.
        filato_dir_str = self._prefs.get("po_filato_export_dir") or erp_dir_str
        if filato_dir_str and Path(filato_dir_str).is_dir():
            self._po_filato_export_dir = Path(filato_dir_str)
            self._po_lbl_filato_dir.config(text=filato_dir_str, foreground="black")

        self._po_update_erp_file.set(self._prefs.get("po_update_erp_file", True))
        self._po_update_filato_file.set(self._prefs.get("po_update_filato_file", True))

        if getattr(self, "_situazione_tab", None):
            self.after_idle(self._situazione_tab.sync_shared_async)
            if hasattr(self._situazione_tab, "sync_remaining_shared_sources"):
                self.after_idle(self._situazione_tab.sync_remaining_shared_sources)
        if getattr(self, "_settimana_tab", None):
            self.after_idle(self._settimana_tab.sync_shared_async)
































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
