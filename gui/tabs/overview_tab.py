"""
overview_tab.py — Overview dashboard.

Read-only cross-tab summary, built entirely from data the app already has
in memory — nothing new to upload. Current sources:

  • Situazione Generale (situazione_tab.current_df) — columns include
    cliente, colore, titolo, partita, rocche, comment, new_comment.
      - "comment" starting with "PG-X"        -> yarn shortage (Filato)
      - "new_comment" containing "pronto da
         spedire"                              -> ready to ship, not yet
                                                   marked as shipped
                                                   (no Data Uscita yet, i.e.
                                                   no exit document made)
  • Magazino Filato (magazino_tab.magazino_summary) — columns articolo,
    partita, mag_rocche, mag_peso. articolo prefix identifies the client
    family (G130 = Elvy, G170 = Kamal) since Magazino has no Cliente column.

There's no historical/time-series table in situazione_db.py (partita_state
only tracks current state, not day-by-day snapshots), so the charts here
are current-snapshot breakdowns (by client / status) rather than true
trend-over-time lines. If a history table gets added later this can grow
a real trend chart on top of it.

Auto-refreshes when the tab is shown (call `on_shown()` from the notebook's
<<NotebookTabChanged>> handler) and has a manual Refresh button too.

Every Treeview on this page can export exactly what it's showing (after
search/filter) to an .xlsx file — right-click a row, or use the "Export to
Excel" button above each table. `export_treeview_to_excel()` is a small,
tree-agnostic helper: it only reads `tree["columns"]`, `tree.heading()`
and `tree.get_children()`, so it can be reused as-is on any other
ttk.Treeview in the app, not just the ones built here.
"""
from __future__ import annotations
from utility.excel_io import safe_save_workbook

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from utility.utils import load_settings, parse_number
from utility import notifications
import calculate.situazione as business_logic

ARTICOLO_CLIENT_LABELS = {"G130": "Elvy", "G170": "Kamal"}
READY_MARK = "pronto da spedire"
SHORTAGE_MARK = "PG-X"
DELIVERY_ALERT_DAYS = 3
QUALITY_READY_CODES = {"AA", "AC", "AU", "AT"}
HEADERS_IT = {
    "cliente": "Cliente", "colore": "Colore", "titolo": "Titolo",
    "articolo": "Articolo", "codice": "Codice", "ordine": "Ordine",
    "riga": "Riga", "data": "Data", "consegna": "Consegna",
    "delivery_date": "Delivery Date",
    "partita": "Partita", "rocche": "Rocche", "mc": "M/C",
    "comment": "Commento", "raw_yarn_match": "Filato Disponibile",
    "cq": "C.Q", "tinto": "Tinto", "bagno": "Bagno",
    "old_comment": "Vecchio commento", "new_comment": "Nuovo commento",
    "planedate": "PlaneDate", "data_qualita": "Data Qualità",
    "data_uscita": "Data Uscita", "custom": "Controllo",
    "days_in_qc": "Giorni in C.Q",
    "days_to_delivery": "Giorni alla consegna",
    "ritardo_consegna": "Ritardo (gg)",
    "prezzo": "Prezzo Ord.", "prezzo_lisini": "Prezzo Listini", "issue": "Problema",
}
CHECK_COLUMNS = [
    "cliente", "articolo", "titolo", "codice", "colore", "ordine", "riga",
    "data", "consegna", "partita", "rocche", "mc", "comment", "raw_yarn_match",
    "delivery_date",
    "cq", "tinto", "bagno", "old_comment", "new_comment", "planedate",
    "data_qualita", "data_uscita", "custom", "days_in_qc",
]


# ----------------------------------------------------------------------
# Generic "export any Treeview to Excel" helper — reusable elsewhere.
# ----------------------------------------------------------------------
def treeview_to_dataframe(tree: ttk.Treeview) -> pd.DataFrame:
    """Read a Treeview's current columns/rows (as displayed) into a DataFrame."""
    columns = tree["columns"]
    headers = [str(tree.heading(c)["text"] or c) for c in columns]
    rows = [tree.item(item)["values"] for item in tree.get_children()]
    return pd.DataFrame(rows, columns=headers)


def _infer_column_type_from_values(values) -> str:
    """Classify a column purely from its own data, not its header text:
    all-numeric -> "number", all-date-parseable -> "date", a mix of
    digits and letters/symbols -> "general" (Excel decides per cell,
    which is friendlier than forcing text on something like a Bagno code
    that is mostly digits with the odd letter), anything else -> "text"."""
    non_blank = [str(v).strip() for v in values if str(v).strip()]
    if not non_blank:
        return "text"

    def _is_number(token: str) -> bool:
        try:
            float(token.replace(",", "."))
            return True
        except ValueError:
            return False

    if all(_is_number(v) for v in non_blank):
        return "number"
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed_dates = pd.to_datetime(pd.Series(non_blank), errors="coerce", dayfirst=False)
    if parsed_dates.notna().all():
        return "date"
    has_digit = any(any(ch.isdigit() for ch in v) for v in non_blank)
    has_alpha = any(any(ch.isalpha() for ch in v) for v in non_blank)
    return "general" if (has_digit and has_alpha) else "text"


def export_dataframe_typed(df: pd.DataFrame, default_filename: str, parent: tk.Widget | None = None,
                            sheet_title: str = "Export") -> None:
    """Like export_treeview_to_excel, but classifies each column's Excel
    number format from its actual values (see _infer_column_type_from_values)
    instead of guessing from the header text."""
    if df.empty:
        messagebox.showinfo("No Data", "There is no data to export.", parent=parent)
        return
    path = filedialog.asksaveasfilename(
        title="Export to Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=default_filename,
    )
    if not path:
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        ws.append(list(df.columns))
        header_fills = {"date": "70AD47", "number": "5B9BD5", "general": "A5A5A5", "text": "ED7D31"}
        column_types = [_infer_column_type_from_values(df[col].tolist()) for col in df.columns]
        for row in df.itertuples(index=False, name=None):
            values = []
            for value, kind in zip(row, column_types):
                if value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == "":
                    values.append(None)
                elif kind == "date":
                    parsed = pd.to_datetime(value, errors="coerce")
                    values.append(str(value) if pd.isna(parsed) else parsed.strftime("%d/%m/%Y"))
                elif kind == "number":
                    parsed = parse_number(value)
                    values.append(parsed if parsed is not None else str(value))
                else:
                    values.append(str(value))
            ws.append(values)
        for index, (col, kind) in enumerate(zip(df.columns, column_types), start=1):
            letter = get_column_letter(index)
            header_cell = ws.cell(row=1, column=index)
            header_cell.font = openpyxl.styles.Font(color="FFFFFF", bold=True)
            header_cell.fill = openpyxl.styles.PatternFill("solid", fgColor=header_fills[kind])
            header_cell.alignment = openpyxl.styles.Alignment(horizontal="center")
            if kind == "date":
                # Exported as text on purpose: Excel can otherwise reinterpret
                # day/month order based on the computer's locale.
                for cell in ws[letter][1:]:
                    cell.number_format = "@"
            elif kind == "number":
                for cell in ws[letter][1:]:
                    cell.number_format = "#,##0"
            elif kind == "general":
                for cell in ws[letter][1:]:
                    cell.number_format = "General"
            max_len = max(
                [len(str(col))]
                + [len(str(ws.cell(row=r, column=index).value or "")) for r in range(2, ws.max_row + 1)]
            )
            ws.column_dimensions[letter].width = max(10, min(45, max_len + 2))
        last_col = get_column_letter(ws.max_column)
        last_row = ws.max_row
        table = Table(displayName="TypedExport", ref=f"A1:{last_col}{last_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(table)
        ws.freeze_panes = "A2"
        safe_save_workbook(wb, path)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Export Failed", str(exc), parent=parent)
        return
    messagebox.showinfo("Export Completed", f"Exported to:\n{path}", parent=parent)


def export_treeview_to_excel(tree: ttk.Treeview, default_filename: str, parent: tk.Widget | None = None) -> None:
    """Export a Treeview's currently displayed rows/columns to a new .xlsx file."""
    df = treeview_to_dataframe(tree)
    if df.empty:
        messagebox.showinfo("No Data", "There is no data to export.", parent=parent)
        return
    path = filedialog.asksaveasfilename(
        title="Export to Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=default_filename,
    )
    if not path:
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Overview"
        _write_typed_excel_table(ws, df)
        ws.freeze_panes = "A2"
        safe_save_workbook(wb, path)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Export Failed", str(exc), parent=parent)
        return
    messagebox.showinfo("Export Completed", f"Exported to:\n{path}", parent=parent)


def _write_typed_excel_table(ws, df: pd.DataFrame) -> None:
    """Write an Overview export with typed cells, styling and Excel filters."""
    ws.append(list(df.columns))

    date_headers = {"data", "consegna", "data qualità", "data uscita", "planedate"}
    number_headers = {
        "ordine", "riga", "rocche", "m/c",
        "giorni in c.q", "giorni alla consegna", "extra %",
    }
    general_headers = {"cliente", "codice", "partita", "titolo", "colore"}
    header_fills = {
        "date": "70AD47",    # green
        "number": "5B9BD5",  # blue
        "general": "A5A5A5",  # gray
        "text": "ED7D31",    # orange
    }

    column_types = []
    for col in df.columns:
        header = str(col).strip().casefold()
        if header in date_headers or "data" in header or "date" in header:
            column_types.append("date")
        elif header in general_headers:
            column_types.append("general")
        elif header in number_headers or "giorni" in header:
            column_types.append("number")
        else:
            column_types.append("text")

    for row in df.itertuples(index=False, name=None):
        values = []
        for value, kind in zip(row, column_types):
            if value is None or (isinstance(value, float) and pd.isna(value)):
                values.append(None)
            elif kind == "date":
                parsed = pd.to_datetime(value, errors="coerce")
                values.append(
                    str(value) if pd.isna(parsed) else parsed.strftime("%d/%m/%Y")
                )
            elif kind == "number":
                parsed = parse_number(value)
                values.append(parsed if parsed is not None else str(value))
            else:
                values.append(str(value))
        ws.append(values)

    for index, (col, kind) in enumerate(zip(df.columns, column_types), start=1):
        letter = get_column_letter(index)
        header_cell = ws.cell(row=1, column=index)
        header_cell.font = openpyxl.styles.Font(color="FFFFFF", bold=True)
        header_cell.fill = openpyxl.styles.PatternFill(
            "solid", fgColor=header_fills[kind]
        )
        header_cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        if kind == "date":
            # Dates are exported as text deliberately. Excel otherwise may
            # reinterpret the day/month order using the computer locale.
            for cell in ws[letter][1:]:
                cell.number_format = "@"
        elif kind == "number":
            for cell in ws[letter][1:]:
                # Overview quantities and identifiers are whole numbers:
                # show 6, not 6.00, while keeping the cells numeric.
                cell.number_format = "#,##0"
        max_len = max(
            [len(str(col))]
            + [len(str(ws.cell(row=row, column=index).value or "")) for row in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[letter].width = max(10, min(45, max_len + 2))

    last_col = get_column_letter(ws.max_column)
    last_row = ws.max_row
    table = Table(displayName="OverviewExport", ref=f"A1:{last_col}{last_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(table)


def attach_excel_export(tree: ttk.Treeview, default_filename: str) -> None:
    """Right-click on any row of *tree* -> "Export to Excel"."""
    menu = tk.Menu(tree, tearoff=0)
    menu.add_command(
        label="📤 Export to Excel",
        command=lambda: export_treeview_to_excel(tree, default_filename, parent=tree),
    )

    def _popup(event):
        row_id = tree.identify_row(event.y)
        if row_id:
            tree.selection_set(row_id)
        menu.tk_popup(event.x_root, event.y_root)

    tree.bind("<Button-3>", _popup)


def _format_display_dates(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Format report dates for people while keeping blanks blank."""
    result = df.copy()
    for column in columns:
        if column not in result.columns:
            continue
        original = result[column].fillna("").astype(str)
        parsed = pd.to_datetime(result[column], errors="coerce", dayfirst=False)
        formatted = parsed.dt.strftime("%d/%m/%Y")
        result[column] = formatted.where(parsed.notna(), original)
    return result


class OverviewTab(ttk.Frame):
    """Dashboard: KPI cards + charts + detail lists, all read-only."""

    AUTO_REFRESH_MS = 5000
    CARD_GRID_COLUMNS = 20

    def __init__(self, master, situazione_tab, magazino_tab, biglietti_tab=None, prezzi_tab=None, save_prefs=None, prefs=None, on_shared_cache_changed=None, settimana_tab=None):
        super().__init__(master)
        self.situazione_tab = situazione_tab
        self.magazino_tab = magazino_tab
        self.biglietti_tab = biglietti_tab
        self.prezzi_tab = prezzi_tab
        self.settimana_tab = settimana_tab
        self._save_prefs = save_prefs
        self._prefs = prefs or {}
        self._on_shared_cache_changed = on_shared_cache_changed
        p_dir = self._prefs.get("master_data_dir")
        self._data_folder = Path(p_dir) if p_dir and Path(p_dir).is_dir() else None

        self._data_signature = None
        self._auto_refresh_id = None
        self._build_ui()
        self.refresh(force=True)
        self._schedule_auto_refresh()

    # ------------------------------------------------------------------
    # UI scaffolding
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.configure("Overview.Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("Overview.CardTitle.TLabel", background="#ffffff", foreground="#526173",
                        font=("Segoe UI", 10, "bold"))
        style.configure("Overview.CardValue.TLabel", background="#ffffff", foreground="#16324f",
                        font=("Segoe UI", 24, "bold"))
        style.configure("Overview.Table.TLabelframe", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("Overview.Table.TLabelframe.Label", background="#ffffff", foreground="#16324f",
                        font=("Segoe UI", 10, "bold"))
        style.configure("Overview.Treeview", background="#ffffff", fieldbackground="#ffffff",
                        foreground="#25364d", rowheight=28, font=("Segoe UI", 9))
        style.configure("Overview.Treeview.Heading", background="#16324f", foreground="#ffffff",
                        font=("Segoe UI", 9, "bold"), padding=(7, 6))
        style.map("Overview.Treeview", background=[("selected", "#2f80ed")],
                  foreground=[("selected", "#ffffff")])
        style.configure("Overview.AlertCard.TFrame", background="#fff7ed", relief="solid", borderwidth=1)
        style.configure("Overview.AlertCardTitle.TLabel", background="#fff7ed", foreground="#c2410c",
                        font=("Segoe UI", 10, "bold"))
        style.configure("Overview.AlertCardValue.TLabel", background="#fff7ed", foreground="#9a3412",
                        font=("Segoe UI", 24, "bold"))
        style.configure("Overview.DangerCardValue.TLabel", background="#fff7ed", foreground="#c62828",
                font=("Segoe UI", 24, "bold"))

        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", padx=12, pady=(10, 4))
        ttk.Label(
            toolbar, text="📊  Operations Overview", foreground="#0f172a",
            font=("Segoe UI", 14, "bold")
        ).pack(side="left")
        self._lbl_updated = ttk.Label(toolbar, text="", foreground="#64748b", font=("Segoe UI", 9))
        self._lbl_updated.pack(side="right")

        self._notification_frame = None

        outer = ttk.Frame(self)
        outer.pack(side="top", fill="both", expand=True, padx=8, pady=(0, 8))

        canvas = tk.Canvas(outer, highlightthickness=0, bg="#f8fafc")
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._body = ttk.Frame(canvas, padding=(8, 4))
        self._body.columnconfigure(0, weight=1)
        body_window = canvas.create_window((0, 0), window=self._body, anchor="nw")

        def _on_body_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(body_window, width=event.width)

        self._body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        # ── Master Data Synchronization Card ──
        sync_frame = ttk.LabelFrame(self._body, text=" 📂 Master Data Synchronization (Update All Data) ", padding=10)
        sync_frame.pack(side="top", fill="x", padx=4, pady=(2, 8))
        sync_frame.columnconfigure(1, weight=1)

        ttk.Button(sync_frame, text="📂  Upload All Data from Folder...", command=self._on_choose_data_folder, width=28).grid(row=0, column=0, sticky="w", pady=2)

        folder_text = str(self._data_folder) if self._data_folder else "No data folder selected"
        folder_color = "#111827" if self._data_folder else "grey"
        self._lbl_folder = ttk.Label(sync_frame, text=folder_text, foreground=folder_color, anchor="w")
        self._lbl_folder.grid(row=0, column=1, sticky="ew", padx=(10, 10), pady=2)

        self._btn_sync = ttk.Button(sync_frame, text="🔄  Update All Data", command=self._on_sync_all_data, width=20)
        self._btn_sync.grid(row=0, column=2, sticky="e", pady=2)

        sec_row = ttk.Frame(sync_frame)
        sec_row.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        sec_row.columnconfigure(0, weight=1)

        self._lbl_sync_status = ttk.Label(sec_row, text="● Ready to sync", foreground="#2E7D32", font=("Segoe UI", 9))
        self._lbl_sync_status.grid(row=0, column=1, sticky="e", padx=(10, 0))

        self._cards_frame = ttk.Frame(self._body, padding=(0, 2))
        self._cards_frame.pack(side="top", fill="x", padx=2, pady=(2, 8))
        for column in range(self.CARD_GRID_COLUMNS):
            self._cards_frame.columnconfigure(column, weight=1, uniform="overview_card")

        self._tables_frame = ttk.Frame(self._body, padding=(0, 2))
        self._table_count = 0

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def on_shown(self) -> None:
        """Call this when the tab becomes visible — keeps it auto-fresh."""
        self._schedule_auto_refresh()
        if self._data_signature != self._current_data_signature():
            self.refresh()

    def _schedule_auto_refresh(self) -> None:
        if self._auto_refresh_id is None and self.winfo_exists():
            self._auto_refresh_id = self.after(self.AUTO_REFRESH_MS, self._auto_refresh_tick)

    def _auto_refresh_tick(self) -> None:
        self._auto_refresh_id = None
        if not self.winfo_exists():
            return
        # Do not touch the Treeviews while the user is on another page.
        if self.winfo_ismapped() and self._data_signature != self._current_data_signature():
            self.refresh()
        self._schedule_auto_refresh()

    def _current_data_signature(self):
        """Return cheap revision counters maintained by the source tabs."""
        prezzi_df = getattr(self.prezzi_tab, "_base_df", None)
        return (
            getattr(self.situazione_tab, "_data_revision", 0),
            getattr(self.situazione_tab, "_copertura_revision", 0),
            getattr(self.magazino_tab, "_data_revision", 0),
            getattr(self.prezzi_tab, "_loaded_source_path", ""),
            len(prezzi_df) if isinstance(prezzi_df, pd.DataFrame) else 0,
            len(notifications.list_open()),
        )

    def _on_choose_data_folder(self) -> None:
        initial = str(self._data_folder) if self._data_folder else None
        path = filedialog.askdirectory(
            title="Select Data Folder containing factory files (Articoli, DFM, WINCOINT, etc.)",
            initialdir=initial,
        )
        if not path:
            return
        self._data_folder = Path(path).expanduser().resolve()
        self._lbl_folder.config(text=str(self._data_folder), foreground="#111827")
        # Persist immediately in the same settings store used at startup.
        # The absolute path also avoids a changed working directory making a
        # valid selected folder look missing on the next launch.
        self._prefs["master_data_dir"] = str(self._data_folder)
        if self._save_prefs:
            self._save_prefs(master_data_dir=str(self._data_folder))
        # Selecting the folder only stores the master directory. The update
        # button is the only action that reads and distributes its files.
        self._lbl_sync_status.config(text="● Folder saved — press Update All Data", foreground="#1565C0")

    def _on_sync_all_data(self) -> None:
        from tkinter import messagebox
        # Reload the persisted value on every update click. This covers the
        # case where Overview was rebuilt or another tab replaced its prefs
        # mapping after the folder was selected.
        saved_folder = load_settings().get("master_data_dir")
        if saved_folder:
            candidate = Path(str(saved_folder)).expanduser()
            if candidate.is_dir():
                self._data_folder = candidate.resolve()
                self._lbl_folder.config(text=str(self._data_folder), foreground="#111827")
        if not self._data_folder or not self._data_folder.is_dir():
            messagebox.showwarning(
                "Data folder not selected",
                "Select the master data folder first with Upload All Data from Folder.",
            )
            return

        self._btn_sync.config(state="disabled")
        self._lbl_sync_status.config(text="⏳ Synchronizing all factory files in background...", foreground="#1565C0")

        import threading
        import pipelines.master_import as master_import

        def worker():
            try:
                loaded, skipped = master_import.import_master_directory(
                    self._data_folder,
                    self.situazione_tab,
                    self.magazino_tab,
                    self.biglietti_tab,
                    self.prezzi_tab,
                    # Order input files belong to Create (EXCEL+Biglietti)
                    # and must never be auto-selected by Overview.
                    skip_keys={"data_ordine", "dispo_bagno"},
                    settimana_tab=self.settimana_tab,
                )
                err = None
            except Exception as exc:  # noqa: BLE001
                loaded, skipped, err = [], [], str(exc)

            def apply_result():
                self._btn_sync.config(state="normal")
                if err:
                    self._lbl_sync_status.config(text=f"✖ Sync error: {err}", foreground="#C62828")
                    messagebox.showerror("Sync Failed", f"Failed to sync data folder:\n{err}")
                    return

                now_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                self._lbl_sync_status.config(text=f"✔ Synced ({len(loaded)} sources) at {now_str}", foreground="#2E7D32")
                self._lbl_updated.config(text=f"Last sync: {now_str}")

                # The individual _handle_upload/_on_upload_magazino calls
                # inside import_master_directory populate each tab's
                # in-memory frames and caches, but Situazione's own table
                # (compute_situation + its Treeview) only recomputes when
                # its own Refresh runs -- do that now too, silently, if
                # every source it needs actually came through this sync.
                try:
                    st = self.situazione_tab
                    required = {"wincoint", "dfm", "data_prod", "copertura", "uscita", "qualita"}
                    if st is not None and required.issubset(getattr(st, "loaded_frames", {}).keys()):
                        st._on_refresh()
                except Exception:
                    pass

                # Propagate to every other tab that keeps its own live
                # in-memory copy of a shared source (Kamal/Ordine MED/
                # Magazino Filato/Situazione Settimanale/Ordine Elvy/Prezzi)
                # -- otherwise they'd only pick up what this sync just wrote
                # next time the user happens to interact with them.
                if self._on_shared_cache_changed:
                    try:
                        self._on_shared_cache_changed()
                    except Exception:
                        pass
                    # Magazino and the shared DFM/Produzione consumers finish
                    # their own parsing in workers. Re-run the fan-out after
                    # those caches have had time to become available.
                    self.after(1000, self._refresh_shared_tabs)

                msg_parts = []
                if loaded:
                    msg_parts.append("✅ Successfully Loaded & Distributed:\n" + "\n".join(f"  • {item}" for item in loaded))
                if skipped:
                    msg_parts.append("\n⚠️ Missing / Not Found in Folder:\n" + "\n".join(f"  • {item}" for item in skipped))

                summary = "\n".join(msg_parts) if msg_parts else "No recognized files found in folder."
                messagebox.showinfo("Data Synchronization Complete", summary)
                self.refresh(force=True)

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_shared_tabs(self, attempt: int = 0) -> None:
        """Retry shared-source fan-out after background upload handlers finish."""
        if not self.winfo_exists():
            return
        if self._on_shared_cache_changed:
            try:
                self._on_shared_cache_changed()
            except Exception:
                pass
        if attempt < 7:
            self.after(1000, lambda: self._refresh_shared_tabs(attempt + 1))

    def refresh(self, force: bool = False) -> None:
        signature = self._current_data_signature()
        if not force and self._data_signature == signature:
            return
        df = getattr(self.situazione_tab, "current_df", None)
        magazino = getattr(self.magazino_tab, "magazino_summary", None)
        df = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
        magazino = magazino.copy() if isinstance(magazino, pd.DataFrame) else pd.DataFrame()
        self._render(df, magazino)
        self._data_signature = signature
        self._lbl_updated.config(text=f"Ultimo aggiornamento: {pd.Timestamp.now():%Y-%m-%d %H:%M:%S}")

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------
    def _render_notifications(self) -> None:
        if not getattr(self, "_notification_frame", None):
            return
        for child in self._notification_frame.winfo_children():
            child.destroy()
        items = notifications.list_open()
        if not items:
            ttk.Label(self._notification_frame, text="No open notifications", foreground="#2e7d32").pack(anchor="w")
            return
        for item in items[:8]:
            row = ttk.Frame(self._notification_frame)
            row.pack(fill="x", pady=1)
            color = "#c62828" if item.get("severity") == "high" else "#9a3412"
            ttk.Label(row, text=f"{item.get('title', 'Notification')}: ", foreground=color, font=("Segoe UI", 9, "bold")).pack(side="left")
            ttk.Label(row, text=item.get("message", ""), foreground="#344054").pack(side="left", fill="x", expand=True)
            key = item.get("key", "")
            ttk.Button(row, text="Resolve", width=9, command=lambda key=key: self._resolve_notification(key)).pack(side="right")

    def _resolve_notification(self, key: str) -> None:
        notifications.resolve(key)
        self.refresh(force=True)

    def _render(self, df: pd.DataFrame, magazino: pd.DataFrame) -> None:
        self._render_notifications()
        for frame in (self._cards_frame, self._tables_frame):
            for widget in frame.winfo_children():
                widget.destroy()
        self._has_rendered = True
        self._table_count = 0

        if not df.empty:
            for col in ("cliente", "colore", "titolo", "comment", "new_comment", "custom"):
                values = df[col] if col in df.columns else pd.Series("", index=df.index)
                df[col] = values.fillna("").astype(str).str.strip()

        # Keep every shortage row whose Commento contains PG-X.  The raw-yarn
        # match is only an informational column; it must never filter out a
        # shortage that has no available yarn match.
        shortage_df = df.iloc[0:0].copy()
        if not df.empty and "comment" in df.columns:
            shortage_mask = df["comment"].astype(str).str.contains(
                SHORTAGE_MARK, case=False, regex=False, na=False
            )
            shortage_df = df.loc[shortage_mask].copy()
            if "raw_yarn_match" not in shortage_df.columns:
                shortage_df["raw_yarn_match"] = ""
        # "Colori pronti" means: quality is completed with a ready code,
        # more than three days have passed since the quality date, and no
        # warehouse exit/invoice has been recorded yet.  Do not rely only on
        # New Comment, because that status can be carried from old history.
        ready_df = df.iloc[0:0].copy()
        if not df.empty:
            cq_ready = df.get(
                "cq", pd.Series("", index=df.index)
            ).astype(str).str.strip().str.upper().isin(QUALITY_READY_CODES)
            quality_dates = pd.to_datetime(
                df.get("data_qualita", pd.Series(index=df.index)), errors="coerce"
            )
            quality_older_than_3_days = quality_dates.notna() & (
                (pd.Timestamp.now().normalize() - quality_dates.dt.normalize()).dt.days > 3
            )
            no_invoice = pd.to_datetime(
                df.get("data_uscita", pd.Series(index=df.index)), errors="coerce"
            ).isna()
            ready_df = df.loc[cq_ready & quality_older_than_3_days & no_invoice].copy()
        # Quality-delay alert is intentionally based on the three visible
        # report fields, so stale Custom values cannot leak old rows into it.
        check_df = df.iloc[0:0].copy()
        if not df.empty:
            cq_oo = df.get("cq", pd.Series("", index=df.index)).astype(str).str.strip().str.upper().eq("OO")
            new_comment_cq = df.get("new_comment", pd.Series("", index=df.index)).astype(str).str.strip().str.casefold().eq("c.q")
            days_in_qc = pd.to_numeric(df.get("days_in_qc", pd.Series(index=df.index)), errors="coerce")
            check_df = df.loc[cq_oo & new_comment_cq & days_in_qc.gt(4)].copy()

        # Delivery alert: include overdue orders and orders due within the
        # next DELIVERY_ALERT_DAYS days, while excluding already shipped rows.
        # The current situation data stores dates as ISO strings, so parse them
        # explicitly instead of comparing display text.
        delivery_df = df.iloc[0:0].copy()
        if not df.empty:
            is_elvy = df.get("cliente", pd.Series("", index=df.index)).astype(str).str.upper().isin({"ELVY", "3009", "3009.0"})
            derived_dates = pd.to_datetime(df.get("delivery_date", pd.Series(index=df.index)), errors="coerce")
            original_dates = pd.to_datetime(df.get("consegna", pd.Series(index=df.index)), errors="coerce")
            due_dates = original_dates.where(~is_elvy, derived_dates)
            today = pd.Timestamp.now().normalize()
            deadline = today + pd.Timedelta(days=DELIVERY_ALERT_DAYS)
            # Delivery risk is limited to batches still queued in quality
            # (C.Q = OO), as requested. Include overdue and due-within-window
            # rows, but never rows already shipped.
            not_shipped = ~df.get("new_comment", pd.Series("", index=df.index)).str.casefold().eq("spedita")
            still_in_quality = df.get("cq", pd.Series("", index=df.index)).astype(str).str.strip().eq("OO")
            due_mask = due_dates.notna() & due_dates.le(deadline) & not_shipped & still_in_quality
            delivery_df = df.loc[due_mask].copy()
            delivery_df["days_to_delivery"] = (due_dates.loc[due_mask] - today).dt.days
            delivery_df = delivery_df.sort_values("days_to_delivery")

        n_shortage_colors = shortage_df["colore"].nunique() if not shortage_df.empty else 0
        n_ready_colors = ready_df["colore"].nunique() if not ready_df.empty else 0
        elvy_mask = df.get("cliente", pd.Series("", index=df.index)).astype(str).str.upper().isin({"ELVY", "3009", "3009.0"})
        elvy_df = df.loc[elvy_mask].copy()
        elvy_delivery = pd.to_datetime(elvy_df.get("delivery_date", pd.Series(index=elvy_df.index)), errors="coerce")
        elvy_total_colors = elvy_df["colore"].nunique() if not elvy_df.empty and "colore" in elvy_df else 0
        elvy_due_colors = int((elvy_delivery.notna() & elvy_delivery.le(pd.Timestamp.now().normalize() + pd.Timedelta(days=DELIVERY_ALERT_DAYS))).sum())
        total_yarn_kg = (
            magazino["mag_peso"].sum() if not magazino.empty and "mag_peso" in magazino else 0.0
        )

        totale_cols = ["cliente", "articolo", "titolo", "codice", "colore", "partita", "rocche",
                       "mc", "cq", "bagno", "new_comment", "delivery_date"]
        self._add_card(0, "📋 Totale partite", f"{len(df):,}",
                        on_click=lambda: self._open_data_ticket("Totale partite", df, totale_cols, "totale_partite.xlsx"))
        magazino_cols = ["articolo", "titolo", "partita", "mag_rocche", "mag_peso"]
        self._add_card(1, "🧵 Filato disponibile", f"{total_yarn_kg:,.0f} kg",
                        on_click=lambda: self._open_data_ticket("Filato disponibile", magazino, magazino_cols, "filato_disponibile.xlsx"))
        ready_cols = ["cliente", "codice", "colore", "titolo", "prezzo", "prezzo_lisini", "partita", "rocche", "mc", "bagno"]
        self._add_card(2, "📦 Pronte da spedire", str(n_ready_colors),
                        on_click=lambda: self._open_data_ticket("Colori pronti da spedire (senza uscita)", ready_df, ready_cols, "colori_pronti.xlsx"))
        shortage_cols = ["cliente", "codice", "colore", "titolo", "ordine", "riga", "partita", "rocche", "comment", "raw_yarn_match"]
        self._add_card(3, "⚠️ In attesa di filato", str(n_shortage_colors),
                        on_click=lambda: self._open_data_ticket("Colori in attesa di filato", shortage_df, shortage_cols, "colori_filato_mancante.xlsx"))
        check_cols = ["cliente", "codice", "colore", "titolo", "partita", "rocche", "days_in_qc", "new_comment"]
        self._add_card(4, "⚠️ Ritardo in Q.C", f"{len(check_df):,}", alert=True,
                        on_click=lambda: self._open_data_ticket("Ritardo in Q.C (C.Q = OO, oltre 4 giorni)", check_df, check_cols, "ritardo_in_qc.xlsx"))
        delivery_cols = ["cliente", "codice", "colore", "titolo", "partita", "rocche", "delivery_date", "days_to_delivery", "new_comment"]
        self._add_card(5, "📅 Ritardo Consegna", f"{len(delivery_df):,}", alert=True,
                        on_click=lambda: self._open_data_ticket(f"Ritardo consegna / entro {DELIVERY_ALERT_DAYS} giorni (C.Q = OO)", delivery_df, delivery_cols, "ordini_urgenti_consegna.xlsx"))
        elvy_cols = ["articolo", "titolo", "codice", "colore", "prezzo", "ordine", "riga", "data",
                     "partita", "rocche", "mc", "comment", "cq", "bagno", "new_comment", "delivery_date"]
        elvy_due_mask = elvy_delivery.notna() & elvy_delivery.le(pd.Timestamp.now().normalize() + pd.Timedelta(days=DELIVERY_ALERT_DAYS))
        elvy_due_df = elvy_df.loc[elvy_due_mask]
        self._add_card(6, "🎨 Elvy: totale colori", f"{elvy_total_colors:,}",
                        on_click=lambda: self._open_data_ticket("Tutti i colori Elvy", elvy_df, elvy_cols, "colori_elvy.xlsx"))
        self._add_card(7, f"📅 Elvy: entro {DELIVERY_ALERT_DAYS} giorni", f"{elvy_due_colors:,}", alert=elvy_due_colors > 0,
                        on_click=lambda: self._open_data_ticket(f"Elvy: consegna entro {DELIVERY_ALERT_DAYS} giorni", elvy_due_df, elvy_cols, "elvy_entro_giorni.xlsx"))

        price_anomalies = business_logic.find_price_anomalies(df)
        price_problem_colors = len({
            str(item.get("colore", "")).strip()
            for item in price_anomalies
            if str(item.get("colore", "")).strip()
        })
        price_anomalies_cols = ["cliente", "articolo", "codice", "colore", "prezzo", "prezzo_lisini", "ordine", "riga", "bagno", "mc", "issue"]
        self._add_card(
            8, "💲 Errori Prezzo — colori", str(price_problem_colors),
            alert=price_problem_colors > 0,
            on_click=lambda: self._open_data_ticket(
                "Errori Prezzo (Prezzo mancante o sospetto)", pd.DataFrame(price_anomalies),
                price_anomalies_cols, "errori_prezzo.xlsx",
            ),
        )
        price_change_df = self._price_change_df()
        price_change_cols = ["CLARTICOLO", "CLCOLORE", "CLDESCR", "old_price", "new_price", "pct_change", "changed_on"]
        self._add_card(
            9, "⚠ Price Changes", f"{len(price_change_df):,}",
            alert=len(price_change_df) > 0, danger=len(price_change_df) > 0,
            on_click=lambda: self._open_data_ticket("Price Changes (Listini)", price_change_df, price_change_cols, "price_changes.xlsx"),
        )

        self._add_machine_summary(df)

    def _add_machine_summary(self, situation_df: pd.DataFrame) -> None:
        """Render clickable Copertura machine totals as schedule cards."""
        copertura = getattr(self.situazione_tab, "loaded_frames", {}).get("copertura")
        schedule = business_logic.build_machine_schedule(situation_df, copertura)
        counts = {machine: 0 for machine in range(3, 13)}
        if not schedule.empty:
            counts.update(schedule["machine"].value_counts().to_dict())

        frame = tk.LabelFrame(
            self._cards_frame, text="Copertura macchine", bg="#f8fafc", fg="#16324f",
            font=("Segoe UI", 10, "bold"), padx=8, pady=6,
        )
        frame.grid(
            row=2, column=0, columnspan=self.CARD_GRID_COLUMNS,
            padx=2, pady=(0, 10), sticky="ew",
        )
        for column in range(self.CARD_GRID_COLUMNS):
            self._cards_frame.columnconfigure(column, weight=1, uniform="overview_card")

        covered_until = business_logic.machine_coverage_until

        for index, machine in enumerate(range(3, 13)):
            count = int(counts[machine])
            empty = count == 0
            card = tk.Frame(
                frame, bg="#fee2e2" if empty else "#eaf2f8",
                highlightbackground="#ef4444" if empty else "#cbd5e1",
                highlightthickness=1, width=110, height=82,
            )
            card.grid(row=0, column=index, padx=3, pady=3, sticky="ew")
            card.grid_propagate(False)
            frame.columnconfigure(index, weight=1, minsize=106)
            click = lambda _event, selected_machine=machine: self._open_machine_schedule(selected_machine, schedule)
            card.bind("<Button-1>", click)
            tk.Label(card, text=f"M{machine}", bg=card["bg"], fg="#991b1b" if empty else "#16324f",
                     font=("Segoe UI", 10, "bold"), anchor="center").pack(fill="x")
            for child in card.winfo_children():
                child.bind("<Button-1>", click)
            tk.Label(card, text=f"{count} colori", bg=card["bg"], fg="#991b1b" if empty else "#344054",
                     font=("Segoe UI", 12, "bold"), anchor="center").pack(fill="x")
            tk.Label(card, text=f"Fino al: {covered_until(count)}", bg=card["bg"],
                     fg="#991b1b" if empty else "#667085",
                     font=("Segoe UI", 9, "bold"), anchor="center").pack(fill="x", pady=(3, 0))

            for child in card.winfo_children():
                child.bind("<Button-1>", click)

    def _open_machine_schedule(self, machine: int, schedule: pd.DataFrame) -> None:
        """Show the Copertura-ordered dyeing sequence for one machine."""
        window = tk.Toplevel(self)
        window.title(f"Copertura macchine — M{machine}")
        window.geometry("1180x620")
        window.minsize(850, 420)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text=f"Dyeing schedule for Machine {machine} — 2 colors/day; Friday off", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=8)
        toolbar = ttk.Frame(window)
        toolbar.grid(row=0, column=0, sticky="e", padx=10, pady=8)
        search = tk.StringVar()
        ttk.Label(toolbar, text="Search:").pack(side="left")
        entry = ttk.Entry(toolbar, textvariable=search, width=28)
        entry.pack(side="left", padx=5)
        ttk.Button(toolbar, text="Clear", command=lambda: search.set("")).pack(side="left", padx=(5, 0))
        # Same field order as Situazione Generale's own COLUMN_SPEC.
        columns = ["cliente", "articolo", "titolo", "colore", "ordine", "riga",
                   "partita", "rocche", "machine", "dye_date", "bagno"]
        labels = {"dye_date": "Data tintura", "machine": "M/C", "bagno": "Bagno", "colore": "Colore",
                   "titolo": "Titolo", "articolo": "Articolo", "partita": "Partita", "rocche": "Rocche",
                   "cliente": "Cliente", "ordine": "Ordine", "riga": "Riga"}
        tree = ttk.Treeview(window, columns=columns, show="headings")
        sort_state: dict[str, bool] = {}

        def sort_col(col):
            nonlocal machine_df
            if machine_df.empty:
                return
            ascending = sort_state.get(col, True)
            machine_df = machine_df.sort_values(by=col, ascending=ascending, key=lambda s: s.astype(str))
            sort_state[col] = not ascending
            render()

        for col in columns:
            tree.heading(col, text=labels[col], command=lambda c=col: sort_col(c))
            tree.column(col, width=115 if col not in {"titolo", "colore"} else 170, anchor="center")
        tree.grid(row=1, column=0, sticky="nsew", padx=10)
        scroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)

        # Select by column NAME (not position) so the Treeview's column
        # order can never drift out of sync with the row values again --
        # that mismatch was exactly what swapped M/C and Data tintura before.
        machine_df = schedule[schedule["machine"] == machine][columns].copy() if not schedule.empty else schedule.reindex(columns=columns)

        def render(*_args):
            tree.delete(*tree.get_children())
            term = search.get().strip().casefold()
            for row in machine_df.itertuples(index=False, name=None):
                values = tuple("" if value is None else value for value in row)
                if not term or term in " ".join(str(value).casefold() for value in values):
                    tree.insert("", "end", values=values)
        search.trace_add("write", render)
        render()
        buttons = ttk.Frame(window, padding=10)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Button(
            buttons, text="Export Excel",
            command=lambda: export_dataframe_typed(
                machine_df.rename(columns=labels), f"machine_{machine}_dyeing_schedule.xlsx", window,
                sheet_title=f"M{machine}",
            ),
        ).pack(side="left")
        ttk.Button(buttons, text="Close", command=window.destroy).pack(side="right")

    def _price_change_df(self) -> pd.DataFrame:
        """Listini price transitions over the alert threshold (see detect_price_anomalies)."""
        prezzi_tab = self.prezzi_tab
        base_df = getattr(prezzi_tab, "_base_df", None)
        if not isinstance(base_df, pd.DataFrame) or base_df.empty:
            return pd.DataFrame(columns=["CLARTICOLO", "CLCOLORE", "CLDESCR", "old_price", "new_price", "pct_change", "changed_on"])
        from calculate.prezzi import detect_price_anomalies
        return detect_price_anomalies(base_df, min_pct_change=10.0)

    def _add_card(self, col: int, title: str, value: str, alert: bool = False, danger: bool = False,
                  on_click=None) -> None:
        card_style = "Overview.AlertCard.TFrame" if alert else "Overview.Card.TFrame"
        title_style = "Overview.AlertCardTitle.TLabel" if alert else "Overview.CardTitle.TLabel"
        value_style = "Overview.DangerCardValue.TLabel" if danger else ("Overview.AlertCardValue.TLabel" if alert else "Overview.CardValue.TLabel")
        card = ttk.Frame(self._cards_frame, style=card_style, padding=(14, 12))
        grid_row = col // 5
        item_in_row = col % 5
        # Keep five equal-width cards on every row, including the row that
        # contains Price Changes.
        span = 4
        grid_column = item_in_row * span
        card.grid(
            row=grid_row, column=grid_column, columnspan=span,
            padx=3, pady=4, sticky="nsew",
        )
        self._cards_frame.rowconfigure(grid_row, weight=1, minsize=100)
        title_label = ttk.Label(card, text=title, style=title_style, wraplength=150)
        title_label.pack(anchor="w")
        value_label = ttk.Label(card, text=value, style=value_style)
        value_label.pack(anchor="w", pady=(4, 0))
        if on_click is not None:
            for widget in (card, title_label, value_label):
                widget.configure(cursor="hand2")
                widget.bind("<Button-1>", lambda _event, cb=on_click: cb())

    def _open_data_ticket(self, title: str, df: pd.DataFrame, cols: list[str], export_filename: str) -> None:
        """Generic ticket popup for an Overview card: search/clear toolbar,
        a Treeview in the same column order given, and an Export Excel that
        classifies each column's format from its own data -- the same
        pattern used for the Copertura macchine per-machine ticket."""
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("1180x620")
        window.minsize(850, 420)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text=title, font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=8)
        toolbar = ttk.Frame(window)
        toolbar.grid(row=0, column=0, sticky="e", padx=10, pady=8)
        search = tk.StringVar()
        ttk.Label(toolbar, text="Search:").pack(side="left")
        ttk.Entry(toolbar, textvariable=search, width=28).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Clear", command=lambda: search.set("")).pack(side="left", padx=(5, 0))

        display_df = _format_display_dates(df, ["data", "consegna", "delivery_date", "tinto", "data_qualita", "data_uscita"])
        available_cols = [c for c in cols if display_df.empty or c in display_df.columns] or cols
        tree = ttk.Treeview(window, columns=available_cols, show="headings")
        sort_state: dict[str, bool] = {}

        def sort_col(col):
            nonlocal subset_df
            if subset_df.empty:
                return
            ascending = sort_state.get(col, True)
            subset_df = subset_df.sort_values(by=col, ascending=ascending, key=lambda s: s.astype(str))
            sort_state[col] = not ascending
            render()

        for col in available_cols:
            tree.heading(col, text=HEADERS_IT.get(col, col.capitalize()), command=lambda c=col: sort_col(c))
            tree.column(col, width=115 if col not in {"titolo", "colore", "comment", "new_comment", "raw_yarn_match", "issue"} else 170, anchor="center")
        tree.grid(row=1, column=0, sticky="nsew", padx=10)
        scroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)
        subset_df = display_df[available_cols] if not display_df.empty else display_df.reindex(columns=available_cols)

        def render(*_args):
            tree.delete(*tree.get_children())
            term = search.get().strip().casefold()
            for row in subset_df.itertuples(index=False, name=None):
                values = tuple("" if value is None else value for value in row)
                if not term or term in " ".join(str(value).casefold() for value in values):
                    tree.insert("", "end", values=values)
        search.trace_add("write", render)
        render()
        buttons = ttk.Frame(window, padding=10)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Button(
            buttons, text="Export Excel",
            command=lambda: export_dataframe_typed(
                subset_df.rename(columns=lambda c: HEADERS_IT.get(c, c.capitalize())),
                export_filename, window, sheet_title=title[:31] or "Export",
            ),
        ).pack(side="left")
        ttk.Button(buttons, text="Close", command=window.destroy).pack(side="right")

