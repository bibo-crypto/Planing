"""
prezzi_tab.py — "Prezzi" tab: browse Delta's Listini (price list) export.

Upload the Listini file once (its path is cached and auto-restored on the
next launch, same as Magazino Filato). Prices are looked up by Articolo
(CLARTICOLO) + colour code (CLCOLORE) together -- the same Articolo can
have several colours, each priced differently, since a different colour
often means a different raw yarn.
"""
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from calculate import prezzi as logic
from utility.prezzi_cache import load_prezzi_cache, save_prezzi_cache
from utility.excel_io import safe_save_workbook
from utility.path_manager import save_source
from utility.utils import keep_window_on_top, logger, bind_escape_to_close

COLUMNS = logic.DISPLAY_COLUMNS
HEADERS = logic.HEADERS
MAX_RENDER_ROWS = 1500


def _filter_color_code(df: pd.DataFrame, query: str) -> pd.DataFrame:
    code = query.strip()
    if not code:
        return df

    colors = df["CLCOLORE"].fillna("").astype(str)
    if code.isdigit():
        normalized_query = code.lstrip("0") or "0"
        normalized_colors = colors.map(
            lambda value: (value.lstrip("0") or "0") if value.isdigit() else value
        )
        exact_matches = normalized_colors.eq(normalized_query)
        if exact_matches.any():
            return df[exact_matches]

    return df[colors.str.casefold().str.contains(code.casefold(), regex=False)]


class PrezziTab(ttk.Frame):
    """Embeddable 'Prezzi' tab."""

    def __init__(self, master, on_shared_cache_changed=None, on_notification=None, on_notifications=None):
        super().__init__(master)
        self._on_shared_cache_changed = on_shared_cache_changed
        self._on_notification = on_notification
        self._on_notifications = on_notifications
        self._base_df = pd.DataFrame()
        self._validation_issues: list[dict[str, str]] = []
        self.prezzi_df = pd.DataFrame()
        self.summary_df = pd.DataFrame()
        self._uploading = False
        self._filter_after_id = None
        self._render_after_id = None
        self._render_generation = 0
        self._loaded_source_path = ""
        self._loaded_file_name = ""
        self._sort_column = ""
        self._sort_reverse = False
        self._articolo_var = tk.StringVar()
        self._codice_var = tk.StringVar()
        self._category_var = tk.StringVar()
        self._anomalies_window = None

        self._configure_styles()
        self._build_upload_panel()
        self._build_treeview()
        self.after_idle(self._restore_from_cache)

    # ------------------------------------------------------------------
    def _configure_styles(self) -> None:
        # Matches the navy heading style already used elsewhere in the app
        # (ui/gui.py sets the "clam" theme globally at startup, needed for
        # a Treeview heading's background colour to actually show).
        style = ttk.Style(self)
        style.configure("Prezzi.Treeview", background="#ffffff", fieldbackground="#ffffff",
                         foreground="#1d2939", rowheight=27, font=("Segoe UI", 9))
        style.configure("Prezzi.Treeview.Heading", background="#16324f", foreground="#ffffff",
                         relief="raised", borderwidth=1, padding=(8, 6),
                         font=("Segoe UI", 9, "bold"))
        style.map("Prezzi.Treeview.Heading",
                  foreground=[("pressed", "#ffffff"), ("active", "#ffffff"), ("!active", "#ffffff")],
                  background=[("pressed", "#0b2239"), ("active", "#244b70"), ("!active", "#16324f")])
        style.map("Prezzi.Treeview", background=[("selected", "#2563eb")],
                  foreground=[("selected", "#ffffff")])

    def _build_upload_panel(self) -> None:
        panel = ttk.LabelFrame(self, text="1) Upload Listini")
        panel.pack(side="top", fill="x", padx=8, pady=6)

        row = ttk.Frame(panel)
        row.pack(fill="x", padx=4, pady=4)
        self.status_var = tk.StringVar(value="No Listini file uploaded")
        self._btn_upload = ttk.Button(row, text="📤 Upload Listini", command=self._on_upload)
        self._btn_upload.pack(side="left")
        ttk.Label(row, textvariable=self.status_var, foreground="#666666").pack(side="left", padx=8)

        row_search = ttk.Frame(panel)
        row_search.pack(fill="x", padx=4, pady=4)
        ttk.Label(row_search, text="Articolo:").pack(side="left")
        ttk.Entry(row_search, textvariable=self._articolo_var, width=20).pack(side="left", padx=(4, 16))
        ttk.Label(row_search, text="Codice:").pack(side="left")
        ttk.Entry(row_search, textvariable=self._codice_var, width=14).pack(side="left", padx=(4, 16))
        ttk.Label(row_search, text="Category:").pack(side="left")
        ttk.Entry(row_search, textvariable=self._category_var, width=18).pack(side="left", padx=(4, 16))
        self._articolo_var.trace_add("write", lambda *_: self._on_search_changed())
        self._codice_var.trace_add("write", lambda *_: self._on_search_changed())
        self._category_var.trace_add("write", lambda *_: self._on_search_changed())

        ttk.Button(row_search, text="Clear", command=self._clear_filters).pack(side="left", padx=4)

        self._btn_export = ttk.Button(row_search, text="📤 Export to Excel", command=self._on_export)
        self._btn_export.pack(side="right")

        self._btn_price_changes = ttk.Button(
            row_search, text="⚠ Price Changes", command=self._open_price_anomalies,
        )
        self._btn_price_changes.pack(side="right", padx=(0, 8))

    def _build_treeview(self) -> None:
        frame = ttk.Frame(self)
        frame.pack(side="top", fill="both", expand=True, padx=8, pady=6)

        self.tree = ttk.Treeview(frame, columns=COLUMNS, show="headings", selectmode="browse",
                                  style="Prezzi.Treeview")
        for c in COLUMNS:
            self.tree.heading(c, text=HEADERS[c], command=lambda c=c: self._sort_by(c))
            self.tree.column(c, width=150 if c in ("DESCRIZARTICOLOLI", "CLDESCR") else 120,
                              anchor="w" if c in ("CLARTICOLO", "DESCRIZARTICOLOLI", "CLDESCR") else "center")
        self.tree.tag_configure("oddrow", background="#ffffff")
        self.tree.tag_configure("evenrow", background="#eaf1fb")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Upload + startup restore
    # ------------------------------------------------------------------
    def _on_upload(self) -> None:
        if self._uploading:
            return
        path = filedialog.askopenfilename(
            title="Select the Listini file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self._load_path(path, save_cache=True)

    def _restore_from_cache(self) -> None:
        cache = load_prezzi_cache()
        source_path = cache.get("source_path")
        if source_path and Path(source_path).is_file():
            self._load_path(source_path, save_cache=False)

    def _load_path(self, path: str, save_cache: bool) -> None:
        normalized_path = str(Path(path).resolve())
        if self._base_df is not None and not self._base_df.empty and normalized_path == self._loaded_source_path:
            self.status_var.set(f"{Path(path).name} — {len(self._base_df)} price rows")
            return
        self._uploading = True
        self._btn_upload.config(state="disabled")
        self.status_var.set(f"Loading {Path(path).name}…")

        def worker():
            try:
                df, errors = logic.load_prezzi(path)
            except Exception as exc:  # noqa: BLE001
                errors = [f"An error occurred while reading the file: {exc}"]
                df = None
            issues = []
            validation_error = None
            if df is not None:
                try:
                    issues = logic.validate_price_data(df)
                except Exception as exc:  # noqa: BLE001
                    validation_error = exc

            def apply_result():
                self._uploading = False
                self._btn_upload.config(state="normal")
                if df is None:
                    self.status_var.set("Failed to load Listini file")
                    if save_cache:
                        messagebox.showerror("Error", "\n".join(errors))
                    else:
                        logger.warning("Prezzi: could not restore cached file: %s", "; ".join(errors))
                    return
                # Articoli may be uploaded after Listini; enrich again here
                # so the Category column and category rules are immediately
                # refreshed without requiring a second Listini upload.
                self._base_df = df
                self._validation_issues = list(issues)
                # Public, unfiltered-by-search accessor for other tabs
                # (Ordine Elvy's Livello/Prezzo lookup) -- self.summary_df
                # changes with the search boxes, this doesn't.
                self.prezzi_df = self._base_df
                self._loaded_source_path = normalized_path
                self._loaded_file_name = Path(path).name
                self._update_status(len(self._base_df))
                if validation_error is not None:
                    logger.warning("Prezzi: price validation failed: %s", validation_error)
                if self._on_notifications:
                    self._on_notifications([
                        {**issue, "page": "Prezzi"}
                        for issue in issues
                    ])
                elif self._on_notification:
                    for issue in issues:
                        self._on_notification(issue["key"], issue["title"], issue["message"], "Prezzi", issue["severity"])
                if save_cache:
                    save_prezzi_cache(path)
                    save_source("listini", path)
                    if self._on_shared_cache_changed:
                        self._on_shared_cache_changed()
                self._apply_search_and_sort()

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Search + sort + render
    # ------------------------------------------------------------------
    def _on_search_changed(self) -> None:
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except tk.TclError:
                pass
        self._filter_after_id = self.after(180, self._apply_search_and_sort)

    def _clear_filters(self) -> None:
        self._articolo_var.set("")
        self._codice_var.set("")
        self._category_var.set("")
        self._sort_column = ""
        self._sort_reverse = False
        self._apply_search_and_sort()

    def _apply_search_and_sort(self) -> None:
        self._filter_after_id = None
        df = self._base_df
        if df.empty:
            self.summary_df = pd.DataFrame(columns=COLUMNS)
            self._update_status(0)
            self._render()
            return

        articolo_q = self._articolo_var.get().strip().lower()
        if articolo_q:
            df = df[df["CLARTICOLO"].str.lower().str.contains(articolo_q, regex=False)]
        codice_q = self._codice_var.get().strip().lower()
        if codice_q:
            df = _filter_color_code(df, codice_q)
        category_q = self._category_var.get().strip().lower()
        if category_q:
            df = df[df["CATEGORY"].str.lower().str.contains(category_q, regex=False)]

        if self._sort_column:
            ascending = not self._sort_reverse
            df = df.sort_values(self._sort_column, ascending=ascending, kind="mergesort")

        self.summary_df = df.reset_index(drop=True)
        self._update_status(len(df))
        self._render()

    def _update_status(self, matched_rows: int) -> None:
        visible_rows = min(matched_rows, MAX_RENDER_ROWS)
        suffix = f" — showing {visible_rows:,} of {matched_rows:,}" if visible_rows < matched_rows else f" — {matched_rows:,} price rows"
        self.status_var.set(f"{self._loaded_file_name}{suffix}")

    def _sort_by(self, column: str) -> None:
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
        self._apply_search_and_sort()

    def _render(self) -> None:
        self._render_generation += 1
        generation = self._render_generation
        if self._render_after_id is not None:
            try:
                self.after_cancel(self._render_after_id)
            except tk.TclError:
                pass
            self._render_after_id = None
        self.tree.delete(*self.tree.get_children())
        if self.summary_df.empty:
            return
        rows = list(self.summary_df[COLUMNS].head(MAX_RENDER_ROWS).itertuples(index=False, name=None))

        def insert_chunk(start=0):
            if generation != self._render_generation:
                return
            end = min(start + 150, len(rows))
            for idx in range(start, end):
                row = rows[idx]
                # COLUMNS = Articolo, Descrizione, Codice, Colore,
                # Category, Livello, Prezzo.  Category is text and must
                # never be formatted with the numeric :g formatter.
                livello = "" if pd.isna(row[5]) else f"{float(row[5]):g}"
                prezzo = "" if pd.isna(row[6]) else f"{float(row[6]):.2f}"
                self.tree.insert("", "end", values=(
                    row[0], row[1], row[2], row[3], row[4], livello, prezzo,
                ), tags=("evenrow" if idx % 2 == 0 else "oddrow",))
            if end < len(rows):
                self._render_after_id = self.after(1, insert_chunk, end)
            else:
                self._render_after_id = None

        insert_chunk()

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def _on_export(self) -> None:
        if self.summary_df.empty:
            messagebox.showinfo("No data", "Upload the Listini file first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Prezzi.xlsx"
        )
        if not path:
            return

        # Snapshot the visible result and move all workbook work off the Tk
        # thread.  openpyxl can be slow on large Listini exports.
        export_df = self.summary_df[COLUMNS].copy()
        self._btn_export.config(state="disabled")

        def worker():
            error = None
            try:
                self._write_export_file(path, export_df)
            except Exception as exc:  # noqa: BLE001
                error = exc

            def finish():
                self._btn_export.config(state="normal")
                if error is not None:
                    messagebox.showerror("Export failed", str(error), parent=self)
                    return
                logger.info("Prezzi: exported to %s", path)
                messagebox.showinfo("Completed", f"Export completed successfully:\n{path}", parent=self)

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _write_export_file(path: str, export_df: pd.DataFrame) -> None:
        """Write a Prezzi workbook without blocking the Tk event loop."""

        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        thin = Side(style="thin", color="B0B0B0")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center")
        header_fill = PatternFill(start_color="16324F", end_color="16324F", fill_type="solid")

        wb = Workbook()
        ws = wb.active
        ws.title = "Prezzi"
        ws.append([HEADERS[c] for c in COLUMNS])
        for cell in ws[1]:
            cell.font = Font(bold=True, name="Arial", color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        for row in export_df.itertuples(index=False, name=None):
            values = list(row)
            for index in (5, 6):
                if pd.isna(values[index]):
                    values[index] = None
            ws.append(values)

        last_row = ws.max_row
        last_col_letter = get_column_letter(len(COLUMNS))
        if last_row > 1:
            ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"
        ws.freeze_panes = "A2"
        widths = {"CLARTICOLO": 16, "DESCRIZARTICOLOLI": 24, "CLCOLORE": 14,
                  "CLDESCR": 20, "CATEGORY": 20, "LIVELLOLPZ": 12, "PREZZOLPZ": 12}
        for i, c in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 16)

        safe_save_workbook(wb, path)

    # ------------------------------------------------------------------
    # Price-change anomalies
    # ------------------------------------------------------------------
    def _open_price_anomalies(self) -> None:
        """Calculate the report off the Tk thread, then render it."""
        if self._base_df is None or self._base_df.empty:
            messagebox.showinfo("Price Changes", "Upload the Listini file first.", parent=self)
            return
        if self._anomalies_window is not None:
            try:
                if self._anomalies_window.winfo_exists():
                    self._anomalies_window.lift()
                    self._anomalies_window.focus_force()
                    return
            except tk.TclError:
                self._anomalies_window = None
        snapshot = self._base_df.copy(deep=True)
        self._btn_price_changes.config(state="disabled", text="Calculating…")

        def worker():
            try:
                result = logic.detect_price_anomalies(snapshot, min_pct_change=10.0)
                error = None
            except Exception as exc:  # noqa: BLE001
                result, error = pd.DataFrame(), exc

            def finish():
                self._btn_price_changes.config(state="normal", text="⚠ Price Changes")
                if error is not None:
                    messagebox.showerror("Price Changes", str(error), parent=self)
                    return
                self._show_price_anomalies(result)
            self.after(0, finish)

        threading.Thread(target=worker, name="prezzi-price-changes", daemon=True).start()

    def _show_price_anomalies(self, anomalies: pd.DataFrame) -> None:
        """Flag Articolo+Colore pairs whose price jumped/dropped by 10%+
        between two dated Listini entries -- catches pricing mistakes and
        genuine repricings alike, sorted by the biggest change first."""
        if self._anomalies_window is not None:
            try:
                if self._anomalies_window.winfo_exists():
                    self._anomalies_window.lift()
                    self._anomalies_window.focus_force()
                    return
            except tk.TclError:
                pass
            self._anomalies_window = None

        window = tk.Toplevel(self)
        bind_escape_to_close(window)
        keep_window_on_top(window)
        self._anomalies_window = window
        window.title("Listini — Price Changes (10%+)")
        window.geometry("760x460")
        window.minsize(560, 320)
        ttk.Label(
            window, text="Articolo + Colore pairs whose price changed by 10% or more between two dated entries",
            font=("Segoe UI", 10, "bold"), wraplength=720,
        ).pack(anchor="w", padx=10, pady=(10, 6))

        frame = ttk.Frame(window)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        columns = ("CLARTICOLO", "CLCOLORE", "CLDESCR", "CATEGORY", "old_price", "new_price", "pct_change", "changed_on", "issue")
        labels = {
            "CLARTICOLO": "Articolo", "CLCOLORE": "Colore", "CLDESCR": "Descr. Colore",
            "CATEGORY": "Category", "old_price": "Old Price", "new_price": "New Price", "pct_change": "Change %",
            "changed_on": "Changed On", "issue": "Issue",
        }
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        widths = [110, 90, 150, 130, 90, 90, 90, 100, 240]
        for column, width in zip(columns, widths):
            tree.heading(column, text=labels[column])
            tree.column(column, width=width, anchor="center")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree.tag_configure("up", foreground="#b91c1c")
        tree.tag_configure("down", foreground="#15803d")

        if anomalies.empty:
            ttk.Label(frame, text="No price changes of 10% or more found.").grid(row=0, column=0)
        else:
            for _, row in anomalies.iterrows():
                try:
                    tag = "up" if float(row["pct_change"]) > 0 else "down"
                except (TypeError, ValueError):
                    tag = "up"
                tree.insert("", "end", values=(
                    row["CLARTICOLO"], row["CLCOLORE"], row["CLDESCR"], row.get("CATEGORY", ""),
                    f"{float(row['old_price']):.2f}" if str(row["old_price"]) else "",
                    f"{float(row['new_price']):.2f}" if str(row["new_price"]) else "",
                    f"{float(row['pct_change']):+.1f}%" if str(row["pct_change"]) else "", row["changed_on"], row.get("issue", ""),
                ), tags=(tag,))

        def export_anomalies():
            if anomalies.empty:
                messagebox.showinfo("Price Changes", "Nothing to export.", parent=window)
                return
            path = filedialog.asksaveasfilename(
                parent=window, title="Export Price Changes", defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")], initialfile="price_changes.xlsx",
            )
            if not path:
                return
            export_df = anomalies.rename(columns=labels).copy()
            export_button.config(state="disabled", text="Exporting…")

            def worker():
                try:
                    export_df.to_excel(path, index=False, sheet_name="Price Changes")
                    error = None
                except Exception as exc:  # noqa: BLE001
                    error = exc

                def finish():
                    export_button.config(state="normal", text="📤 Export to Excel")
                    if error is not None:
                        messagebox.showerror("Price Changes", str(error), parent=window)
                    else:
                        messagebox.showinfo("Price Changes", f"Export completed:\n{path}", parent=window)
                window.after(0, finish)

            threading.Thread(target=worker, name="prezzi-price-changes-export", daemon=True).start()

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=(0, 10))
        export_button = ttk.Button(buttons, text="📤 Export to Excel", command=export_anomalies)
        export_button.pack(side="right", padx=3)
