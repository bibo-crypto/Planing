"""
magazino_filato_tab.py
Raw yarn warehouse summary tab ("Magazino Filato"): upload the Magazino
export, apply the filter rules in magazino_logic.py, and show one row per
raw-yarn Partita with its total cones (Mag.rocche) and weight (Mag.peso).
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

import magazino_logic as logic
from utils import logger
from magazino_cache import load_magazino_cache, save_magazino_cache

COLUMNS = ["articolo", "partita", "mag_rocche", "mag_peso"]
HEADERS = {"articolo": "Articolo", "partita": "Partita", "mag_rocche": "Mag.rocche", "mag_peso": "Mag.peso"}


class MagazinoFilatoTab(ttk.Frame):
    """Embeddable 'Magazino Filato' tab."""

    def __init__(self, master):
        super().__init__(master)
        self.summary_df = pd.DataFrame()
        self._shared_path = ""
        self._syncing = False

        self._build_upload_panel()
        self._build_treeview()
        self._restore_cached_data()

    def _build_upload_panel(self):
        panel = ttk.LabelFrame(self, text="1) Upload the Magazino export")
        panel.pack(side="top", fill="x", padx=8, pady=6)

        row = ttk.Frame(panel)
        row.pack(fill="x", padx=4, pady=4)

        self.status_var = tk.StringVar(value="No file uploaded")
        ttk.Button(row, text="Select Magazino File", command=self._on_upload).pack(side="left")
        ttk.Label(row, textvariable=self.status_var, foreground="#666666").pack(side="left", padx=8)

        ttk.Button(row, text="Export Excel", command=self._on_export).pack(side="right")

    def _build_treeview(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", padx=8, pady=(2, 0))
        self.search_var = tk.StringVar()
        self._search_after_id = None
        self.search_var.trace_add("write", self._on_search_changed)
        ttk.Label(toolbar, text="Search:").pack(side="left", padx=(0, 5))
        ttk.Entry(toolbar, textvariable=self.search_var, width=30).pack(side="left")
        ttk.Button(toolbar, text="Clear", command=lambda: self.search_var.set("")).pack(side="left", padx=6)

        frame = ttk.Frame(self)
        frame.pack(side="top", fill="both", expand=True, padx=8, pady=6)

        style = ttk.Style(self)
        style.configure("Magazino.Treeview", font=("Segoe UI", 9, "bold"), rowheight=25)
        style.configure("Magazino.Treeview.Heading", font=("Segoe UI", 9, "bold"),
                        background="#16324F", foreground="#FFFFFF")

        self.tree = ttk.Treeview(frame, columns=COLUMNS, show="headings", selectmode="browse",
                                 style="Magazino.Treeview")
        for c in COLUMNS:
            self.tree.heading(c, text=HEADERS[c], command=lambda col=c: self._on_sort(col))
            self.tree.column(c, width=170, anchor="center")

        self.tree.tag_configure("evenrow", background="#EAF1FB", foreground="#111827")
        self.tree.tag_configure("oddrow", background="#FFFFFF", foreground="#111827")
        self._sort_column = "articolo"
        self._sort_reverse = False

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _on_upload(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df, errors = logic.load_magazino(path)
        except Exception as exc:  # noqa: BLE001
            errors = [f"An error occurred while reading the file: {exc}"]
            df = None

        if errors or df is None or df.empty:
            msg = "; ".join(errors) if errors else "The file is empty after filtering"
            self.status_var.set(f"❌ {msg}")
            messagebox.showerror("Error", msg)
            return

        self.summary_df = logic.summarize_by_partita(df)
        self._shared_path = str(path)
        save_magazino_cache(path, self.summary_df)
        self.status_var.set(f"✅ {len(self.summary_df)} batches - {os.path.basename(path)}")
        logger.info("Magazino Filato: loaded %d batches from %s", len(self.summary_df), os.path.basename(path))
        self._render()

    def sync_shared_async(self):
        """Load a Magazino selected on the Ordini ELVY page without blocking Tk."""
        cache = load_magazino_cache()
        source = str(cache.get("source_path", ""))
        if not source or not os.path.isfile(source) or source == self._shared_path or self._syncing:
            return
        self._syncing = True

        def worker():
            try:
                df, errors = logic.load_magazino(source)
                summary = logic.summarize_by_partita(df) if not errors and df is not None else None
            except Exception as exc:  # noqa: BLE001
                errors, summary = [str(exc)], None

            def apply_result():
                self._syncing = False
                if errors or summary is None or summary.empty:
                    return
                self.summary_df = summary
                self._shared_path = source
                save_magazino_cache(source, summary)
                self.status_var.set(f"✅ {len(summary)} batches - {os.path.basename(source)}")
                self._render()

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    def _render(self):
        self.tree.delete(*self.tree.get_children())
        query = self.search_var.get().strip().lower()
        visible = self.summary_df
        if query:
            values = visible.reindex(columns=COLUMNS, fill_value="").astype(str)
            visible = visible[values.apply(
                lambda col: col.str.contains(query, case=False, regex=False)
            ).any(axis=1)]
        visible = visible.sort_values(
            by=self._sort_column,
            key=lambda col: col.astype(str).str.lower(),
            ascending=not self._sort_reverse,
        )
        for i, (_, r) in enumerate(visible.iterrows()):
            self.tree.insert("", "end", values=(
                r["articolo"], r["partita"], f"{r['mag_rocche']:g}", f"{r['mag_peso']:g}",
            ), tags=("evenrow" if i % 2 == 0 else "oddrow",))

    def _apply_filter(self):
        self._search_after_id = None
        self._render()

    def _on_search_changed(self, *_args):
        """Filter automatically while typing, with a small debounce."""
        if self._search_after_id is not None:
            self.after_cancel(self._search_after_id)
        self._search_after_id = self.after(120, self._apply_filter)

    def _on_sort(self, column):
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
        self._render()

    def _restore_cached_data(self):
        cache = load_magazino_cache()
        source = str(cache.get("source_path", ""))
        rows = cache.get("summary_rows")
        if not source or not os.path.isfile(source) or not isinstance(rows, list) or not rows:
            return
        try:
            self.summary_df = pd.DataFrame(rows).reindex(columns=COLUMNS)
            self._shared_path = source
            self.status_var.set(f"✅ {len(self.summary_df)} batches - {os.path.basename(source)}")
            self._render()
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not restore Magazino table: %s", exc)

    def _on_export(self):
        if self.summary_df.empty:
            messagebox.showinfo("No data", "Upload the Magazino file first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             initialfile="Magazino_Filato.xlsx")
        if not path:
            return

        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        thin = Side(style="thin", color="B0B0B0")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center")
        header_fill = PatternFill(start_color="16324F", end_color="16324F", fill_type="solid")

        wb = Workbook()
        ws = wb.active
        ws.title = "Magazino Filato"
        ws.append([HEADERS[c] for c in COLUMNS])
        for cell in ws[1]:
            cell.font = Font(bold=True, name="Arial", color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        for _, r in self.summary_df.iterrows():
            ws.append([r["articolo"], r["partita"], r["mag_rocche"], r["mag_peso"]])
            for cell in ws[ws.max_row]:
                cell.font = Font(name="Arial")
                cell.alignment = center
                cell.border = border

        last_row = ws.max_row
        last_col_letter = get_column_letter(len(COLUMNS))
        if last_row > 1:
            ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"
        ws.freeze_panes = "A2"
        for i in range(1, len(COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 16

        wb.save(path)
        logger.info("Magazino Filato: exported to %s", path)
        messagebox.showinfo("Completed", f"Export completed successfully:\n{path}")
