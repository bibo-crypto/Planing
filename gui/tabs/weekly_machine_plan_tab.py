"""Weekly machine operating plan derived from shared production/DFM data."""
from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

from utility.excel_io import safe_save_workbook


class WeeklyMachinePlanTab(ttk.Frame):
    """Operational weekly view: one row per bagno/batch assigned to a machine."""

    def __init__(self, master, settimana_tab):
        super().__init__(master, padding=10)
        self.settimana_tab = settimana_tab
        self.week_var = tk.StringVar(value="All")
        self.status_var = tk.StringVar(value="Open Situazione Settimanale and calculate data first.")
        self._build()
        self.after(500, self.refresh)

    def _build(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Label(bar, text="Week:").pack(side="left")
        self.week_combo = ttk.Combobox(bar, textvariable=self.week_var, state="readonly", width=10)
        self.week_combo.pack(side="left", padx=5)
        self.week_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh())
        ttk.Button(bar, text="Refresh", command=self.refresh).pack(side="left", padx=5)
        ttk.Button(bar, text="Export Machine Plan", command=self.export).pack(side="left", padx=5)
        ttk.Label(bar, textvariable=self.status_var, foreground="#475467").pack(side="right")

        summary_frame = ttk.LabelFrame(self, text="Machine totals", padding=5)
        summary_frame.pack(fill="x", pady=(0, 8))
        self.summary_tree = ttk.Treeview(summary_frame, columns=("machine", "client", "weight", "batches"), show="headings", height=4)
        for col, title, width in (("machine", "Machine", 130), ("client", "Cliente", 150), ("weight", "Total Peso", 130), ("batches", "Batches", 100)):
            self.summary_tree.heading(col, text=title)
            self.summary_tree.column(col, width=width, anchor="center")
        self.summary_tree.pack(fill="x")

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frame, columns=("week", "machine", "cliente", "bagno", "peso", "status"), show="headings")
        for col, title, width in (("week", "Week", 75), ("machine", "Machine", 130), ("cliente", "Cliente", 140), ("bagno", "Bagno / Batch", 130), ("peso", "Peso", 110), ("status", "Status", 150)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _data(self) -> pd.DataFrame:
        df = getattr(self.settimana_tab, "batch_weights", pd.DataFrame())
        if not isinstance(df, pd.DataFrame):
            return pd.DataFrame()
        df = df.copy()
        if df.empty:
            return df
        selected = self.week_var.get()
        if selected and selected != "All":
            df = df[df["week_of_year"].astype(str) == str(selected)]
        return df.sort_values(["week_of_year", "machine_name", "cliente", "bagno"], na_position="last")

    def refresh(self):
        source = getattr(self.settimana_tab, "batch_weights", pd.DataFrame())
        weeks = [] if not isinstance(source, pd.DataFrame) or source.empty else sorted(str(x) for x in source["week_of_year"].dropna().unique())
        self.week_combo["values"] = ["All"] + weeks
        if self.week_var.get() not in self.week_combo["values"]:
            self.week_var.set("All")
        for tree in (self.tree, self.summary_tree):
            tree.delete(*tree.get_children())
        df = self._data()
        if df.empty:
            self.status_var.set("No calculated machine data")
            return
        for row in df.itertuples(index=False):
            self.tree.insert("", "end", values=(getattr(row, "week_of_year", ""), getattr(row, "machine_name", ""), getattr(row, "cliente", ""), getattr(row, "bagno", ""), round(float(getattr(row, "peso", 0) or 0), 2), "Planned"))
        summary = df.groupby(["machine_name", "cliente"], dropna=False).agg(total_peso=("peso", "sum"), batches=("bagno", "count")).reset_index()
        for row in summary.itertuples(index=False):
            self.summary_tree.insert("", "end", values=(row.machine_name, row.cliente, round(float(row.total_peso or 0), 2), int(row.batches)))
        self.status_var.set(f"{len(df)} planned batch(es)")

    def export(self):
        df = self._data()
        if df.empty:
            return messagebox.showinfo("Machine Plan", "There is no calculated weekly machine data to export.", parent=self)
        path = filedialog.asksaveasfilename(parent=self, title="Export Weekly Machine Plan", initialfile="Weekly Machine Plan.xlsx", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Machine Plan"
        headers = ["Week", "Machine", "Cliente", "Bagno / Batch", "Peso", "Status"]
        ws.append(headers)
        for row in df.itertuples(index=False):
            ws.append([getattr(row, "week_of_year", ""), getattr(row, "machine_name", ""), getattr(row, "cliente", ""), getattr(row, "bagno", ""), getattr(row, "peso", 0), "Planned"])
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        safe_save_workbook(wb, Path(path))
        wb.close()
        messagebox.showinfo("Machine Plan", f"Weekly machine plan exported to:\n{path}", parent=self)
