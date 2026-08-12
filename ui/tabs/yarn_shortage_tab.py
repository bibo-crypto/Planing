"""Yarn shortage summary derived from the current Situazione table."""
from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


class YarnShortageTab(ttk.Frame):
    """Group PG-X rows by client and Titolo, summing Rocche."""

    COLUMNS = ("cliente", "titolo", "rocche")
    HEADERS = {"cliente": "CLIENTE", "titolo": "Titolo", "rocche": "Rocche mancanti"}

    def __init__(self, master, situazione_tab):
        super().__init__(master)
        self.situazione_tab = situazione_tab
        self.source_df = pd.DataFrame()
        self.current_df = pd.DataFrame(columns=self.COLUMNS)
        self.sort_state = {}
        self._filter_after_id = None

        self._configure_style()
        self._build_toolbar()
        self._build_tree()
        self.refresh()

    def _configure_style(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Shortage.Treeview", background="#ffffff", fieldbackground="#ffffff",
                        foreground="#1d2939", rowheight=29, borderwidth=1, relief="solid",
                        bordercolor="#b8c6d6", lightcolor="#b8c6d6", darkcolor="#b8c6d6",
                        font=("Segoe UI", 9, "bold"))
        style.configure("Shortage.Treeview.Heading", background="#16324f", foreground="#ffffff",
                        relief="raised", borderwidth=1, padding=(8, 8),
                        font=("Segoe UI", 9, "bold"))
        style.map("Shortage.Treeview", background=[("selected", "#2563eb")],
                  foreground=[("selected", "#ffffff")])

    def _build_toolbar(self):
        bar = ttk.Frame(self)
        bar.pack(side="top", fill="x", padx=8, pady=(8, 4))
        ttk.Button(bar, text="Refresh from Situazione", command=self.refresh).pack(side="left", padx=4)
        ttk.Button(bar, text="Export to Excel", command=self.export).pack(side="left", padx=4)

        ttk.Label(bar, text="Cliente:").pack(side="left", padx=(18, 4))
        self.client_var = tk.StringVar(value="Tutti")
        self.client_combo = ttk.Combobox(bar, textvariable=self.client_var, state="readonly", width=24)
        self.client_combo.pack(side="left")
        self.client_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_filter())

        ttk.Label(bar, text="Search:").pack(side="left", padx=(18, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)
        ttk.Entry(bar, textvariable=self.search_var, width=28).pack(side="left")
        self.summary_lbl = ttk.Label(bar, text="")
        self.summary_lbl.pack(side="right", padx=8)

        ttk.Label(self, text="Rows source: Situazione where Comment contains PG-X",
                  foreground="#667085", anchor="w").pack(side="top", fill="x", padx=12, pady=(0, 4))

    def _build_tree(self):
        frame = ttk.Frame(self, borderwidth=1, relief="solid")
        frame.pack(side="top", fill="both", expand=True, padx=8, pady=(2, 8))
        self.tree = ttk.Treeview(frame, columns=self.COLUMNS, show="tree headings",
                                 style="Shortage.Treeview")
        self.tree.heading("#0", text="Group", anchor="center")
        self.tree.column("#0", width=55, minwidth=55, stretch=False, anchor="center")
        for col in self.COLUMNS:
            self.tree.heading(col, text=self.HEADERS[col], anchor="center",
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=160 if col != "rocche" else 130,
                             minwidth=80, anchor="center")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.tree.tag_configure("stripe", background="#f3f6fa", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("client", background="#dbeafe", font=("Segoe UI", 9, "bold"))

    @staticmethod
    def _number(value):
        # Preserve numeric values exactly; stripping dots is only for text
        # values written with European thousands separators.
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return 0.0 if pd.isna(value) else float(value)
        try:
            if pd.isna(value):
                return 0.0
        except (TypeError, ValueError):
            pass
        text = str(value).strip().replace(".", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return 0.0

    def refresh(self):
        source = getattr(self.situazione_tab, "current_df", pd.DataFrame())
        if source is None or source.empty:
            self.source_df = pd.DataFrame()
            self.current_df = pd.DataFrame(columns=self.COLUMNS)
        else:
            source = source.copy()
            comment = source.get("comment", pd.Series("", index=source.index)).fillna("").astype(str)
            # PG-X is intentionally matched case-insensitively and as a substring,
            # allowing values such as "UG PG-X" or "PG-X / ...".
            source = source[comment.str.contains("PG-X", case=False, regex=False)].copy()
            source["cliente"] = source.get("cliente", "").fillna("").astype(str).str.strip()
            source["titolo"] = source.get("titolo", "").fillna("").astype(str).str.strip()
            source["rocche"] = source.get("rocche", 0).map(self._number)
            self.source_df = source
            grouped = (source.groupby(["cliente", "titolo"], dropna=False, as_index=False)["rocche"]
                        .sum().rename(columns={"rocche": "rocche"}))
            self.current_df = grouped.reindex(columns=self.COLUMNS)

        clients = sorted({str(x) for x in self.current_df.get("cliente", []) if str(x).strip()})
        previous = self.client_var.get()
        values = ["Tutti"] + clients
        self.client_combo["values"] = values
        self.client_var.set(previous if previous in values else "Tutti")
        self._apply_filter()

    def _on_search_changed(self, *_args):
        if self._filter_after_id is not None:
            self.after_cancel(self._filter_after_id)
        self._filter_after_id = self.after(180, self._apply_filter)

    def _apply_filter(self):
        self._filter_after_id = None
        df = self.current_df.copy()
        client = self.client_var.get()
        if client and client != "Tutti":
            df = df[df["cliente"].astype(str) == client]
        query = self.search_var.get().strip().lower()
        if query and not df.empty:
            searchable = df.astype(str)
            df = df[searchable.apply(lambda col: col.str.contains(query, case=False, regex=False)).any(axis=1)]
        self._render(df)

    def _sort_by(self, col):
        if self.current_df.empty:
            return
        ascending = self.sort_state.get(col, True)
        self.current_df = self.current_df.sort_values(col, ascending=ascending,
                                                       key=lambda s: s.astype(str))
        self.sort_state[col] = not ascending
        self._apply_filter()

    def _render(self, df):
        self.tree.delete(*self.tree.get_children())
        if df.empty:
            self.summary_lbl.config(text="0 Titoli | 0 Rocche")
            return
        total = float(df["rocche"].sum())
        self.summary_lbl.config(text=f"{len(df)} Titoli | {total:g} Rocche")
        for row_index, (client, client_df) in enumerate(df.groupby("cliente", sort=False)):
            parent = self.tree.insert("", "end", text="", values=(client, "", ""), tags=("client",), open=True)
            for _, row in client_df.iterrows():
                self.tree.insert(parent, "end", text="", values=("", row["titolo"], row["rocche"]),
                                 tags=(("stripe",) if row_index % 2 else ()))

    def export(self):
        if self.current_df.empty:
            messagebox.showinfo("No data", "Refresh the data after loading Situazione.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="Mancanza_Filato.xlsx")
        if not path:
            return
        client = self.client_var.get()
        export_df = self.current_df.copy()
        if client and client != "Tutti":
            export_df = export_df[export_df["cliente"].astype(str) == client]
        query = self.search_var.get().strip().lower()
        if query and not export_df.empty:
            searchable = export_df.astype(str)
            export_df = export_df[searchable.apply(lambda col: col.str.contains(query, case=False, regex=False)).any(axis=1)]
        export_df = export_df.rename(columns=self.HEADERS)
        export_df.to_excel(path, index=False, sheet_name="Yarn Shortage")
        messagebox.showinfo("Completed", f"Export completed successfully:\n{path}")
