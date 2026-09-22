from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from utility import master_data


class MasterDataTab(ttk.Frame):
    """Small local CRUD screen for controlled reference data."""

    SCHEMAS = {
        "Customers": ("customers", ("code", "name")),
        "Machines": ("machines", ("code", "number", "rocche")),
        "Delave -> Lino": ("delave_map", ("delave_articolo", "raw_articolo")),
    }

    def __init__(self, parent, on_data_changed=None):
        super().__init__(parent)
        self._data = master_data.load()
        self._on_data_changed = on_data_changed or (lambda: None)
        self._trees = {}
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1); self.rowconfigure(0, weight=1)
        book = ttk.Notebook(self); book.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        for label, (key, columns) in self.SCHEMAS.items():
            frame = ttk.Frame(book, padding=8); frame.columnconfigure(0, weight=1); frame.rowconfigure(0, weight=1)
            tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
            for col in columns:
                tree.heading(col, text=col.replace("_", " ").title())
                tree.column(col, width=180, anchor="w")
            tree.grid(row=0, column=0, columnspan=3, sticky="nsew")
            ttk.Button(frame, text="Add", command=lambda k=key: self._edit(k)).grid(row=1, column=0, sticky="w", pady=(8, 0))
            ttk.Button(frame, text="Edit", command=lambda k=key: self._edit(k)).grid(row=1, column=1, padx=5, pady=(8, 0))
            ttk.Button(frame, text="Delete", command=lambda k=key: self._delete(k)).grid(row=1, column=2, sticky="e", pady=(8, 0))
            book.add(frame, text=label); self._trees[key] = tree
            self._refresh(key)

    def _refresh(self, key):
        tree = self._trees[key]; tree.delete(*tree.get_children())
        columns = self.SCHEMAS[next(label for label, (k, _) in self.SCHEMAS.items() if k == key)][1]
        for item in self._data[key]: tree.insert("", "end", values=[item.get(c, "") for c in columns])

    def _edit(self, key):
        label, columns = next((label, schema[1]) for label, schema in self.SCHEMAS.items() if schema[0] == key)
        tree = self._trees[key]; selected = tree.selection()
        old = self._data[key][tree.index(selected[0])] if selected else {}
        win = tk.Toplevel(self); win.title(f"{label} - Edit"); win.transient(self); win.grab_set()
        entries = {}
        for row, col in enumerate(columns):
            ttk.Label(win, text=col.replace("_", " ").title()).grid(row=row, column=0, padx=8, pady=5, sticky="w")
            entry = ttk.Entry(win, width=36); entry.insert(0, str(old.get(col, ""))); entry.grid(row=row, column=1, padx=8, pady=5); entries[col] = entry
        def save_item():
            item = {col: entries[col].get().strip() for col in columns}
            if not item[columns[0]]:
                messagebox.showwarning("Missing value", f"{columns[0]} is required.", parent=win); return
            if selected: self._data[key][tree.index(selected[0])] = item
            else: self._data[key].append(item)
            master_data.save(self._data); self._refresh(key); self._on_data_changed(); win.destroy()
        ttk.Button(win, text="Save", command=save_item).grid(row=len(columns), column=0, columnspan=2, pady=10)

    def _delete(self, key):
        tree = self._trees[key]; selected = tree.selection()
        if not selected: return
        if not messagebox.askyesno("Confirm delete", "Delete selected item?", parent=self): return
        del self._data[key][tree.index(selected[0])]; master_data.save(self._data); self._refresh(key); self._on_data_changed()
