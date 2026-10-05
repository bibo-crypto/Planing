from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from calculate import prezzi as prezzi_logic
from utility import master_data
from utility.utils import bind_escape_to_close, logger


class MasterDataTab(ttk.Frame):
    """Small local CRUD screen for controlled reference data."""

    SCHEMAS = {
        "Customers": ("customers", ("code", "name")),
        "Machines": ("machines", ("code", "number", "rocche")),
        "Delave -> Lino": ("delave_map", ("delave_articolo", "raw_articolo", "titolo")),
    }

    REVIEW_COLUMNS = (
        ("articolo", "Articolo", 110), ("descrizione", "Descrizione", 220),
        ("customer", "Client", 80), ("colori", "Colori", 60),
        ("suggestion", "Suggested category", 200), ("evidence", "Matches", 70),
    )

    def __init__(self, parent, on_data_changed=None, prezzi_tab=None):
        super().__init__(parent)
        self._data = master_data.load()
        self._on_data_changed = on_data_changed or (lambda: None)
        self._prezzi_tab = prezzi_tab
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
            scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            scroll.grid(row=0, column=3, sticky="ns")
            tree.configure(yscrollcommand=scroll.set)
            tree.grid(row=0, column=0, columnspan=3, sticky="nsew")
            ttk.Button(frame, text="Add", command=lambda k=key: self._edit(k)).grid(row=1, column=0, sticky="w", pady=(8, 0))
            ttk.Button(frame, text="Edit", command=lambda k=key: self._edit(k)).grid(row=1, column=1, padx=5, pady=(8, 0))
            ttk.Button(frame, text="Delete", command=lambda k=key: self._delete(k)).grid(row=1, column=2, sticky="e", pady=(8, 0))
            book.add(frame, text=label); self._trees[key] = tree
            self._refresh(key)
        self._build_category_review(book)
        self.bind("<Map>", lambda _event: self._refresh_category_review())

    def _refresh(self, key):
        tree = self._trees[key]; tree.delete(*tree.get_children())
        columns = self.SCHEMAS[next(label for label, (k, _) in self.SCHEMAS.items() if k == key)][1]
        for item in self._data[key]: tree.insert("", "end", values=[item.get(c, "") for c in columns])

    def _edit(self, key):
        label, columns = next((label, schema[1]) for label, schema in self.SCHEMAS.items() if schema[0] == key)
        tree = self._trees[key]; selected = tree.selection()
        old = self._data[key][tree.index(selected[0])] if selected else {}
        win = tk.Toplevel(self); win.title(f"{label} - Edit"); win.transient(self); win.grab_set()
        bind_escape_to_close(win)
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

    def _build_category_review(self, book):
        frame = ttk.Frame(book, padding=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=3)
        frame.rowconfigure(5, weight=1)
        self._review_status = tk.StringVar()
        ttk.Label(
            frame, textvariable=self._review_status, wraplength=760, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        columns = [name for name, _label, _width in self.REVIEW_COLUMNS]
        self._review_tree = ttk.Treeview(
            frame, columns=columns, show="headings", selectmode="browse", height=6,
        )
        for name, label, width in self.REVIEW_COLUMNS:
            self._review_tree.heading(name, text=label)
            self._review_tree.column(name, width=width, anchor="w")
        self._review_tree.grid(row=1, column=0, columnspan=3, sticky="nsew")
        review_scroll = ttk.Scrollbar(frame, orient="vertical", command=self._review_tree.yview)
        review_scroll.grid(row=1, column=3, sticky="ns")
        self._review_tree.configure(yscrollcommand=review_scroll.set)
        self._review_tree.bind("<<TreeviewSelect>>", self._on_review_select)

        controls = ttk.Frame(frame)
        controls.grid(row=2, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(controls, text="Category:").pack(side="left")
        self._review_category = ttk.Combobox(controls, width=40)
        self._review_category.pack(side="left", padx=6)
        ttk.Button(
            controls, text="Assign category", command=self._assign_category,
        ).pack(side="left", padx=(0, 5))
        ttk.Button(
            controls, text="Refresh", command=self._refresh_category_review,
        ).pack(side="left")

        ttk.Label(frame, text="Assigned manually", font=("Segoe UI", 9, "bold")).grid(
            row=4, column=0, sticky="w",
        )
        self._override_tree = ttk.Treeview(
            frame, columns=("articolo", "category"), show="headings",
            selectmode="browse", height=3,
        )
        self._override_tree.heading("articolo", text="Articolo")
        self._override_tree.column("articolo", width=140, anchor="w")
        self._override_tree.heading("category", text="Category")
        self._override_tree.column("category", width=260, anchor="w")
        self._override_tree.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
        override_scroll = ttk.Scrollbar(frame, orient="vertical", command=self._override_tree.yview)
        override_scroll.grid(row=5, column=3, sticky="ns", pady=(4, 0))
        self._override_tree.configure(yscrollcommand=override_scroll.set)
        ttk.Button(
            frame, text="Remove assignment", command=self._remove_override,
        ).grid(row=6, column=0, sticky="w", pady=(8, 0))
        book.add(frame, text="Category Review")

    def _refresh_category_review(self):
        if not hasattr(self, "_review_tree"):
            return
        df = getattr(self._prezzi_tab, "prezzi_df", None)
        table = prezzi_logic.category_review_table(df)
        self._review_tree.delete(*self._review_tree.get_children())
        for row in table.itertuples(index=False):
            self._review_tree.insert("", "end", iid=row.articolo, values=[
                row.articolo, row.descrizione, row.customer, row.colori,
                row.suggestion, row.evidence or "",
            ])
        if df is None or getattr(df, "empty", True):
            self._review_status.set(
                "Load a Listini file in the Prezzi tab to see articles that still need a Category."
            )
        elif table.empty:
            self._review_status.set("No articles need review.")
        else:
            self._review_status.set(
                f"{len(table)} article(s) have no Category. Select an article and assign a category. "
                "A suggestion is prefilled when the price evidence points to one category."
            )
        self._review_category.configure(values=prezzi_logic.known_categories(df))
        self._override_tree.delete(*self._override_tree.get_children())
        for article, category in sorted(prezzi_logic.load_category_overrides().items()):
            self._override_tree.insert("", "end", iid=article, values=[article, category])

    def _on_review_select(self, _event=None):
        selected = self._review_tree.selection()
        if selected:
            self._review_category.set(self._review_tree.set(selected[0], "suggestion"))

    def _assign_category(self):
        selected = self._review_tree.selection()
        if not selected:
            messagebox.showinfo("Category Review", "Select an article first.", parent=self)
            return
        article = selected[0]
        category = self._review_category.get().strip()
        if not category:
            messagebox.showwarning("Category Review", "Choose or type a category.", parent=self)
            return
        warning = prezzi_logic.override_warning(article, category)
        if warning and not messagebox.askyesno(
            "Check this category", f"{warning}\n\nAssign it anyway?", parent=self,
        ):
            return
        known = {name.casefold() for name in prezzi_logic.known_categories()}
        if category.casefold() not in known and not messagebox.askyesno(
            "New category", f"'{category}' is not an existing category.\n\nCreate it for {article}?",
            parent=self,
        ):
            return
        try:
            prezzi_logic.save_category_override(article, category)
        except (OSError, ValueError) as exc:
            logger.exception("Category Review: couldn't save the assignment")
            messagebox.showerror("Category Review", f"Couldn't save the assignment:\n{exc}", parent=self)
            return
        self._apply_category_change()

    def _remove_override(self):
        selected = self._override_tree.selection()
        if not selected:
            return
        article = selected[0]
        if not messagebox.askyesno(
            "Remove assignment", f"Remove the manual category for {article}?", parent=self,
        ):
            return
        try:
            prezzi_logic.remove_category_override(article)
        except OSError as exc:
            logger.exception("Category Review: couldn't remove the assignment")
            messagebox.showerror("Category Review", f"Couldn't remove the assignment:\n{exc}", parent=self)
            return
        self._apply_category_change()

    def _apply_category_change(self):
        self._on_data_changed()
        if self._prezzi_tab is not None:
            self._prezzi_tab.reapply_categories(on_done=self._refresh_category_review)
            return
        self._refresh_category_review()
