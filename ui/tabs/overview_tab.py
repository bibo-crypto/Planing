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

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except ImportError:  # pragma: no cover - matplotlib not installed yet
    _HAS_MPL = False

ARTICOLO_CLIENT_LABELS = {"G130": "Elvy", "G170": "Kamal"}
READY_MARK = "pronto da spedire"
SHORTAGE_MARK = "PG-X"


# ----------------------------------------------------------------------
# Generic "export any Treeview to Excel" helper — reusable elsewhere.
# ----------------------------------------------------------------------
def treeview_to_dataframe(tree: ttk.Treeview) -> pd.DataFrame:
    """Read a Treeview's current columns/rows (as displayed) into a DataFrame."""
    columns = tree["columns"]
    headers = [str(tree.heading(c)["text"] or c) for c in columns]
    rows = [tree.item(item)["values"] for item in tree.get_children()]
    return pd.DataFrame(rows, columns=headers)


def export_treeview_to_excel(tree: ttk.Treeview, default_filename: str, parent: tk.Widget | None = None) -> None:
    """Export a Treeview's currently displayed rows/columns to a new .xlsx file."""
    df = treeview_to_dataframe(tree)
    if df.empty:
        messagebox.showinfo("No data", "There's nothing to export.", parent=parent)
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
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append(list(row))
        for i, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col]])
            ws.column_dimensions[get_column_letter(i)].width = max(10, min(45, max_len + 2))
        ws.freeze_panes = "A2"
        wb.save(path)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Export failed", str(exc), parent=parent)
        return
    messagebox.showinfo("Done", f"Exported to:\n{path}", parent=parent)


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


class OverviewTab(ttk.Frame):
    """Dashboard: KPI cards + charts + detail lists, all read-only."""

    def __init__(self, master, situazione_tab, magazino_tab):
        super().__init__(master)
        self.situazione_tab = situazione_tab
        self.magazino_tab = magazino_tab
        self._build_ui()
        self.refresh()

    # ------------------------------------------------------------------
    # UI scaffolding
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", padx=8, pady=(8, 4))
        ttk.Button(toolbar, text="🔄 Refresh", command=self.refresh).pack(side="left")
        self._lbl_updated = ttk.Label(toolbar, text="", foreground="#667085")
        self._lbl_updated.pack(side="right")

        outer = ttk.Frame(self)
        outer.pack(side="top", fill="both", expand=True, padx=4, pady=4)

        canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._body = ttk.Frame(canvas)
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

        self._cards_frame = ttk.Frame(self._body)
        self._cards_frame.pack(side="top", fill="x", padx=8, pady=(4, 10))

        self._charts_frame = ttk.Frame(self._body)
        self._charts_frame.pack(side="top", fill="x", padx=8, pady=(0, 10))

        self._tables_frame = ttk.Frame(self._body)
        self._tables_frame.pack(side="top", fill="both", expand=True, padx=8, pady=(0, 10))

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def on_shown(self) -> None:
        """Call this when the tab becomes visible — keeps it auto-fresh."""
        self.refresh()

    def refresh(self) -> None:
        df = getattr(self.situazione_tab, "current_df", None)
        magazino = getattr(self.magazino_tab, "magazino_summary", None)
        df = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
        magazino = magazino.copy() if isinstance(magazino, pd.DataFrame) else pd.DataFrame()
        self._render(df, magazino)
        self._lbl_updated.config(text=f"Last updated: {pd.Timestamp.now():%Y-%m-%d %H:%M:%S}")

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------
    def _render(self, df: pd.DataFrame, magazino: pd.DataFrame) -> None:
        for frame in (self._cards_frame, self._charts_frame, self._tables_frame):
            for widget in frame.winfo_children():
                widget.destroy()

        if not df.empty:
            for col in ("cliente", "colore", "titolo", "comment", "new_comment"):
                df[col] = df.get(col, "").fillna("").astype(str).str.strip()

        shortage_df = (
            df[df["comment"].str.upper().str.startswith(SHORTAGE_MARK)]
            if not df.empty else df
        )
        ready_df = (
            df[df["new_comment"].str.lower().str.contains(READY_MARK)]
            if not df.empty else df
        )

        n_shortage_colors = shortage_df["colore"].nunique() if not shortage_df.empty else 0
        n_ready_colors = ready_df["colore"].nunique() if not ready_df.empty else 0
        total_yarn_kg = (
            magazino["mag_peso"].sum() if not magazino.empty and "mag_peso" in magazino else 0.0
        )

        self._add_card(0, "🧵 Raw yarn in warehouse", f"{total_yarn_kg:,.0f} kg")
        self._add_card(1, "📦 Colors ready to ship, no exit invoice", str(n_ready_colors))
        self._add_card(2, "⚠️ Colors needing yarn", str(n_shortage_colors))

        if _HAS_MPL:
            self._draw_yarn_by_client_chart(magazino)
            self._draw_grouped_chart(ready_df, "Ready-to-ship colors per client")
            self._draw_grouped_chart(shortage_df, "Colors needing yarn per client")
        else:
            ttk.Label(
                self._charts_frame,
                text="To enable charts: pip install -r requirements.txt (matplotlib)",
                foreground="#b42318",
            ).pack(pady=10)

        self._add_table(
            "Colors ready to ship (no exit invoice yet)",
            ready_df, ["cliente", "colore", "titolo", "partita", "rocche"],
            "ready_to_ship.xlsx",
        )
        self._add_table(
            "Colors needing yarn",
            shortage_df, ["cliente", "colore", "titolo", "rocche"],
            "yarn_needed.xlsx",
        )
        self._add_yarn_table(magazino)

    def _add_card(self, col: int, title: str, value: str) -> None:
        self._cards_frame.columnconfigure(col, weight=1)
        card = ttk.Frame(self._cards_frame, relief="solid", borderwidth=1, padding=12)
        card.grid(row=0, column=col, padx=6, sticky="nsew")
        ttk.Label(card, text=title, foreground="#667085", wraplength=200).pack(anchor="w")
        ttk.Label(card, text=value, font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(4, 0))

    def _bar_chart(self, title: str, series: pd.Series) -> None:
        fig = Figure(figsize=(4.0, 2.8), dpi=90)
        ax = fig.add_subplot(111)
        if series.empty:
            ax.text(0.5, 0.5, "No data", ha="center", va="center")
            ax.set_xticks([])
            ax.set_yticks([])
        else:
            series.plot(kind="bar", ax=ax, color="#1976D2")
            ax.tick_params(axis="x", rotation=0)
        ax.set_title(title, fontsize=10)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self._charts_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="left", padx=8, pady=4)

    def _draw_yarn_by_client_chart(self, magazino: pd.DataFrame) -> None:
        if magazino.empty or "articolo" not in magazino or "mag_peso" not in magazino:
            series = pd.Series(dtype=float)
        else:
            m = magazino.copy()
            m["client"] = (
                m["articolo"].astype(str).str[:4].map(ARTICOLO_CLIENT_LABELS)
                .fillna(m["articolo"])
            )
            series = m.groupby("client")["mag_peso"].sum()
        self._bar_chart("Raw yarn per client (kg)", series)

    def _draw_grouped_chart(self, source_df: pd.DataFrame, title: str) -> None:
        if source_df.empty or "cliente" not in source_df or "colore" not in source_df:
            series = pd.Series(dtype=float)
        else:
            series = source_df.groupby("cliente")["colore"].nunique()
        self._bar_chart(title, series)

    def _add_table(self, title: str, df: pd.DataFrame, cols: list[str], export_filename: str) -> None:
        frame = ttk.LabelFrame(self._tables_frame, text=title, padding=6)
        frame.pack(side="top", fill="both", expand=True, pady=6)

        header = ttk.Frame(frame)
        header.pack(side="top", fill="x", pady=(0, 4))

        available_cols = [c for c in cols if df.empty or c in df.columns] or cols
        tree = ttk.Treeview(frame, columns=available_cols, show="headings", height=6)
        for c in available_cols:
            tree.heading(c, text=c.capitalize())
            tree.column(c, width=120, anchor="center")
        tree.pack(fill="both", expand=True)
        if not df.empty:
            for _, row in df.iterrows():
                tree.insert("", "end", values=[row.get(c, "") for c in available_cols])

        ttk.Button(
            header, text="📤 Export to Excel",
            command=lambda: export_treeview_to_excel(tree, export_filename, parent=self),
        ).pack(side="right")
        attach_excel_export(tree, export_filename)

    def _add_yarn_table(self, magazino: pd.DataFrame) -> None:
        frame = ttk.LabelFrame(self._tables_frame, text="Raw yarn in warehouse, per client/partita", padding=6)
        frame.pack(side="top", fill="both", expand=True, pady=6)

        header = ttk.Frame(frame)
        header.pack(side="top", fill="x", pady=(0, 4))

        cols = ["client", "articolo", "partita", "mag_rocche", "mag_peso"]
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=6)
        for c in cols:
            tree.heading(c, text=c.capitalize())
            tree.column(c, width=120, anchor="center")
        tree.pack(fill="both", expand=True)
        if not magazino.empty:
            m = magazino.copy()
            m["client"] = m["articolo"].astype(str).str[:4].map(ARTICOLO_CLIENT_LABELS).fillna(m["articolo"])
            for _, row in m.iterrows():
                tree.insert("", "end", values=[row.get(c, "") for c in cols])

        ttk.Button(
            header, text="📤 Export to Excel",
            command=lambda: export_treeview_to_excel(tree, "raw_yarn_by_client.xlsx", parent=self),
        ).pack(side="right")
        attach_excel_export(tree, "raw_yarn_by_client.xlsx")
