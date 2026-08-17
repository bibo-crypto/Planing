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
from openpyxl.worksheet.table import Table, TableStyleInfo

from utils import parse_number
import situazione_logic as business_logic

ARTICOLO_CLIENT_LABELS = {"G130": "Elvy", "G170": "Kamal"}
READY_MARK = "pronto da spedire"
SHORTAGE_MARK = "PG-X"
DELIVERY_ALERT_DAYS = 3
QUALITY_READY_CODES = {"AA", "AC", "AU", "AT"}
HEADERS_IT = {
    "cliente": "Cliente", "colore": "Colore", "titolo": "Titolo",
    "articolo": "Articolo", "codice": "Codice", "ordine": "Ordine",
    "riga": "Riga", "data": "Data", "consegna": "Consegna",
    "partita": "Partita", "rocche": "Rocche", "mc": "M/C",
    "comment": "Commento", "raw_yarn_match": "Filato Disponibile",
    "cq": "C.Q", "tinto": "Tinto", "bagno": "Bagno",
    "old_comment": "Vecchio commento", "new_comment": "Nuovo commento",
    "planedate": "PlaneDate", "data_qualita": "Data Qualità",
    "data_uscita": "Data Uscita", "custom": "Controllo",
    "days_in_qc": "Giorni in C.Q",
    "days_to_delivery": "Giorni alla consegna",
}
CHECK_COLUMNS = [
    "cliente", "articolo", "titolo", "codice", "colore", "ordine", "riga",
    "data", "consegna", "partita", "rocche", "mc", "comment", "raw_yarn_match",
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


def export_treeview_to_excel(tree: ttk.Treeview, default_filename: str, parent: tk.Widget | None = None) -> None:
    """Export a Treeview's currently displayed rows/columns to a new .xlsx file."""
    df = treeview_to_dataframe(tree)
    if df.empty:
        messagebox.showinfo("Nessun dato", "Non ci sono dati da esportare.", parent=parent)
        return
    path = filedialog.asksaveasfilename(
        title="Esporta in Excel",
        defaultextension=".xlsx",
        filetypes=[("File Excel", "*.xlsx")],
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
        wb.save(path)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Esportazione non riuscita", str(exc), parent=parent)
        return
    messagebox.showinfo("Completato", f"Esportato in:\n{path}", parent=parent)


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
                values.append(None if pd.isna(parsed) else parsed.to_pydatetime())
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
            for cell in ws[letter][1:]:
                cell.number_format = "yyyy-mm-dd"
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
        label="📤 Esporta in Excel",
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

    AUTO_REFRESH_MS = 5000

    def __init__(self, master, situazione_tab, magazino_tab):
        super().__init__(master)
        self.situazione_tab = situazione_tab
        self.magazino_tab = magazino_tab
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
        style.configure("Overview.CardTitle.TLabel", background="#ffffff", foreground="#667085",
                        font=("Segoe UI", 9))
        style.configure("Overview.CardValue.TLabel", background="#ffffff", foreground="#16324f",
                        font=("Segoe UI", 19, "bold"))
        style.configure("Overview.AlertCard.TFrame", background="#fff7ed", relief="solid", borderwidth=1)
        style.configure("Overview.AlertCardTitle.TLabel", background="#fff7ed", foreground="#c2410c",
                        font=("Segoe UI", 9, "bold"))
        style.configure("Overview.AlertCardValue.TLabel", background="#fff7ed", foreground="#9a3412",
                        font=("Segoe UI", 19, "bold"))

        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", padx=8, pady=(8, 4))
        ttk.Button(toolbar, text="🔄 Refresh", command=self.refresh).pack(side="left")
        ttk.Label(toolbar, text="Riepilogo operativo", foreground="#16324f",
                  font=("Segoe UI", 13, "bold")).pack(side="left", padx=(14, 0))
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

        self._tables_frame = ttk.Frame(self._body)
        self._tables_frame.pack(side="top", fill="both", expand=True, padx=8, pady=(0, 10))
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
        return (
            getattr(self.situazione_tab, "_data_revision", 0),
            getattr(self.situazione_tab, "_copertura_revision", 0),
            getattr(self.magazino_tab, "_data_revision", 0),
        )

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
    def _render(self, df: pd.DataFrame, magazino: pd.DataFrame) -> None:
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
        if not df.empty and "consegna" in df.columns:
            due_dates = pd.to_datetime(df["consegna"], errors="coerce")
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
        total_yarn_kg = (
            magazino["mag_peso"].sum() if not magazino.empty and "mag_peso" in magazino else 0.0
        )

        self._add_card(0, "📋 Totale partite", f"{len(df):,}")
        self._add_card(1, "🧵 Filato disponibile", f"{total_yarn_kg:,.0f} kg")
        self._add_card(2, "📦 Pronte da spedire", str(n_ready_colors))
        self._add_card(3, "⚠️ In attesa di filato", str(n_shortage_colors))
        self._add_card(4, "⚠️ Ritardo in Q.C", f"{len(check_df):,}", alert=True)
        self._add_card(5, "📅 Ritardo Consegna", f"{len(delivery_df):,}", alert=True)

        self._add_machine_summary(df)

        self._add_table(
            "Colori pronti da spedire (senza uscita)",
            ready_df, ["cliente", "codice", "colore", "titolo", "partita", "rocche", "mc", "bagno"],
            "colori_pronti.xlsx",
        )
        self._add_table(
            "Colori in attesa di filato",
            shortage_df, ["cliente", "codice", "colore", "titolo", "ordine", "riga", "partita", "rocche",
                          "comment", "raw_yarn_match"],
            "colori_filato_mancante.xlsx",
        )
        self._add_table(
            "Ritardo in Q.C (C.Q = OO, oltre 4 giorni)",
            check_df,
            ["cliente", "codice", "colore", "titolo", "partita", "rocche",
             "days_in_qc", "new_comment"],
            "ritardo_in_qc.xlsx",
        )
        self._add_table(
            f"📅 Ritardo consegna / entro {DELIVERY_ALERT_DAYS} giorni (C.Q = OO)",
            delivery_df, ["cliente", "codice", "colore", "titolo", "partita", "rocche",
                          "consegna", "days_to_delivery", "new_comment"],
            "ordini_urgenti_consegna.xlsx",
        )

    def _add_machine_summary(self, situation_df: pd.DataFrame) -> None:
        """Render Copertura machine totals as compact Overview cards."""
        copertura = getattr(self.situazione_tab, "loaded_frames", {}).get("copertura")
        counts = {machine: 0 for machine in range(3, 13)}
        if isinstance(copertura, pd.DataFrame) and not copertura.empty and not situation_df.empty:
            if {"bagno", "machine"}.issubset(copertura.columns) and "bagno" in situation_df.columns:
                def key(value):
                    digits = "".join(ch for ch in str(value or "") if ch.isdigit()).lstrip("0")
                    return digits or str(value or "").strip().casefold()

                left = situation_df.copy()
                right = copertura.copy()
                left["_bagno_key"] = left["bagno"].map(key)
                right["_bagno_key"] = right["bagno"].map(key)
                joined = left.merge(right[["_bagno_key", "machine"]].drop_duplicates("_bagno_key"), on="_bagno_key", how="inner")

                # Same parser the Copertura popup uses (business_logic.
                # machine_number_from_label) -- kept in one place so a
                # "5" vs "3305" style label is read identically in both
                # views instead of silently under-counting here.
                joined["_machine_number"] = joined["machine"].map(business_logic.machine_number_from_label)
                valid = joined[joined["_machine_number"].between(3, 12, inclusive="both")]
                counts.update(valid["_machine_number"].value_counts().to_dict())

        frame = tk.LabelFrame(
            self._cards_frame, text="Copertura macchine", bg="#f8fafc", fg="#16324f",
            font=("Segoe UI", 10, "bold"), padx=8, pady=6,
        )
        frame.grid(row=2, column=0, columnspan=3, padx=6, pady=(0, 10), sticky="ew")
        self._cards_frame.columnconfigure(0, weight=1)
        self._cards_frame.columnconfigure(1, weight=1)
        self._cards_frame.columnconfigure(2, weight=1)

        covered_until = business_logic.machine_coverage_until

        for index, machine in enumerate(range(3, 13)):
            count = int(counts[machine])
            empty = count == 0
            card = tk.Frame(
                frame, bg="#fee2e2" if empty else "#eaf2f8",
                highlightbackground="#ef4444" if empty else "#cbd5e1",
                highlightthickness=1, width=106, height=76,
            )
            card.grid(row=0, column=index, padx=3, pady=3, sticky="ew")
            card.grid_propagate(False)
            frame.columnconfigure(index, weight=1, minsize=106)
            tk.Label(card, text=f"M{machine}", bg=card["bg"], fg="#991b1b" if empty else "#16324f",
                     font=("Segoe UI", 9, "bold"), anchor="center").pack(fill="x")
            tk.Label(card, text=f"{count} colori", bg=card["bg"], fg="#991b1b" if empty else "#344054",
                     font=("Segoe UI", 11, "bold"), anchor="center").pack(fill="x")
            tk.Label(card, text=f"Fino al: {covered_until(count)}", bg=card["bg"],
                     fg="#991b1b" if empty else "#667085",
                     font=("Segoe UI", 8, "bold"), anchor="center").pack(fill="x", pady=(2, 0))

    def _add_card(self, col: int, title: str, value: str, alert: bool = False) -> None:
        self._cards_frame.columnconfigure(col, weight=1)
        card_style = "Overview.AlertCard.TFrame" if alert else "Overview.Card.TFrame"
        title_style = "Overview.AlertCardTitle.TLabel" if alert else "Overview.CardTitle.TLabel"
        value_style = "Overview.AlertCardValue.TLabel" if alert else "Overview.CardValue.TLabel"
        card = ttk.Frame(self._cards_frame, style=card_style, padding=12)
        card.grid(row=col // 3, column=col % 3, padx=6, pady=4, sticky="nsew")
        ttk.Label(card, text=title, style=title_style, wraplength=180).pack(anchor="w")
        ttk.Label(card, text=value, style=value_style).pack(anchor="w", pady=(4, 0))

    def _add_table(self, title: str, df: pd.DataFrame, cols: list[str], export_filename: str) -> None:
        frame = ttk.LabelFrame(self._tables_frame, text=title, padding=6)
        col = self._table_count % 2
        row = self._table_count // 2
        self._tables_frame.columnconfigure(col, weight=1)
        frame.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
        self._table_count += 1

        header = ttk.Frame(frame)
        header.pack(side="top", fill="x", pady=(0, 4))

        available_cols = [c for c in cols if df.empty or c in df.columns] or cols
        search_var = tk.StringVar()
        ttk.Label(header, text="Cerca:").pack(side="left", padx=(2, 5))
        search_entry = ttk.Entry(header, textvariable=search_var, width=28)
        search_entry.pack(side="left", padx=(0, 8))

        table_area = ttk.Frame(frame)
        table_area.pack(fill="both", expand=True)
        tree = ttk.Treeview(table_area, columns=available_cols, show="headings", height=7)
        for c in available_cols:
            tree.heading(c, text=HEADERS_IT.get(c, c.capitalize()))
            tree.column(c, width=120, anchor="center")
        vsb = ttk.Scrollbar(table_area, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_area, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_area.rowconfigure(0, weight=1)
        table_area.columnconfigure(0, weight=1)

        def render_filtered_rows(*_args):
            tree.delete(*tree.get_children())
            visible = df
            query = search_var.get().strip()
            if query and not df.empty:
                mask = df[available_cols].astype(str).apply(
                    lambda column: column.str.contains(query, case=False, regex=False, na=False)
                ).any(axis=1)
                visible = df.loc[mask]
            for _, row in visible.iterrows():
                tree.insert("", "end", values=[row.get(c, "") for c in available_cols])

        search_var.trace_add("write", render_filtered_rows)
        render_filtered_rows()

        ttk.Button(
            header, text="📤 Esporta in Excel",
            command=lambda: export_treeview_to_excel(tree, export_filename, parent=self),
        ).pack(side="right")
        attach_excel_export(tree, export_filename)
