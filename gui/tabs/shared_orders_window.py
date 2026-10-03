from __future__ import annotations

import threading
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from calculate.order_report import export_report, pg_x_demand, report_path, schedule_due, schedule_stamp
from utility.email_compose import EmailTemplate, open_outlook_email, send_outlook_email
from utility.excel_io import safe_save_workbook
from utility.master_data import finished_articolo_for, raw_articolo_for
from utility.utils import bind_escape_to_close
from .biglietti_exports import (
    deduplicate_create_excel,
    load_articoli_marca_lookup,
    load_articoli_titolo_map,
    load_create_excel_records,
    sync_workbook_history,
)


class SharedOrdersWindowMixin:
    def _show_shared_orders(self):
        if not self.shared_excel_path or not self.shared_excel_path.is_file():
            return messagebox.showwarning("Missing Shared Excel", "Select the shared Excel file first.")
        if getattr(self, "_orders_loading", False):
            return
        self._orders_loading = True
        loading = tk.Toplevel(self)
        loading.title("Show Orders")
        loading.geometry("360x120")
        loading.resizable(False, False)
        loading.transient(self.winfo_toplevel())
        ttk.Label(loading, text="Loading orders...", anchor="center").pack(expand=True, fill="both", padx=20, pady=20)
        threading.Thread(target=self._load_shared_orders_worker, args=(self.shared_excel_path, loading), daemon=True).start()
    def _show_shipped_history(self):
        """List every archived order whose invoice/shipment already went
        out (SQLite history -- see utility.orders_db.list_shipped), with
        the same search/clear/sort/export toolbar as the other lists."""
        from utility.orders_db import list_shipped
        import json as _json
        try:
            rows = list_shipped()
        except Exception as exc:
            return messagebox.showerror("Extract Shipped/Invoiced History", str(exc))
        if not rows:
            return messagebox.showinfo("Extract Shipped/Invoiced History", "No shipped/invoiced orders in history yet.")

        parsed = []
        for row in rows:
            try:
                payload = _json.loads(row.get("order_json") or "{}")
            except (TypeError, ValueError):
                payload = {}
            payload = dict(payload)
            payload["Sheet"] = row.get("sheet_name", "")
            payload["Archived At"] = row.get("archived_at", "")
            parsed.append(payload)

        # Union of columns across every row, in first-seen order, with the
        # archival metadata columns pinned last.
        columns: list[str] = []
        for payload in parsed:
            for key in payload:
                if key not in columns:
                    columns.append(key)
        for pinned in ("Sheet", "Archived At"):
            if pinned in columns:
                columns.remove(pinned)
                columns.append(pinned)

        window = tk.Toplevel(self)
        bind_escape_to_close(window)
        window.title("Extract Shipped/Invoiced History")
        window.geometry("1180x620")
        window.minsize(850, 420)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text=f"{len(parsed)} shipped/invoiced row(s)", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=8)
        toolbar = ttk.Frame(window)
        toolbar.grid(row=0, column=0, sticky="e", padx=10, pady=8)
        search = tk.StringVar()
        ttk.Label(toolbar, text="Search:").pack(side="left")
        ttk.Entry(toolbar, textvariable=search, width=28).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Clear", command=lambda: search.set("")).pack(side="left", padx=(5, 0))

        tree = ttk.Treeview(window, columns=columns, show="headings")
        sort_state: dict[str, bool] = {}

        def sort_col(col):
            nonlocal parsed
            ascending = sort_state.get(col, True)
            parsed = sorted(parsed, key=lambda p: str(p.get(col, "")), reverse=not ascending)
            sort_state[col] = not ascending
            render()

        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: sort_col(c))
            tree.column(col, width=120, anchor="center")
        tree.grid(row=1, column=0, sticky="nsew", padx=10)
        vscroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        vscroll.grid(row=1, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(window, orient="horizontal", command=tree.xview)
        hscroll.grid(row=2, column=0, sticky="ew", padx=10)
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        def render(*_args):
            tree.delete(*tree.get_children())
            term = search.get().strip().casefold()
            for payload in parsed:
                values = tuple(str(payload.get(c, "")) for c in columns)
                if not term or term in " ".join(values).casefold():
                    tree.insert("", "end", values=values)
        search.trace_add("write", render)
        render()

        buttons = ttk.Frame(window, padding=10)
        buttons.grid(row=3, column=0, columnspan=2, sticky="ew")

        def export():
            path = filedialog.asksaveasfilename(
                title="Export to Excel", defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")], initialfile="Shipped_Invoiced_History.xlsx",
            )
            if not path:
                return
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "History"
                ws.append(columns)
                for payload in parsed:
                    ws.append([payload.get(c, "") for c in columns])
                safe_save_workbook(wb, path)
                wb.close()
            except Exception as exc:
                return messagebox.showerror("Export Failed", str(exc), parent=window)
            messagebox.showinfo("Export Completed", f"Exported to:\n{path}", parent=window)

        ttk.Button(buttons, text="Export Excel", command=export).pack(side="left")
        ttk.Button(buttons, text="Close", command=window.destroy).pack(side="right")
    def _show_pgx_report(self):
        """Show total open PG-X demand per Titolo and configure delivery."""
        if not self.shared_excel_path or not self.shared_excel_path.is_file():
            return messagebox.showwarning("Missing Shared Excel", "Select the shared Excel file first.")
        try:
            records = load_create_excel_records(self.shared_excel_path, sheet_name="PG-X")
            frame = pg_x_demand(records)
        except Exception as exc:
            return messagebox.showerror("PG-X Report Error", str(exc))
        if frame.empty:
            return messagebox.showinfo("PG-X Report", "There are no open PG-X rows.")

        window = tk.Toplevel(self)
        bind_escape_to_close(window)
        window.title("PG-X Orders Report")
        window.geometry("980x600")
        window.minsize(760, 420)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text="Open raw-yarn demand grouped by Titolo", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=10)
        table_frame = ttk.Frame(window, padding=(12, 0, 12, 8))
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1); table_frame.columnconfigure(0, weight=1)
        columns = list(frame.columns)
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=180 if col == "Titolo" else 120, anchor="center")
        for row in frame.itertuples(index=False, name=None):
            tree.insert("", "end", values=tuple(round(v, 2) if isinstance(v, float) else v for v in row))
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        settings = dict(self._prefs.get("pgx_report_schedule", {}) or {})
        form = ttk.LabelFrame(window, text="Email report schedule", padding=8)
        form.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        for col in range(7): form.columnconfigure(col, weight=1 if col == 1 else 0)
        recipient = tk.StringVar(value=str(settings.get("recipient", "")))
        frequency = tk.StringVar(value=str(settings.get("frequency", "daily")))
        send_time = tk.StringVar(value=str(settings.get("time", "08:00")))
        enabled = tk.BooleanVar(value=bool(settings.get("enabled", False)))
        ttk.Label(form, text="To:").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=recipient, width=34).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Label(form, text="Frequency:").grid(row=0, column=2, sticky="e", padx=(8, 3))
        ttk.Combobox(form, textvariable=frequency, values=("daily", "weekly"), state="readonly", width=10).grid(row=0, column=3, sticky="w")
        ttk.Label(form, text="Time (HH:MM):").grid(row=0, column=4, sticky="e", padx=(8, 3))
        ttk.Entry(form, textvariable=send_time, width=8).grid(row=0, column=5, sticky="w")
        ttk.Checkbutton(form, text="Enable", variable=enabled).grid(row=0, column=6, sticky="w", padx=8)

        def save_schedule():
            value = {"enabled": enabled.get(), "recipient": recipient.get().strip(), "frequency": frequency.get(), "time": send_time.get().strip(), "last_sent": settings.get("last_sent", "")}
            self._save_prefs(pgx_report_schedule=value)
            settings.update(value)
            messagebox.showinfo("PG-X Report", "Schedule saved. The application will prepare the Outlook report at the selected time.", parent=window)

        def prepare_email():
            path = export_report(frame, report_path(self.shared_excel_path))
            template = EmailTemplate(to=recipient.get().strip(), subject="PG-X Orders Report", body="Attached is the current PG-X raw-yarn demand report.")
            try:
                open_outlook_email(template, [path], self._sender_email)
            except Exception as exc:
                messagebox.showerror("Email", str(exc), parent=window)

        def send_now():
            if not recipient.get().strip():
                return messagebox.showwarning("Email", "Enter a recipient first.", parent=window)
            if not messagebox.askyesno("Send PG-X report", f"Send the report now to {recipient.get().strip()}?", parent=window):
                return
            try:
                path = export_report(frame, report_path(self.shared_excel_path))
                send_outlook_email(EmailTemplate(to=recipient.get().strip(), subject="PG-X Orders Report", body="Attached is the current PG-X raw-yarn demand report."), [path], self._sender_email)
                messagebox.showinfo("Email", "Report sent successfully.", parent=window)
            except Exception as exc:
                messagebox.showerror("Email", str(exc), parent=window)

        buttons = ttk.Frame(window, padding=(12, 0, 12, 12)); buttons.grid(row=3, column=0, sticky="ew")
        ttk.Button(buttons, text="Save Schedule", command=save_schedule).pack(side="left")
        ttk.Button(buttons, text="Prepare Email Now", command=prepare_email).pack(side="left", padx=8)
        ttk.Button(buttons, text="Send Email Now", command=send_now).pack(side="left")
        ttk.Button(buttons, text="Export Excel", command=lambda: messagebox.showinfo("PG-X Report", f"Saved to:\n{export_report(frame, report_path(self.shared_excel_path))}", parent=window)).pack(side="left")
        ttk.Button(buttons, text="Close", command=window.destroy).pack(side="right")
    def _run_pgx_report_schedule(self):
        """Prepare the configured report once per scheduled day."""
        try:
            schedule = dict(self._prefs.get("pgx_report_schedule", {}) or {})
            if schedule_due(schedule):
                records = load_create_excel_records(self.shared_excel_path, sheet_name="PG-X") if self.shared_excel_path and self.shared_excel_path.is_file() else []
                frame = pg_x_demand(records)
                if not frame.empty:
                    path = export_report(frame, report_path(self.shared_excel_path))
                    template = EmailTemplate(to=schedule["recipient"], subject="PG-X Orders Report", body="Attached is the scheduled PG-X raw-yarn demand report.")
                    send_outlook_email(template, [path], self._sender_email)
                    schedule["last_sent"] = schedule_stamp()
                    self._save_prefs(pgx_report_schedule=schedule)
        except Exception as exc:
            self._logger.exception("Scheduled PG-X report failed: %s", exc)
        finally:
            self.after(60000, self._run_pgx_report_schedule)
    def _load_shared_orders_worker(self, path: Path, loading: tk.Toplevel):
        try:
            deduplicate_create_excel(path)
            sync_workbook_history(path)
            datasets = {
                "Orders": load_create_excel_records(path, sheet_name="Orders"),
                "PG-X": load_create_excel_records(path, sheet_name="PG-X"),
            }
            self.after(0, lambda: self._finish_shared_orders_load(datasets, loading))
        except Exception as exc:
            self.after(0, lambda exc=exc: self._finish_shared_orders_error(exc, loading))
    def _finish_shared_orders_error(self, exc: Exception, loading: tk.Toplevel):
        self._orders_loading = False
        if loading.winfo_exists():
            loading.destroy()
        messagebox.showerror("Shared Excel Error", str(exc))
    def _finish_shared_orders_load(self, datasets, loading: tk.Toplevel):
        self._orders_loading = False
        if loading.winfo_exists():
            loading.destroy()
        if not any(datasets.values()):
            return messagebox.showinfo("Show Orders", "The shared Excel has no orders yet.")
        self._open_shared_orders_window(datasets)
    def _open_shared_orders_window(self, datasets):

        window = tk.Toplevel(self)
        bind_escape_to_close(window)
        window.title("Select Orders for Biglietti")
        window.geometry("1100x560")
        window.minsize(850, 420)
        window.resizable(True, True)
        window.grab_set()
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)

        sheets = ttk.Notebook(window)
        sheets.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 0))
        orders_page = ttk.Frame(sheets)
        pgx_page = ttk.Frame(sheets)
        sheets.add(orders_page, text="Orders")
        sheets.add(pgx_page, text="PG-X")

        frame = ttk.Frame(window, padding=10)
        frame.grid(row=1, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=0)
        frame.rowconfigure(1, weight=1)
        search_bar = ttk.Frame(frame)
        search_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        search_bar.columnconfigure(1, weight=1)
        ttk.Label(search_bar, text="🔎 Search:").grid(row=0, column=0, sticky="w")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_bar, textvariable=search_var)
        search_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(search_bar, text="Clear", width=10, command=lambda: search_var.set("")).grid(row=0, column=2, sticky="e")
        ttk.Label(search_bar, text="Partita GG:").grid(row=0, column=3, sticky="w", padx=(18, 4))
        partita_gg_var = tk.StringVar()
        ttk.Entry(search_bar, textvariable=partita_gg_var, width=14).grid(row=0, column=4, sticky="w")
        ttk.Button(search_bar, text="SAVE", width=10, command=lambda: save_pg_x()).grid(row=0, column=5, sticky="e", padx=(6, 0))
        ttk.Button(
            search_bar, text="⚡ Smart Auto-Assign",
            command=lambda: self._smart_auto_assign_pg_x(window, datasets, refresh_records, rebuild, partita_gg_var),
        ).grid(row=0, column=6, sticky="e", padx=(6, 0))
        ttk.Button(
            search_bar, text="📤 Upload PG-X GG",
            command=lambda: self._upload_pgx_gg_file(window, datasets, refresh_records, rebuild, partita_gg_var),
        ).grid(row=0, column=7, sticky="e", padx=(6, 0))
        excel_columns = (
            "Dispo/Riga", "Cliente", "Articolo", "Titolo", "Formato", "Ordine", "Codice",
            "Colore", "Rocche", "KG", "M/C", "Partita Col", "Consegna", "Commento",
            "Bagno", "Partita GG", "Delivery Date", "Partita MED", "Cliente MED", "POLMON",
            "Color Tube", "VMM22", "Prezzo", "Densita` (360-390)", "Print",
        )
        columns = ("select",) + tuple(f"excel_{index}" for index in range(len(excel_columns)))
        tree_style = ttk.Style(window)
        tree_style.configure("Orders.Treeview", rowheight=28, background="#ffffff", fieldbackground="#ffffff")
        tree_style.configure("Orders.Treeview.Heading", font=("Segoe UI", 9, "bold"))
        tree_style.configure("Danger.TButton", foreground="#ffffff", background="#c62828", font=("Segoe UI", 9, "bold"))
        tree_style.map("Danger.TButton", background=[("active", "#8e0000"), ("pressed", "#8e0000")], foreground=[("disabled", "#eeeeee"), ("!disabled", "#ffffff")])
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse", style="Orders.Treeview")
        tree.tag_configure("odd", background="#e8f1fa", foreground="#172b4d")
        tree.tag_configure("even", background="#ffffff", foreground="#172b4d")
        headings = {"select": "Print"}
        headings.update({f"excel_{index}": name for index, name in enumerate(excel_columns)})
        widths = {"select": 70}
        widths.update({f"excel_{index}": max(90, min(220, len(name) * 10 + 20)) for index, name in enumerate(excel_columns)})
        for column in columns:
            tree.heading(column, text=headings[column], command=lambda c=column: sort_by(c))
            tree.column(column, width=widths[column], anchor="center")
        tree.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        hscrollbar = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        hscrollbar.grid(row=2, column=0, sticky="ew")
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=hscrollbar.set)

        selected: dict[str, bool] = {}
        record_by_iid: dict[str, object] = {}
        sheet_by_iid: dict[str, str] = {}
        partita_col_tree_column = f"excel_{excel_columns.index('Partita Col')}"
        sort_state = {"column": partita_col_tree_column, "reverse": False}
        current_sheet = {"name": "Orders"}
        for sheet_name, sheet_records in datasets.items():
            for index, record in enumerate(sheet_records):
                iid = f"{sheet_name}-{index}"
                selected[iid] = False
                record_by_iid[iid] = record
                sheet_by_iid[iid] = sheet_name

        def refresh_records(new_datasets):
            datasets.clear()
            datasets.update(new_datasets)
            selected.clear()
            record_by_iid.clear()
            sheet_by_iid.clear()
            for sheet_name, sheet_records in datasets.items():
                for index, record in enumerate(sheet_records):
                    iid = f"{sheet_name}-{index}"
                    selected[iid] = False
                    record_by_iid[iid] = record
                    sheet_by_iid[iid] = sheet_name

        def values_for(record, iid):
            polmon = ""
            if record.customer_code == "3004":
                parts = [part.strip() for part in str(record.additional_raw or "").split("-")]
                polmon = parts[4] if len(parts) > 4 and "POLMON" in parts[4].upper() else ""
            values = {
                "Dispo/Riga": record.dispo, "Cliente": record.customer_name,
                "Articolo": record.article, "Titolo": record.title, "Formato": record.formato,
                "Ordine": record.order_no, "Codice": record.color_code, "Colore": record.color_name,
                "Rocche": record.quantity_cones, "KG": record.kg, "M/C": record.machine,
                "Partita Col": record.colored_batch, "Consegna": record.delivery,
                "Commento": record.commento, "Bagno": record.bagno, "Partita GG": record.raw_batch,
                "Delivery Date": record.delivery_date, "Partita MED": record.partita_med,
                "Cliente MED": record.cliente_med, "POLMON": polmon, "Color Tube": record.color_tube,
                "VMM22": record.vmm22, "Prezzo": record.prezzo, "Densita` (360-390)": record.densita,
                "Print": record.print_flag,
            }

            def display_value(name, value):
                if name not in {"Consegna", "Delivery Date"} or value in (None, ""):
                    return value
                if isinstance(value, datetime):
                    return value.strftime("%d/%m/%Y")
                if isinstance(value, date):
                    return value.strftime("%d/%m/%Y")
                text = str(value).strip()
                for parser in (
                    lambda: datetime.fromisoformat(text.replace("Z", "+00:00")),
                    lambda: datetime.strptime(text, "%Y/%m/%d"),
                    lambda: datetime.strptime(text, "%d/%m/%Y"),
                ):
                    try:
                        return parser().strftime("%d/%m/%Y")
                    except (ValueError, TypeError):
                        pass
                return text

            return tuple(
                ["☑" if selected[iid] else "☐"]
                + [display_value(name, values.get(name, "")) for name in excel_columns]
            )

        def sort_by(column):
            if sort_state["column"] == column:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["column"] = column
                sort_state["reverse"] = False
            rebuild()

        def rebuild(*_args):
            query = search_var.get().strip().casefold()
            for iid in tree.get_children(""):
                tree.delete(iid)
            visible_index = 0
            visible_rows = []
            for iid, record in record_by_iid.items():
                if sheet_by_iid[iid] != current_sheet["name"]:
                    continue
                searchable = " ".join(str(value or "") for value in values_for(record, iid)).casefold()
                if query and query not in searchable:
                    continue
                visible_rows.append((iid, record))
            if sort_state["column"] is not None:
                sort_index = columns.index(sort_state["column"])
                if sort_state["column"] == partita_col_tree_column:
                    def partita_sort_key(pair):
                        value = str(pair[1].colored_batch or "").strip().replace(",", ".")
                        try:
                            number = float(value)
                            return (False, number, "")
                        except ValueError:
                            return (True, 0, value.casefold())

                    visible_rows.sort(key=partita_sort_key, reverse=sort_state["reverse"])
                else:
                    visible_rows.sort(key=lambda pair: str(values_for(pair[1], pair[0])[sort_index] or "").casefold(), reverse=sort_state["reverse"])
            for iid, record in visible_rows:
                tree.insert("", "end", iid=iid, values=values_for(record, iid), tags=("odd" if visible_index % 2 else "even",))
                visible_index += 1

        search_var.trace_add("write", rebuild)
        rebuild()

        def switch_sheet(_event=None):
            current_sheet["name"] = "PG-X" if sheets.index(sheets.select()) == 1 else "Orders"
            search_var.set("")
            partita_gg_var.set("")
            rebuild()

        sheets.bind("<<NotebookTabChanged>>", switch_sheet)

        def save_pg_x(allow_article_mismatch=False):
            if current_sheet["name"] != "PG-X":
                return messagebox.showwarning("PG-X Only", "Partita GG can be saved only from the PG-X page.", parent=window)
            partita_col = search_var.get().strip()
            partita_gg = partita_gg_var.get().strip()
            if not partita_col or not partita_gg:
                return messagebox.showwarning("Missing Data", "Search for one Partita Col and enter its Partita GG.", parent=window)
            self._save_pg_x_in_background(window, partita_col, partita_gg, datasets, refresh_records, rebuild, partita_gg_var, allow_article_mismatch)

        def edit_pg_x_row(event):
            if current_sheet["name"] not in ("PG-X", "Orders"):
                return
            iid = tree.identify_row(event.y)
            if not iid or sheet_by_iid.get(iid) != current_sheet["name"]:
                return
            record = record_by_iid[iid]
            editor = tk.Toplevel(window)
            bind_escape_to_close(editor)
            editor.title(f"Edit {current_sheet['name']} Color")
            editor.geometry("470x350")
            editor.resizable(False, False)
            editor.transient(window)
            editor.grab_set()
            form = ttk.Frame(editor, padding=12)
            form.pack(fill="both", expand=True)
            form.columnconfigure(1, weight=1)
            ttk.Label(form, text="Partita Col:").grid(row=0, column=0, sticky="w", pady=5)
            ttk.Label(form, text=str(record.colored_batch)).grid(row=0, column=1, sticky="w", pady=5)
            fields = (
                ("Articolo", str(record.article or "")),
                ("Titolo", str(record.title or "")),
                ("Partita GG", str(record.raw_batch or "")),
                ("Bagno", str(record.bagno or "")),
                ("Rocche", "" if record.quantity_cones is None else str(record.quantity_cones)),
                ("M/C", str(record.machine or "")),
            )
            entries = {}
            print_var = tk.BooleanVar(value=str(record.print_flag or "").strip().casefold() in {"1", "yes", "si", "true", "☑"})
            machine_choices = tuple(str(machine) for machine in range(3, 13))
            for row_idx, (label, value) in enumerate(fields, start=1):
                ttk.Label(form, text=f"{label}:").grid(row=row_idx, column=0, sticky="w", pady=5)
                if label in {"Articolo", "Titolo", "Partita GG"}:
                    entry = ttk.Combobox(form, width=30, state="normal")
                elif label == "M/C":
                    entry = ttk.Combobox(form, width=30, values=machine_choices, state="normal")
                else:
                    entry = ttk.Entry(form, width=32)
                entry.grid(row=row_idx, column=1, sticky="ew", pady=5)
                entry.insert(0, value)
                entries[label] = entry

            status_row = len(fields) + 1
            ttk.Checkbutton(form, text="Print", variable=print_var).grid(row=status_row, column=1, sticky="w", pady=(2, 4))
            status_row += 1
            expected_raw = str(record.article or "").strip().upper()
            expected_raw = raw_articolo_for(expected_raw)
            availability_var = tk.StringVar(value=f"Expected raw article: {expected_raw} — loading Magazino...")
            availability_label = ttk.Label(form, textvariable=availability_var, foreground="#666666", wraplength=340)
            availability_label.grid(row=status_row, column=0, columnspan=2, sticky="w", pady=(2, 4))
            stock_by_partita = {}
            partita_choices_by_article = {}
            articolo_choices = set()
            titolo_by_articolo = {}
            articolo_by_titolo = {}
            titolo_choices = set()
            sync_state = {"busy": False}
            availability_state = {"ready": False, "valid": False, "article_mismatch": False}

            def normal_partita(value):
                text = str(value or "").strip()
                try:
                    number = float(text.replace(",", "."))
                    return str(int(number)) if number.is_integer() else str(number)
                except ValueError:
                    return text.casefold()

            def refresh_availability(*_args):
                key = normal_partita(entries["Partita GG"].get())
                article = entries["Articolo"].get().strip().upper()
                expected = raw_articolo_for(article)
                if not availability_state["ready"]:
                    availability_var.set(f"Expected raw article: {expected} — loading Magazino...")
                    return
                if not key:
                    availability_state["valid"] = False
                    availability_state["article_mismatch"] = False
                    availability_var.set("Partita GG is blank — the row will remain in PG-X")
                    availability_label.config(foreground="#666666")
                    return
                items = stock_by_partita.get(key, [])
                item = next((candidate for candidate in items if candidate[1] == expected), None)
                if not items:
                    availability_state["valid"] = False
                    availability_state["article_mismatch"] = False
                    availability_var.set(f"Expected raw article: {expected} — Partita not found")
                    availability_label.config(foreground="#c62828")
                else:
                    available, article = item if item is not None else items[0]
                    valid = item is not None
                    availability_state["valid"] = valid
                    availability_state["article_mismatch"] = not valid
                    availability_var.set(f"Available: {available:g} rocche — raw article: {article}")
                    availability_label.config(foreground="#2e7d32" if valid else "#c62828")
                    if not valid:
                        availability_var.set(f"Wrong article. Expected {expected}, found {article} ({available:g} rocche)")

            def refresh_partita_choices(*_args):
                article = entries["Articolo"].get().strip().upper()
                expected = raw_articolo_for(article)
                choices = sorted(
                    partita_choices_by_article.get(expected, set()),
                    key=lambda value: (normal_partita(value).casefold(), str(value)),
                )
                entries["Partita GG"]["values"] = choices

            def refresh_articolo_choices(*_args):
                entries["Articolo"]["values"] = sorted(articolo_choices, key=str.casefold)
                entries["Titolo"]["values"] = sorted(titolo_choices, key=str.casefold)

            def normalize_article(value):
                article = str(value or "").strip().upper()
                return "C" + article[1:] if article.startswith("G") else article

            def refresh_linked_article_title(_event=None):
                if sync_state["busy"]:
                    return
                sync_state["busy"] = True
                try:
                    article = normalize_article(entries["Articolo"].get())
                    title = titolo_by_articolo.get(article, "")
                    if title:
                        entries["Titolo"].set(title)
                finally:
                    sync_state["busy"] = False

            def refresh_linked_title_article(_event=None):
                if sync_state["busy"]:
                    return
                sync_state["busy"] = True
                try:
                    title = entries["Titolo"].get().strip()
                    article = articolo_by_titolo.get(title.casefold(), "")
                    if article:
                        entries["Articolo"].set(article)
                        refresh_partita_choices()
                        refresh_availability()
                finally:
                    sync_state["busy"] = False

            def load_stock():
                try:
                    _codes, _density, _vmm, _prices, summary = self._load_common_sources()
                    try:
                        source_map = load_articoli_marca_lookup() or load_articoli_titolo_map()
                        for source_article, source_title in source_map.items():
                            article_key = normalize_article(source_article)
                            title_text = str(source_title or "").strip()
                            if article_key and title_text:
                                titolo_by_articolo[article_key] = title_text
                                articolo_by_titolo.setdefault(title_text.casefold(), article_key)
                                titolo_choices.add(title_text)
                                articolo_choices.add(article_key)
                    except Exception:
                        pass
                    if summary is not None:
                        for row in summary.itertuples(index=False):
                            article = str(getattr(row, "articolo", "") or "").strip().upper()
                            partita = normal_partita(getattr(row, "partita", ""))
                            available = float(getattr(row, "mag_rocche", 0) or 0)
                            stock_by_partita.setdefault(partita, []).append((available, article))
                            if article:
                                articolo_choices.add(finished_articolo_for(article) if article.startswith("G") else article)
                            if partita:
                                partita_choices_by_article.setdefault(article, set()).add(partita)
                    self.after(0, lambda: (availability_state.update(ready=True), refresh_articolo_choices(), refresh_partita_choices(), refresh_availability()))
                except Exception as exc:
                    self.after(0, lambda exc=exc: availability_var.set(f"Magazino error: {exc}"))

            entries["Partita GG"].bind("<KeyRelease>", refresh_availability)
            entries["Partita GG"].bind("<<ComboboxSelected>>", refresh_availability)
            entries["Articolo"].bind("<KeyRelease>", refresh_articolo_choices)
            entries["Articolo"].bind("<KeyRelease>", refresh_partita_choices, add="+")
            entries["Articolo"].bind("<KeyRelease>", refresh_availability, add="+")
            entries["Articolo"].bind("<<ComboboxSelected>>", refresh_partita_choices)
            entries["Articolo"].bind("<<ComboboxSelected>>", refresh_availability, add="+")
            entries["Articolo"].bind("<KeyRelease>", refresh_linked_article_title, add="+")
            entries["Articolo"].bind("<<ComboboxSelected>>", refresh_linked_article_title, add="+")
            entries["Titolo"].bind("<KeyRelease>", refresh_linked_title_article)
            entries["Titolo"].bind("<<ComboboxSelected>>", refresh_linked_title_article)
            threading.Thread(target=load_stock, daemon=True).start()

            def save_from_editor():
                values = {label: entry.get().strip() for label, entry in entries.items()}
                try:
                    title_map = load_articoli_marca_lookup() or load_articoli_titolo_map()
                    article_key = values.get("Articolo", "").strip().upper()
                    title_key = raw_articolo_for(article_key) if article_key.startswith("C") else article_key
                    values["Titolo"] = title_map.get(article_key, title_map.get(title_key, values.get("Titolo", "")))
                except Exception:
                    pass
                values["Print"] = "☑" if print_var.get() else ""
                value = values["Partita GG"]
                if value and not availability_state["ready"]:
                    return messagebox.showwarning("Magazino", "Wait for Magazino availability to load.", parent=editor)
                if value:
                    refresh_availability()
                allow_mismatch = False
                if value and not availability_state["valid"]:
                    if not availability_state["article_mismatch"]:
                        return messagebox.showerror("Partita Not Found", "This Partita GG was not found in Magazino.", parent=editor)
                    allow_mismatch = messagebox.askyesno(
                        "Article Mismatch",
                        "The raw article is different from the color article. Do you want to save anyway?",
                        parent=editor,
                    )
                    if not allow_mismatch:
                        return
                editor.destroy()
                self._save_pg_x_details_in_background(
                    window, str(record.colored_batch), values, datasets, refresh_records,
                    rebuild, partita_gg_var, sheet_name=current_sheet["name"],
                    allow_article_mismatch=allow_mismatch,
                )

            ttk.Button(form, text="SAVE", command=save_from_editor).grid(row=status_row + 1, column=1, sticky="e", pady=(8, 0))
            entries["Articolo"].focus_set()

        tree.bind("<Double-1>", edit_pg_x_row)

        # toggle() is debounced (delayed, then cancelled if a Double-1
        # follows): every double-click first fires two ordinary
        # click-release events, so double-clicking a row to open its editor
        # in the checkbox column also toggled a checkbox -- and since the
        # two clicks can resolve to different rows, sometimes a neighbour.
        pending_toggle = {"after_id": None}

        def toggle(_event=None):
            iid = tree.identify_row(_event.y) if _event is not None else ""
            column = tree.identify_column(_event.x) if _event is not None else ""
            if not iid or column != "#1":
                return
            if pending_toggle["after_id"] is not None:
                tree.after_cancel(pending_toggle["after_id"])

            def apply_toggle():
                pending_toggle["after_id"] = None
                if iid not in selected or not tree.exists(iid):
                    return
                selected[iid] = not selected[iid]
                values = list(tree.item(iid, "values"))
                values[0] = "☑" if selected[iid] else "☐"
                tree.item(iid, values=values)
                tree.selection_set(iid)

            # Tk's double-click threshold is well under this delay, so a
            # genuine double-click always cancels this before it fires.
            pending_toggle["after_id"] = tree.after(300, apply_toggle)

        def cancel_pending_toggle(_event=None):
            if pending_toggle["after_id"] is not None:
                tree.after_cancel(pending_toggle["after_id"])
                pending_toggle["after_id"] = None

        tree.bind("<Double-1>", cancel_pending_toggle, add="+")
        tree.bind("<ButtonRelease-1>", toggle)
        actions = ttk.Frame(window, padding=(10, 0, 10, 10))
        actions.grid(row=2, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)
        button_bar = ttk.Frame(actions)
        button_bar.grid(row=0, column=0, sticky="ew")
        for col in range(7):
            button_bar.columnconfigure(col, weight=1)
        ttk.Label(actions, text="Click Print to select or unselect each color.").grid(row=1, column=0, sticky="w", pady=(5, 0))
        def pg_x_action(action):
            if current_sheet["name"] != "PG-X":
                return messagebox.showwarning("PG-X Only", "This action is available only on the PG-X page.", parent=window)
            selection = tree.selection()
            if not selection:
                return messagebox.showwarning("No Row Selected", "Click a PG-X row first.", parent=window)
            iid = selection[0]
            record = record_by_iid[iid]
            partita_col = str(record.colored_batch)
            if action == "delete" and not messagebox.askyesno(
                "Delete PG-X Row", f"Delete Partita Col {partita_col} from PG-X?", parent=window
            ):
                return
            raw_b = str(record.raw_batch or "").strip().upper().replace(" ", "")
            if action == "move" and raw_b in {"", "X", "PG-X", "PGX"} and not messagebox.askyesno(
                "No Raw Yarn Assigned",
                f"Partita Col {partita_col} still has no raw yarn (Partita GG) assigned.\n\n"
                "Send it to Orders anyway?",
                parent=window,
            ):
                return
            self._pg_x_row_action_in_background(
                window, action, partita_col, datasets, refresh_records, rebuild,
            )

        ttk.Button(button_bar, text="Send to Orders", command=lambda: pg_x_action("move"), width=16).grid(row=0, column=0, padx=3, sticky="ew")
        ttk.Button(button_bar, text="Delete", command=lambda: pg_x_action("delete"), width=12).grid(row=0, column=1, padx=3, sticky="ew")
        ttk.Button(
            button_bar, text="Delete Shipped Colors", style="Danger.TButton", width=18,
            command=lambda: self._delete_shipped_colors(
                window, current_sheet["name"], record_by_iid, sheet_by_iid,
                datasets, refresh_records, rebuild,
            ),
        ).grid(row=0, column=2, padx=3, sticky="ew")
        ttk.Button(
            button_bar, text="🖨 Print Assigned PG-X", width=19,
            command=lambda: self._print_assigned_pgx(window, current_sheet["name"], record_by_iid, sheet_by_iid),
        ).grid(row=0, column=3, padx=3, sticky="ew")
        ttk.Button(
            button_bar, text="🖨 Print Selected Biglietti", width=21,
            command=lambda: self._print_selected_shared(
                window, selected, record_by_iid,
                move_assigned_pg_x=current_sheet["name"] == "PG-X",
            ),
        ).grid(row=0, column=4, padx=3, sticky="ew")
