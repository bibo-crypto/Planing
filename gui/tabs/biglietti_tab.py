"""Tkinter page for extracting Biglietti, customer workbook and Filato."""

from __future__ import annotations

import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

def _lazy_exporter_call(function_name):
    """Load the heavy exporter module only when a Biglietti action needs it."""
    def call(*args, **kwargs):
        from exporters.biglietti_exporter import __dict__ as exporter_namespace
        return exporter_namespace[function_name](*args, **kwargs)
    return call


for _exporter_function in (
    "build_el_kamal_stem", "build_output_stem", "detect_order_format", "enrich_records",
    "export_word", "export_workbook", "load_articoli_marca_lookup",
    "load_articoli_titolo_map", "load_densita_query", "load_el_kamal_order", "load_order",
    "load_prezzo_lookup", "load_vmm22_ratio_from_magazino", "_filato_rows",
    "append_create_excel", "load_create_excel_records", "save_pg_x_partita",
    "update_pg_x_row", "update_order_row", "move_pg_x_to_orders", "delete_pg_x_row",
    "delete_shipped_shared_rows",
):
    globals()[_exporter_function] = _lazy_exporter_call(_exporter_function)
from utility.articoli_cache import load_articoli_cache, save_articoli_cache
from utility.densita_cache import load_densita_cache, save_densita_cache
from utility.magazino_cache import load_magazino_cache
from utility.prezzi_cache import load_prezzi_cache
from utility.path_manager import save_source, source_path
from utility.email_compose import (
    EmailAutocomplete,
    EmailTemplate,
    extract_email_addresses,
    get_outlook_accounts,
    open_outlook_account_setup,
    open_outlook_email,
)
from utility.utils import keep_window_on_top, lazy_call

# Filato x Tinturia here follows the same convention as Ordine Kamal/Ordine
# ELVY: a fixed filename, cleared and rewritten with only this run's rows
# every time (see pipelines.ordini_elvy.export_filato_full's docstring) --
# not the old per-order-named "{order file}_Filato.xlsx" snapshot.
export_filato_full = lazy_call("pipelines.ordini_elvy", "export_filato_full")
RawYarnMatch = lazy_call("pipelines.ordini_elvy", "RawYarnMatch")


def _split_pg_x_records(records):
    """Return (ready, pg_x) records for separate Word ticket files."""
    ready, pg_x = [], []
    for record in records:
        raw_batch = str(record.raw_batch or "").strip().upper().replace(" ", "")
        (pg_x if raw_batch in {"", "X", "PG-X", "PGX"} else ready).append(record)
    return ready, pg_x


class BigliettiTab(ttk.Frame):
    def __init__(self, parent: tk.Misc, prefs: dict, save_prefs, logger, on_shared_cache_changed=None):
        super().__init__(parent, padding=12)
        self._prefs = prefs
        self._save_prefs = save_prefs
        self._logger = logger
        self._on_shared_cache_changed = on_shared_cache_changed
        self.data_path: Path | None = None
        self.dispo_path: Path | None = None
        self.template_path: Path | None = None
        self.elvy_output_dir: Path | None = None
        self.med_output_dir: Path | None = None
        self.el_kamal_output_dir: Path | None = None
        self.filato_output_dir: Path | None = None
        self.shared_excel_path: Path | None = None
        self.articoli_path: Path | None = None
        self.densita_path: Path | None = None
        self.filato_enabled = tk.BooleanVar(value=False)
        self.email_templates: dict[str, EmailTemplate] = {
            "elvy": EmailTemplate(), "med": EmailTemplate(), "el_kamal": EmailTemplate(),
        }
        self._sender_email = ""
        self._last_client_files: dict[str, list[Path]] = {"elvy": [], "med": [], "el_kamal": []}
        self._build()
        self._restore()

    def _build(self):
        # This page is taller than the available client area on smaller
        # screens. Keep the page itself scrollable so the lower controls are
        # reachable without forcing the whole application to a fixed size.
        self._canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self._scrollbar = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._scrollbar.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        self._scrollbar.pack(side="right", fill="y")
        content = ttk.Frame(self._canvas, padding=12)
        self._canvas_window = self._canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda _event: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda event: self._canvas.itemconfigure(self._canvas_window, width=event.width))
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        content.columnconfigure(0, weight=1)
        style = ttk.Style(self)
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("PageTitle.TLabel", font=("Segoe UI", 17, "bold"), foreground="#16324F")
        style.configure("PageSubtitle.TLabel", font=("Segoe UI", 9), foreground="#5B6B7A")
        style.configure("Section.TLabelframe", padding=10)
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#16324F")
        style.configure("Hint.TLabel", font=("Segoe UI", 8), foreground="#64748B")
        header_frame = ttk.Frame(content)
        header_frame.pack(fill="x", padx=4, pady=(0, 10))
        ttk.Label(header_frame, text="Create Excel + Biglietti", style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(
            header_frame,
            text="Convert order data into dyeing tickets and a shared Excel workbook for ELVY, MED, and EL KAMAL.",
            style="PageSubtitle.TLabel", wraplength=950,
        ).pack(anchor="w", pady=(3, 0))

        input_box = ttk.LabelFrame(content, text=" 1. Main input files (required) ", style="Section.TLabelframe")
        input_box.pack(fill="x", padx=4, pady=(0, 10))
        input_box.columnconfigure(1, weight=1)
        ttk.Button(input_box, text="📂  Select Order Data", command=self._pick_data, width=24).grid(row=0, column=0, sticky="w", pady=4)
        self.biglietti_data_path_label = ttk.Label(input_box, text="No file selected", foreground="grey", anchor="w")
        self.biglietti_data_path_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)
        ttk.Button(input_box, text="📂  Select Dispo-Bagno", command=self._pick_dispo, width=24).grid(row=1, column=0, sticky="w", pady=4)
        self.biglietti_dispo_path_label = ttk.Label(input_box, text="No file selected (optional if embedded)", foreground="grey", anchor="w")
        self.biglietti_dispo_path_label.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)

        template_box = ttk.LabelFrame(content, text=" 2. Ticket template ", style="Section.TLabelframe")
        template_box.pack(fill="x", padx=4, pady=(0, 10))
        template_box.columnconfigure(1, weight=1)
        ttk.Button(template_box, text="📄  Select Forma Biglietti", command=self._pick_template, width=24).grid(row=0, column=0, sticky="w", pady=4)
        self.template_path_label = ttk.Label(template_box, text="Searching for Biglietti.docx...", foreground="grey", anchor="w")
        self.template_path_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)

        opt_box = ttk.LabelFrame(content, text=" 3. Optional shared data sources ", style="Section.TLabelframe")
        opt_box.pack(fill="x", padx=4, pady=(0, 10))
        opt_box.columnconfigure(1, weight=1)
        self.articoli_btn = ttk.Button(opt_box, text="📂  Articles (Titolo)", command=self._pick_articoli, width=24)
        self.articoli_btn.grid(row=0, column=0, sticky="w", pady=4)
        self.articoli_label = ttk.Label(opt_box, text="Shared with Situation tab", foreground="grey", anchor="w")
        self.articoli_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)
        self.densita_btn = ttk.Button(opt_box, text="📂  Density Query", command=self._pick_densita, width=24)
        self.densita_btn.grid(row=1, column=0, sticky="w", pady=4)
        self.densita_label = ttk.Label(opt_box, text="Not selected — KG remains raw yarn weight", foreground="grey", anchor="w")
        self.densita_label.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)
        info_frame = ttk.Frame(opt_box)
        info_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        info_frame.columnconfigure(0, weight=1)
        info_frame.columnconfigure(1, weight=1)
        self.vmm_label = ttk.Label(info_frame, text="• VMM22: Uses Yarn Warehouse (Magazino Filato) loaded in app", foreground="#555555", font=("Segoe UI", 8))
        self.vmm_label.grid(row=0, column=0, sticky="w")
        self.prezzi_label = ttk.Label(info_frame, text="• Price: Uses Price List (Listini) loaded in Prices / Situation", foreground="#555555", font=("Segoe UI", 8))
        self.prezzi_label.grid(row=0, column=1, sticky="w")

        export_box = ttk.LabelFrame(content, text=" 4. Client output folders and email actions ", style="Section.TLabelframe")
        export_box.pack(fill="x", padx=4, pady=(0, 10))
        export_box.columnconfigure(1, weight=1)
        ttk.Label(export_box, text="Client / destination", style="Hint.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(export_box, text="Selected folder", style="Hint.TLabel").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 4))
        ttk.Label(export_box, text="Client-specific actions", style="Hint.TLabel").grid(row=0, column=2, sticky="e", padx=(10, 0), pady=(0, 4))
        self._build_folder_row(export_box, 1, "elvy", "📁  ELVY Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 2, "med", "📁  MED Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 3, "el_kamal", "📁  EL KAMAL Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 4, "filato", "📁  Filato Output Folder", with_checkbox=True)

        shared_box = ttk.LabelFrame(content, text=" 5. Shared Create Excel workbook ", style="Section.TLabelframe")
        shared_box.pack(fill="x", padx=4, pady=(0, 10))
        shared_box.columnconfigure(1, weight=1)
        shared_buttons = ttk.Frame(shared_box)
        shared_buttons.grid(row=0, column=0, sticky="w", pady=4)
        ttk.Button(shared_buttons, text="📂 Select Existing", command=self._pick_shared_excel, width=18).pack(side="left")
        ttk.Button(shared_buttons, text="➕ Create New", command=self._create_shared_excel, width=15).pack(side="left", padx=(6, 0))
        self.shared_excel_label = ttk.Label(shared_box, text="Not selected — Convert will append here", foreground="grey", anchor="w")
        self.shared_excel_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)
        ttk.Button(shared_box, text="📋  Show Orders", command=self._show_shared_orders, width=24).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        action_box = ttk.LabelFrame(content, text=" 6. Generate ", style="Section.TLabelframe")
        action_box.pack(fill="x", padx=4, pady=(8, 4))
        action_box.columnconfigure(0, weight=1)
        self.convert_btn = ttk.Button(action_box, text="⚡  Convert & Generate Tickets (Excel + Word)", command=self._run_convert, width=42, style="Bold.TButton")
        self.convert_btn.pack(side="top", anchor="w", pady=(0, 8))
        self.status = ttk.Label(action_box, text="● Ready", font=("Segoe UI", 9, "bold"), foreground="#2E7D32", anchor="w", wraplength=950)
        self.status.pack(fill="x")

    def _on_mousewheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        if self.winfo_exists() and widget is not None and str(widget).startswith(str(self._w)):
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_folder_row(self, parent, row_idx, kind, btn_text, with_checkbox=False, show_email_buttons=False):
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=row_idx, column=0, sticky="w", pady=4)
        ttk.Button(btn_frame, text=btn_text, command=lambda: self._pick_output_dir(kind), width=24).pack(side="left")
        if with_checkbox:
            ttk.Checkbutton(btn_frame, text="Generate Raw Yarn (Filato)", variable=self.filato_enabled).pack(side="left", padx=(10, 0))
        lab = ttk.Label(parent, text="No folder selected", foreground="grey", anchor="w")
        lab.grid(row=row_idx, column=1, sticky="ew", padx=(10, 0), pady=4)
        setattr(self, f"{kind}_dir_label", lab)
        if show_email_buttons:
            client_name = {"elvy": "ELVY", "med": "MED", "el_kamal": "EL KAMAL"}.get(kind, kind.upper())
            email_frame = ttk.Frame(parent)
            email_frame.grid(row=row_idx, column=2, sticky="e", padx=(10, 0), pady=4)
            ttk.Button(
                email_frame, text=f"✉ {client_name} Template", width=18,
                command=lambda: self._open_email_template_editor(kind),
            ).pack(side="left", padx=(0, 4))
            ttk.Button(
                email_frame, text=f"📧 {client_name} Prepare Email", width=22,
                command=lambda: self._prepare_email(kind),
            ).pack(side="left")

    def _open_email_template_editor(self, kind: str) -> None:
        """TO/CC/SUBJECT/body editor for one client's email -- saved so
        "Prepare Email" can reuse it on every future Convert, not just this
        session."""
        template = self.email_templates.get(kind, EmailTemplate())
        window = tk.Toplevel(self)
        keep_window_on_top(window)
        window.title(f"Email Template — {kind.upper().replace('_', ' ')}")
        window.geometry("520x440")
        window.minsize(460, 380)

        form = ttk.Frame(window, padding=10)
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        to_var = tk.StringVar(value=template.to)
        cc_var = tk.StringVar(value=template.cc)
        subject_var = tk.StringVar(value=template.subject)
        email_values = self._saved_email_addresses()

        ttk.Label(form, text="To:").grid(row=0, column=0, sticky="w", pady=4)
        to_entry = EmailAutocomplete(form, values=email_values, textvariable=to_var)
        to_entry.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Cc:").grid(row=1, column=0, sticky="w", pady=4)
        cc_entry = EmailAutocomplete(form, values=email_values, textvariable=cc_var)
        cc_entry.grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Subject:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=subject_var).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Body:").grid(row=3, column=0, sticky="nw", pady=4)
        body_text = tk.Text(form, height=12, wrap="word")
        body_text.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(0, 4))
        body_text.insert("1.0", template.body)
        form.rowconfigure(4, weight=1)

        sender_text = self._sender_email or "Default Outlook account"
        sender_row = ttk.Frame(form)
        sender_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        sender_row.columnconfigure(0, weight=1)
        sender_label = ttk.Label(sender_row, text=f"Sender: {sender_text}", foreground="#555555")
        sender_label.grid(row=0, column=0, sticky="w")
        ttk.Button(
            sender_row, text="Change sender email", width=22,
            command=lambda: self._choose_sender_email(sender_label),
        ).grid(row=0, column=1, sticky="e")

        ttk.Label(
            form, text="Separate multiple recipients with a semicolon (a@x.com; b@y.com).",
            foreground="grey", font=("Segoe UI", 8),
        ).grid(row=6, column=0, columnspan=2, sticky="w")

        def save():
            self.email_templates[kind] = EmailTemplate(
                to=to_var.get().strip(), cc=cc_var.get().strip(),
                subject=subject_var.get().strip(), body=body_text.get("1.0", "end-1c"),
            )
            self._remember_email_addresses(to_var.get(), cc_var.get())
            self._save_prefs(biglietti_email_templates={
                k: t.to_dict() for k, t in self.email_templates.items()
            })
            window.destroy()

        buttons = ttk.Frame(form)
        buttons.grid(row=7, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=window.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(buttons, text="Save", command=save).pack(side="right")

    def _prepare_email(self, kind: str) -> None:
        """Open an Outlook compose window for *kind*'s saved template, with
        that client's files from the most recent Convert run attached."""
        template = self.email_templates.get(kind)
        if template is None or template.is_blank():
            messagebox.showwarning(
                "Email Template Not Set",
                f"Set up the email template for {kind.upper().replace('_', ' ')} first "
                "(the ✉ Template button next to it).",
            )
            return
        files = [p for p in self._last_client_files.get(kind, []) if p and Path(p).is_file()]
        if not files:
            messagebox.showwarning(
                "No Files Yet",
                f"No {kind.upper().replace('_', ' ')} files from a Convert run yet in this "
                "session — run Convert first, then Prepare Email.",
            )
            return
        try:
            open_outlook_email(template, files, sender_email=self._sender_email)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Could Not Open Email", str(exc))

    def _saved_email_addresses(self) -> list[str]:
        values = self._prefs.get("email_addresses", [])
        return [str(value) for value in values if str(value).strip()] if isinstance(values, list) else []

    def _remember_email_addresses(self, *fields: str) -> None:
        merged = self._saved_email_addresses()
        for address in extract_email_addresses(*fields):
            if not any(address.casefold() == old.casefold() for old in merged):
                merged.append(address)
        self._save_prefs(email_addresses=merged[-200:])

    def _choose_sender_email(self, label=None) -> None:
        try:
            accounts = get_outlook_accounts()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Outlook Accounts", str(exc))
            return
        if not accounts:
            messagebox.showwarning("Outlook Accounts", "No connected Outlook email accounts were found.")
            return
        window = tk.Toplevel(self)
        keep_window_on_top(window)
        window.title("Change sender email")
        window.transient(self.winfo_toplevel())
        window.grab_set()
        form = ttk.Frame(window, padding=12)
        form.pack(fill="both", expand=True)
        ttk.Label(form, text="Send this email from:").pack(anchor="w", pady=(0, 6))
        choices = [f"{address} — {name}" if name and name.casefold() != address.casefold() else address for address, name in accounts]
        current_idx = next((i for i, (address, _name) in enumerate(accounts) if address.casefold() == self._sender_email.casefold()), 0)
        selected = tk.StringVar(value=choices[current_idx])
        ttk.Combobox(form, textvariable=selected, values=choices, state="readonly", width=48).pack(fill="x")
        ttk.Label(
            form,
            text="Don't see an email you already added in Outlook itself? Outlook only reads its "
                 "account list once, when it starts -- fully quit Outlook (check the Windows system "
                 "tray, it can keep running there) and reopen it, then try again.",
            foreground="#555555", wraplength=420, justify="left",
        ).pack(anchor="w", pady=(6, 0))

        def add_account():
            try:
                add_window = tk.Toplevel(window)
                keep_window_on_top(add_window)
                add_window.title("Add / connect email account")
                add_window.transient(window)
                add_form = ttk.Frame(add_window, padding=12)
                add_form.pack(fill="both", expand=True)
                ttk.Label(
                    add_form,
                    text="Enter the email you want to add, then complete the sign-in in Outlook:",
                    wraplength=420,
                ).pack(anchor="w", pady=(0, 8))
                new_email = tk.StringVar()
                ttk.Entry(add_form, textvariable=new_email, width=48).pack(fill="x")
                ttk.Label(
                    add_form,
                    text="Outlook will ask for the password and verification. The app never reads or stores them.",
                    foreground="#555555", wraplength=420,
                ).pack(anchor="w", pady=(8, 10))

                def launch_setup():
                    if not new_email.get().strip() or "@" not in new_email.get():
                        messagebox.showwarning("Email required", "Enter a valid email address first.", parent=add_window)
                        return
                    open_outlook_account_setup()
                    add_window.destroy()
                    window.destroy()
                    messagebox.showinfo(
                        "Complete Outlook setup",
                        f"In the Outlook window, add and connect:\n\n{new_email.get().strip()}\n\n"
                        "When finished, open Change sender email again to select it.\n\n"
                        "Already added it in Outlook but still don't see it here? Outlook keeps "
                        "the account list loaded in memory from when it started, so just adding "
                        "the account isn't enough on its own -- fully quit Outlook (check the "
                        "Windows system tray, it can keep running there after the window is "
                        "closed) and reopen it, then try Change sender email again.",
                    )

                ttk.Button(add_form, text="Open Outlook account setup", command=launch_setup).pack(anchor="e")
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Outlook Account Setup", str(exc))

        ttk.Button(
            form, text="＋ Add / connect new email account", command=add_account,
        ).pack(anchor="w", pady=(8, 0))

        def save_sender():
            self._sender_email = accounts[choices.index(selected.get())][0]
            self._save_prefs(sender_email=self._sender_email)
            if label is not None:
                label.config(text=f"Sender: {self._sender_email}")
            window.destroy()

        ttk.Button(form, text="Save", command=save_sender).pack(anchor="e", pady=(10, 0))

    def _restore(self):
        for key, attr, source in [("biglietti_data_path", "data_path", "data_ordine"), ("biglietti_dispo_path", "dispo_path", "dispo_bagno")]:
            p = self._prefs.get(key) or (str(source_path(source)) if source_path(source) else "")
            if p and Path(p).is_file():
                setattr(self, attr, Path(p))
                getattr(self, f"{key}_label").config(text=p, foreground="#111827")
        template_value = self._prefs.get("biglietti_template_path", "")
        template_candidates = [
            Path(str(template_value)) if template_value else None,
            Path.home() / "OneDrive" / "Desktop" / "Biglietti.docx",
            Path.home() / "Desktop" / "Biglietti.docx",
            Path(__file__).resolve().parent / "Biglietti.docx",
        ]
        self.template_path = next(
            (candidate for candidate in template_candidates if candidate and candidate.is_file()),
            None,
        )
        if self.template_path:
            self.template_path_label.config(text=str(self.template_path), foreground="#111827")
        else:
            self.template_path_label.config(text="Not selected — choose Biglietti.docx", foreground="grey")
        for kind in ("elvy", "med", "el_kamal", "filato"):
            p = self._prefs.get(f"biglietti_{kind}_output_dir")
            if p and Path(p).is_dir():
                setattr(self, f"{kind}_output_dir", Path(p))
                getattr(self, f"{kind}_dir_label").config(text=p, foreground="#111827")
        shared = self._prefs.get("biglietti_shared_excel_path", "")
        if shared and Path(shared).is_file():
            self.shared_excel_path = Path(shared)
            self.shared_excel_label.config(text=shared, foreground="#111827")
        saved_templates = self._prefs.get("biglietti_email_templates", {})
        self._sender_email = str(self._prefs.get("sender_email", "") or "")
        for kind in ("elvy", "med", "el_kamal"):
            self.email_templates[kind] = EmailTemplate.from_dict(saved_templates.get(kind))
        try:
            articoli_cache = load_articoli_cache()
            if articoli_cache.get("source_path"):
                self.articoli_label.config(
                    text=f"Shared: {Path(articoli_cache['source_path']).name}",
                    foreground="#111827",
                )
        except Exception:
            pass
        cache = load_densita_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.densita_path = Path(cache["source_path"])
            self.densita_label.config(text=cache["source_path"], foreground="#111827")
        cache = load_magazino_cache()
        if cache.get("source_path"):
            self.vmm_label.config(text=f"• VMM22: {Path(cache['source_path']).name} (Yarn Warehouse)", foreground="#111827")
        price_cache = load_prezzi_cache()
        if price_cache.get("source_path"):
            self.prezzi_label.config(
                text=f"• Price: {Path(price_cache['source_path']).name} (Price List)",
                foreground="#111827",
            )

    def _pick_file(self, title, filetypes=None):
        return filedialog.askopenfilename(title=title, filetypes=filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])

    def _pick_data(self):
        p = self._pick_file("Select Order Data File")
        if p:
            self.data_path = Path(p)
            self.biglietti_data_path_label.config(text=p, foreground="#111827")
            save_source("data_ordine", self.data_path)
            self._save_prefs(biglietti_data_path=p)

    def _pick_dispo(self):
        p = self._pick_file("Select Dispo-Bagno File", filetypes=[("CSV/Excel", "*.csv;*.xlsx;*.xlsm"), ("All files", "*.*")])
        if p:
            self.dispo_path = Path(p)
            self.biglietti_dispo_path_label.config(text=p, foreground="#111827")
            save_source("dispo_bagno", self.dispo_path)
            self._save_prefs(biglietti_dispo_path=p)

    def set_template_path(self, path: str | Path) -> None:
        """Set and persist the Word ticket template selected by another page."""
        candidate = Path(path)
        if not candidate.is_file():
            return
        self.template_path = candidate
        self.template_path_label.config(text=str(candidate), foreground="#111827")
        self._save_prefs(biglietti_template_path=str(candidate))

    def _pick_template(self):
        path = filedialog.askopenfilename(
            title="Select Forma Biglietti (Word template)",
            filetypes=[("Word document", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.set_template_path(path)

    def _pick_output_dir(self, kind: str):
        p = filedialog.askdirectory(title=f"Select Output Folder ({kind.upper()})")
        if p:
            setattr(self, f"{kind}_output_dir", Path(p))
            getattr(self, f"{kind}_dir_label").config(text=p, foreground="#111827")
            self._save_prefs(**{f"biglietti_{kind}_output_dir": p})

    def _set_shared_excel(self, path: str | Path) -> None:
        candidate = Path(path)
        self.shared_excel_path = candidate
        self.shared_excel_label.config(text=str(candidate), foreground="#111827")
        self._save_prefs(biglietti_shared_excel_path=str(candidate))

    def _pick_shared_excel(self):
        p = filedialog.askopenfilename(
            title="Select shared Create Excel",
            filetypes=[("Excel workbook", "*.xlsx;*.xlsm"), ("All files", "*.*")],
        )
        if p:
            self._set_shared_excel(p)

    def _create_shared_excel(self):
        folder = filedialog.askdirectory(title="Select folder for shared Create Excel and Biglietti")
        if folder:
            self._set_shared_excel(Path(folder) / "Create Orders.xlsx")

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

    def _load_shared_orders_worker(self, path: Path, loading: tk.Toplevel):
        try:
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
        excel_columns = (
            "Dispo/Riga", "Cliente", "Articolo", "Titolo", "Formato", "Ordine", "Codice",
            "Colore", "Rocche", "KG", "M/C", "Partita Col", "Consegna", "Commento",
            "Bagno", "Partita GG", "Delivery Date", "Partita MED", "Cliente MED", "POLMON",
            "Color Tube", "VMM22", "Prezzo", "Densita` (360-390)",
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

        ttk.Button(search_bar, text="SAVE", width=10, command=save_pg_x).grid(row=0, column=5, sticky="e", padx=(6, 0))

        def edit_pg_x_row(event):
            if current_sheet["name"] not in ("PG-X", "Orders"):
                return
            iid = tree.identify_row(event.y)
            if not iid or sheet_by_iid.get(iid) != current_sheet["name"]:
                return
            record = record_by_iid[iid]
            editor = tk.Toplevel(window)
            editor.title(f"Edit {current_sheet['name']} Color")
            editor.geometry("470x300")
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
                ("Partita GG", str(record.raw_batch or "")),
                ("Bagno", str(record.bagno or "")),
                ("Rocche", "" if record.quantity_cones is None else str(record.quantity_cones)),
                ("M/C", str(record.machine or "")),
            )
            entries = {}
            machine_choices = tuple(str(machine) for machine in range(3, 13))
            for row_idx, (label, value) in enumerate(fields, start=1):
                ttk.Label(form, text=f"{label}:").grid(row=row_idx, column=0, sticky="w", pady=5)
                if label in {"Articolo", "Partita GG"}:
                    entry = ttk.Combobox(form, width=30, state="normal")
                elif label == "M/C":
                    entry = ttk.Combobox(form, width=30, values=machine_choices, state="normal")
                else:
                    entry = ttk.Entry(form, width=32)
                entry.grid(row=row_idx, column=1, sticky="ew", pady=5)
                entry.insert(0, value)
                entries[label] = entry

            status_row = len(fields) + 1
            expected_raw = str(record.article or "").strip().upper()
            expected_raw = "G" + expected_raw[1:] if expected_raw.startswith("C") else expected_raw
            availability_var = tk.StringVar(value=f"Expected raw article: {expected_raw} — loading Magazino...")
            availability_label = ttk.Label(form, textvariable=availability_var, foreground="#666666", wraplength=340)
            availability_label.grid(row=status_row, column=0, columnspan=2, sticky="w", pady=(2, 4))
            stock_by_partita = {}
            partita_choices_by_article = {}
            articolo_choices = set()
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
                expected = "G" + article[1:] if article.startswith("C") else article
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
                expected = "G" + article[1:] if article.startswith("C") else article
                choices = sorted(
                    partita_choices_by_article.get(expected, set()),
                    key=lambda value: (normal_partita(value).casefold(), str(value)),
                )
                entries["Partita GG"]["values"] = choices

            def refresh_articolo_choices(*_args):
                entries["Articolo"]["values"] = sorted(articolo_choices, key=str.casefold)

            def load_stock():
                try:
                    _codes, _density, _vmm, _prices, summary = self._load_common_sources()
                    if summary is not None:
                        for row in summary.itertuples(index=False):
                            article = str(getattr(row, "articolo", "") or "").strip().upper()
                            partita = normal_partita(getattr(row, "partita", ""))
                            available = float(getattr(row, "mag_rocche", 0) or 0)
                            stock_by_partita.setdefault(partita, []).append((available, article))
                            if article:
                                articolo_choices.add("C" + article[1:] if article.startswith("G") else article)
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
            threading.Thread(target=load_stock, daemon=True).start()

            def save_from_editor():
                values = {label: entry.get().strip() for label, entry in entries.items()}
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

        def toggle(_event=None):
            iid = tree.identify_row(_event.y) if _event is not None else ""
            column = tree.identify_column(_event.x) if _event is not None else ""
            if not iid or column != "#1":
                return
            selected[iid] = not selected[iid]
            values = list(tree.item(iid, "values"))
            values[0] = "☑" if selected[iid] else "☐"
            tree.item(iid, values=values)
            tree.selection_set(iid)

        tree.bind("<ButtonRelease-1>", toggle)
        actions = ttk.Frame(window, padding=(10, 0, 10, 10))
        actions.grid(row=2, column=0, sticky="ew")
        ttk.Label(actions, text="Click Print to select or unselect each color.").pack(side="left")
        def pg_x_action(action):
            if current_sheet["name"] != "PG-X":
                return messagebox.showwarning("PG-X Only", "This action is available only on the PG-X page.", parent=window)
            selection = tree.selection()
            if not selection:
                return messagebox.showwarning("No Row Selected", "Click a PG-X row first.", parent=window)
            iid = selection[0]
            partita_col = str(record_by_iid[iid].colored_batch)
            if action == "delete" and not messagebox.askyesno(
                "Delete PG-X Row", f"Delete Partita Col {partita_col} from PG-X?", parent=window
            ):
                return
            self._pg_x_row_action_in_background(
                window, action, partita_col, datasets, refresh_records, rebuild,
            )

        ttk.Button(actions, text="Send to Orders", command=lambda: pg_x_action("move")).pack(side="left", padx=(16, 4))
        ttk.Button(actions, text="Delete", command=lambda: pg_x_action("delete")).pack(side="left", padx=4)
        ttk.Button(
            actions, text="Delete Shipped Colors", style="Danger.TButton",
            command=lambda: self._delete_shipped_colors(
                window, current_sheet["name"], record_by_iid, sheet_by_iid,
                datasets, refresh_records, rebuild,
            ),
        ).pack(side="left", padx=(14, 4))
        ttk.Button(
            actions, text="🖨  Print Selected Biglietti",
            command=lambda: self._print_selected_shared(
                window, selected, record_by_iid,
                move_assigned_pg_x=current_sheet["name"] == "PG-X",
            ),
        ).pack(side="right")
        ttk.Button(actions, text="Close", command=window.destroy).pack(side="right", padx=(0, 8))

    @staticmethod
    def _toggle_child_maximize(window: tk.Toplevel) -> None:
        try:
            window.state("normal" if window.state() == "zoomed" else "zoomed")
        except tk.TclError:
            pass

    @staticmethod
    def _start_child_drag(window: tk.Toplevel, event) -> None:
        window._drag_origin = (event.x_root, event.y_root, window.winfo_x(), window.winfo_y())

    @staticmethod
    def _drag_child(window: tk.Toplevel, event) -> None:
        origin = getattr(window, "_drag_origin", None)
        if origin is None:
            return
        start_x, start_y, window_x, window_y = origin
        window.geometry(f"+{window_x + event.x_root - start_x}+{window_y + event.y_root - start_y}")

    @staticmethod
    def _restore_child_window(window: tk.Toplevel) -> None:
        try:
            window.state("normal")
            window.deiconify()
        except tk.TclError:
            pass

    def _save_pg_x_details_in_background(
        self, window, partita_col, values, datasets, refresh_records, rebuild,
        partita_gg_var, sheet_name="PG-X", allow_article_mismatch=False,
    ):
        """Persist the PG-X editor fields, then apply the normal Partita GG flow."""
        if getattr(self, "_pg_x_saving", False):
            return
        self._pg_x_saving = True

        def as_number(text):
            value = str(text or "").strip()
            if not value:
                return ""
            try:
                number = float(value.replace(",", "."))
                return int(number) if number.is_integer() else number
            except ValueError:
                return value

        updates = {
            "Articolo": values.get("Articolo", ""),
            "Partita GG": values.get("Partita GG", ""),
            "Bagno": values.get("Bagno", ""),
            "Rocche": as_number(values.get("Rocche", "")),
            "M/C": values.get("M/C", ""),
        }

        def worker():
            try:
                partita_gg = str(values.get("Partita GG", "")).strip()
                if sheet_name == "Orders":
                    magazino_summary = None
                    if partita_gg:
                        _codes, _densita_map, _vmm_ratio_map, _prices, magazino_summary = self._load_common_sources()
                    result = update_order_row(
                        self.shared_excel_path, partita_col, updates,
                        magazino_summary=magazino_summary,
                        allow_article_mismatch=allow_article_mismatch,
                    )
                    if result["moved_to_pgx"]:
                        message = f"Moved {result['updated']} row(s) to PG-X."
                    else:
                        message = f"Updated {result['updated']} row(s) in Orders."
                else:
                    update_pg_x_row(self.shared_excel_path, partita_col, updates)
                    if partita_gg:
                        _codes, densita_map, vmm_ratio_map, _prices, magazino_summary = self._load_common_sources()
                        result = save_pg_x_partita(
                            self.shared_excel_path, partita_col, partita_gg,
                            densita_map=densita_map, vmm_ratio_map=vmm_ratio_map,
                            magazino_summary=magazino_summary,
                            allow_article_mismatch=allow_article_mismatch,
                        )
                        message = f"Moved {result['updated']} row(s) to Orders. Available raw yarn: {result['available']:g} rocche."
                    else:
                        result = {"updated": 1}
                        message = "PG-X color updated."
                new_datasets = {
                    "Orders": load_create_excel_records(self.shared_excel_path, sheet_name="Orders"),
                    "PG-X": load_create_excel_records(self.shared_excel_path, sheet_name="PG-X"),
                }
                self.after(0, lambda: finish(new_datasets, message))
            except Exception as exc:
                self.after(0, lambda exc=exc: fail(exc))

        def finish(new_datasets, message):
            self._pg_x_saving = False
            refresh_records(new_datasets)
            partita_gg_var.set("")
            rebuild()
            messagebox.showinfo("PG-X Saved", message, parent=window)

        def fail(exc):
            self._pg_x_saving = False
            messagebox.showerror("PG-X Save Error", str(exc), parent=window)

        threading.Thread(target=worker, daemon=True).start()

    def _save_pg_x_in_background(self, window, partita_col, partita_gg, datasets, refresh_records, rebuild, partita_gg_var, allow_article_mismatch=False):
        if getattr(self, "_pg_x_saving", False):
            return
        self._pg_x_saving = True

        def worker():
            try:
                _codes, densita_map, vmm_ratio_map, _prices, magazino_summary = self._load_common_sources()
                result = save_pg_x_partita(
                    self.shared_excel_path, partita_col, partita_gg,
                    densita_map=densita_map, vmm_ratio_map=vmm_ratio_map,
                    magazino_summary=magazino_summary,
                    allow_article_mismatch=allow_article_mismatch,
                )
                new_datasets = {
                    "Orders": load_create_excel_records(self.shared_excel_path, sheet_name="Orders"),
                    "PG-X": load_create_excel_records(self.shared_excel_path, sheet_name="PG-X"),
                }
                self.after(0, lambda: finish(result, new_datasets))
            except Exception as exc:
                self.after(0, lambda exc=exc: fail(exc))

        def finish(result, new_datasets):
            self._pg_x_saving = False
            refresh_records(new_datasets)
            partita_gg_var.set("")
            rebuild()
            messagebox.showinfo(
                "PG-X Saved",
                f"Updated {result['updated']} row(s). Available raw yarn: {result['available']:g} rocche.",
                parent=window,
            )

        def fail(exc):
            self._pg_x_saving = False
            messagebox.showerror("PG-X Save Error", str(exc), parent=window)

        threading.Thread(target=worker, daemon=True).start()

    def _pg_x_row_action_in_background(self, window, action, partita_col, datasets, refresh_records, rebuild):
        if getattr(self, "_pg_x_action_running", False):
            return
        self._pg_x_action_running = True

        def worker():
            try:
                count = move_pg_x_to_orders(self.shared_excel_path, partita_col) if action == "move" else delete_pg_x_row(self.shared_excel_path, partita_col)
                new_datasets = {
                    "Orders": load_create_excel_records(self.shared_excel_path, sheet_name="Orders"),
                    "PG-X": load_create_excel_records(self.shared_excel_path, sheet_name="PG-X"),
                }
                self.after(0, lambda: finish(count, new_datasets))
            except Exception as exc:
                self.after(0, lambda exc=exc: fail(exc))

        def finish(count, new_datasets):
            self._pg_x_action_running = False
            refresh_records(new_datasets)
            rebuild()
            messagebox.showinfo(
                "PG-X",
                f"{count} row(s) {'sent to Orders' if action == 'move' else 'deleted'}.",
                parent=window,
            )

        def fail(exc):
            self._pg_x_action_running = False
            messagebox.showerror("PG-X Error", str(exc), parent=window)

        threading.Thread(target=worker, daemon=True).start()

    def _delete_shipped_colors(
        self, window, sheet_name, record_by_iid, sheet_by_iid,
        datasets, refresh_records, rebuild,
    ):
        if getattr(self, "_delete_shipped_running", False):
            return
        uscita_path = source_path("uscita", existing_only=True)
        if not uscita_path:
            return messagebox.showwarning(
                "Uscita File Required",
                "Upload the Uscita file in Situazione first, then try again.",
                parent=window,
            )
        self._delete_shipped_running = True

        def normalize(value):
            text = str(value or "").strip()
            try:
                number = float(text.replace(",", "."))
                return str(int(number)) if number.is_integer() else str(number)
            except ValueError:
                return text.casefold()

        def worker():
            try:
                from parsers.situazione_loaders import load_uscita
                uscita_df, errors = load_uscita(str(uscita_path))
                if errors or uscita_df is None or uscita_df.empty:
                    detail = "; ".join(errors) if errors else "The Uscita file has no valid shipped rows."
                    raise ValueError(detail)
                shipped = {normalize(value) for value in uscita_df["partita"].tolist() if normalize(value)}
                page_records = [
                    record for iid, record in record_by_iid.items()
                    if sheet_by_iid.get(iid) == sheet_name
                ]
                count = sum(1 for record in page_records if normalize(record.colored_batch) in shipped)
                self.after(0, lambda: confirm_and_delete(count, shipped))
            except Exception as exc:
                self.after(0, lambda exc=exc: shipped_delete_failed(exc, window))

        def confirm_and_delete(count, shipped):
            if not count:
                self._delete_shipped_running = False
                return messagebox.showinfo(
                    "Delete Shipped Colors",
                    f"No shipped colors were found in {sheet_name}.",
                    parent=window,
                )
            if not messagebox.askyesno(
                "Delete Shipped Colors",
                f"{count} color(s) in {sheet_name} have already been shipped.\n\nDelete them now?",
                parent=window,
            ):
                self._delete_shipped_running = False
                return

            def delete_worker():
                try:
                    deleted = delete_shipped_shared_rows(self.shared_excel_path, shipped, sheet_name)
                    new_datasets = {
                        "Orders": load_create_excel_records(self.shared_excel_path, sheet_name="Orders"),
                        "PG-X": load_create_excel_records(self.shared_excel_path, sheet_name="PG-X"),
                    }
                    self.after(0, lambda: delete_finished(deleted, new_datasets))
                except Exception as exc:
                    self.after(0, lambda exc=exc: shipped_delete_failed(exc, window))

            threading.Thread(target=delete_worker, daemon=True).start()

        def delete_finished(deleted, new_datasets):
            self._delete_shipped_running = False
            refresh_records(new_datasets)
            rebuild()
            messagebox.showinfo(
                "Delete Shipped Colors",
                f"Deleted {deleted} shipped color(s) from {sheet_name}.",
                parent=window,
            )

        def shipped_delete_failed(exc, parent):
            self._delete_shipped_running = False
            messagebox.showerror("Delete Shipped Colors", str(exc), parent=parent)

        threading.Thread(target=worker, daemon=True).start()

    def _print_selected_shared(
        self, window, selected: dict[str, bool], record_by_iid: dict[str, object],
        move_assigned_pg_x: bool = False,
    ):
        if not self.template_path or not self.template_path.is_file():
            return messagebox.showwarning("Missing Template", "Select the Biglietti.docx template first.", parent=window)
        records = [record_by_iid[iid] for iid, is_selected in selected.items() if is_selected]
        if not records:
            return messagebox.showwarning("No Orders Selected", "Select at least one color to print.", parent=window)
        # Keep one stable output file: every print replaces the previous
        # document with only the records selected in the current view.
        destination = self.shared_excel_path.parent / "Biglietti_Selected.docx"
        pg_x_partita_cols = []
        if move_assigned_pg_x:
            pg_x_partita_cols = list(dict.fromkeys(
                str(record.colored_batch)
                for record in records
                if str(record.raw_batch or "").strip().upper().replace(" ", "") not in {"", "X", "PG-X", "PGX"}
            ))
        window.destroy()
        self.convert_btn.config(state="disabled")
        self._set_status("Creating selected Biglietti from shared Excel in progress...")
        threading.Thread(
            target=self._worker_shared_biglietti,
            args=(records, destination, pg_x_partita_cols),
            daemon=True,
        ).start()

    def _worker_shared_biglietti(self, records, destination: Path, pg_x_partita_cols=None):
        try:
            export_word(destination, self.template_path, records, stem="Selected Orders")
            moved = 0
            for partita_col in pg_x_partita_cols or []:
                moved += move_pg_x_to_orders(self.shared_excel_path, partita_col)
            self._set_status(f"Created {len(records)} selected Biglietti.")
            moved_text = f"\nMoved {moved} PG-X row(s) to Orders." if moved else ""
            self.after(0, lambda: messagebox.showinfo("Biglietti", f"Created:\n{destination}{moved_text}"))
        except Exception as exc:
            self._logger.exception("Shared Excel Biglietti failed")
            self._set_status(f"Error: {exc}")
            self.after(0, lambda exc=exc: messagebox.showerror("Biglietti Error", str(exc)))
        finally:
            self.after(0, lambda: self.convert_btn.config(state="normal"))

    def _pick_articoli(self):
        p = self._pick_file("Select Articoli.xlsx")
        if not p:
            return
        import exporters.biglietti_exporter as biglietti_exporter
        try:
            marca_map, marca_errors = biglietti_exporter.load_articoli_marca_map(Path(p))
        except Exception as exc:
            self._logger.exception("Articoli (Marca) read failed")
            marca_map, marca_errors = {}, [str(exc)]
        try:
            import utility.situazione_db as situazione_db
            import parsers.situazione_loaders as situazione_loaders
            titolo_df, titolo_errors = situazione_loaders.load_codes(p)
        except Exception as exc:
            self._logger.exception("Articoli (TITOLO) read failed")
            titolo_df, titolo_errors = None, [str(exc)]
        if not marca_map and not titolo_df:
            errors = marca_errors + (titolo_errors or [])
            messagebox.showerror("Articles Error", "; ".join(errors) if errors else "Invalid file: no recognizable Articolo Filato column found.")
            return
        self.articoli_path = Path(p)
        save_source("articoli", self.articoli_path)
        parts = []
        if marca_map:
            save_articoli_cache(p)
            parts.append(f"{len(marca_map)} articles (Brand/Marca)")
        if titolo_df is not None and not titolo_df.empty:
            situazione_db.save_codes(titolo_df)
            parts.append(f"{len(titolo_df)} rows shared with Situation")
        self.articoli_label.config(text=" + ".join(parts) + f" loaded from {Path(p).name}", foreground="#111827")
        if self._on_shared_cache_changed:
            self._on_shared_cache_changed()

    def _pick_densita(self):
        p = self._pick_file("Select Densita' Query.xlsx")
        if p:
            self.densita_path = Path(p)
            self.densita_label.config(text=p, foreground="#111827")
            save_densita_cache(p)
            if self._on_shared_cache_changed:
                self._on_shared_cache_changed()

    def _set_status(self, msg: str):
        color = "#C62828" if "Error" in msg else ("#1565C0" if "progress" in msg or "Converting" in msg else "#2E7D32")
        prefix = "✖ " if "Error" in msg else ("⏳ " if "progress" in msg or "Converting" in msg else "✔ ")
        self.after(0, lambda: self.status.config(text=f"{prefix}{msg}", foreground=color))

    def _load_common_sources(self):
        codes_map = load_articoli_marca_lookup() or load_articoli_titolo_map()
        densita_map = {}
        if self.densita_path and self.densita_path.is_file():
            densita_map, _errors = load_densita_query(self.densita_path)
        vmm_ratio_map = {}
        magazino_path = load_magazino_cache().get("source_path")
        if magazino_path and Path(magazino_path).is_file():
            vmm_ratio_map, _errors = load_vmm22_ratio_from_magazino(Path(magazino_path))
        magazino_summary = None
        if magazino_path and Path(magazino_path).is_file():
            from calculate import magazino as magazino_logic
            magazino_df, _errors = magazino_logic.load_magazino(Path(magazino_path), articolo_prefix=None)
            magazino_summary = magazino_logic.summarize_by_partita(magazino_df)
        price_lookup, _price_source = load_prezzo_lookup()
        return codes_map, densita_map, vmm_ratio_map, price_lookup, magazino_summary

    def _run_convert(self):
        if not self.data_path or not self.data_path.is_file():
            return messagebox.showwarning("Missing Input", "Please select the Order Data file first.")
        try:
            order_format = detect_order_format(self.data_path)
        except Exception as exc:
            return messagebox.showerror("File Error", f"Could not read Order Data file:\n{exc}")
        if order_format == "EL_KAMAL":
            if not self.el_kamal_output_dir or not Path(self.el_kamal_output_dir).is_dir():
                self._pick_output_dir("el_kamal")
                if not self.el_kamal_output_dir or not Path(self.el_kamal_output_dir).is_dir():
                    return
        else:
            try:
                records, _raw = load_order(self.data_path, self.dispo_path)
                has_elvy = any(r.customer_code == "3009" for r in records)
                has_med = any(r.customer_code == "3004" for r in records)
            except Exception as exc:
                return messagebox.showerror("Read Error", f"Could not parse order records:\n{exc}")
            if not has_elvy and not has_med:
                return messagebox.showwarning("No Customer Rows", "No customer rows found for 3004 or 3009 in the selected Data Ordine.")
            if has_elvy and (not self.elvy_output_dir or not Path(self.elvy_output_dir).is_dir()):
                self._pick_output_dir("elvy")
                if not self.elvy_output_dir or not Path(self.elvy_output_dir).is_dir(): return
            if has_med and (not self.med_output_dir or not Path(self.med_output_dir).is_dir()):
                self._pick_output_dir("med")
                if not self.med_output_dir or not Path(self.med_output_dir).is_dir(): return
            if self.filato_enabled.get() and (not self.filato_output_dir or not Path(self.filato_output_dir).is_dir()):
                self._pick_output_dir("filato")
                if not self.filato_output_dir or not Path(self.filato_output_dir).is_dir(): return
            if not self.shared_excel_path:
                self._create_shared_excel()
                if not self.shared_excel_path:
                    return
        self.convert_btn.config(state="disabled")
        self._set_status("Converting orders and generating tickets in progress...")
        threading.Thread(target=self._worker_convert, daemon=True).start()

    def _worker_convert(self):
        created_items: list[str] = []
        self._last_client_files = {"elvy": [], "med": [], "el_kamal": []}
        try:
            order_format = detect_order_format(self.data_path)
            codes_map, densita_map, vmm_ratio_map, price_lookup, magazino_summary = self._load_common_sources()
            template = self.template_path
            if not template or not template.is_file():
                raise ValueError("Forma Biglietti not found. Select the Biglietti.docx template first.")
            if order_format == "EL_KAMAL":
                records, raw = load_el_kamal_order(self.data_path, self.dispo_path)
                stem = build_el_kamal_stem(records); out_dir = Path(self.el_kamal_output_dir)
                xlsx = out_dir / f"{stem}_EL_KAMAL.xlsx"; docx = out_dir / f"{stem}_EL_KAMAL_Biglietti.docx"
                enrich_records(records, "EL_KAMAL", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                export_workbook(xlsx, records, raw, include_filato=True, stem=stem, customer="EL_KAMAL", magazino_summary=magazino_summary)
                export_word(docx, template, records, stem=stem)
                created_items.append(f"• EL KAMAL: {len(records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")
                self._last_client_files["el_kamal"] = [xlsx, docx]
            else:
                records, raw = load_order(self.data_path, self.dispo_path)
                elvy_records = [r for r in records if r.customer_code == "3009"]
                if elvy_records:
                    out_dir = Path(self.elvy_output_dir); stem = build_output_stem(elvy_records, "ELVY")
                    xlsx = out_dir / f"{stem}_ELVY.xlsx"; docx = out_dir / f"{stem}_ELVY_Biglietti.docx"
                    enrich_records(elvy_records, "ELVY", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                    export_workbook(xlsx, elvy_records, raw, include_filato=True, stem=stem, customer="ELVY", magazino_summary=magazino_summary)
                    shared_result = append_create_excel(self.shared_excel_path, elvy_records, "ELVY")
                    elvy_ready, elvy_pgx = _split_pg_x_records(elvy_records)
                    files = [xlsx]
                    if elvy_ready:
                        export_word(docx, template, elvy_ready, stem=stem)
                        files.append(docx)
                    if elvy_pgx:
                        pgx_docx = docx.with_name(f"{docx.stem}-pg-x{docx.suffix}")
                        export_word(pgx_docx, template, elvy_pgx, stem=f"{stem}-PG-X")
                        files.append(pgx_docx)
                    word_text = ", ".join(path.name for path in files[1:]) or "no Word tickets"
                    created_items.append(f"• ELVY: {len(elvy_records)} tickets ({word_text})\n   ↳ Saved to: {xlsx}\n   ↳ Shared Excel: +{shared_result['added']} rows, {shared_result['skipped']} duplicate Partita Col skipped")
                    self._last_client_files["elvy"] = files
                med_records = [r for r in records if r.customer_code == "3004"]
                if med_records:
                    out_dir = Path(self.med_output_dir); stem = build_output_stem(med_records, "MED")
                    xlsx = out_dir / f"{stem}_MED.xlsx"; docx = out_dir / f"{stem}_MED_Biglietti.docx"
                    enrich_records(med_records, "MED", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                    export_workbook(xlsx, med_records, raw, include_filato=True, stem=stem, customer="MED", magazino_summary=magazino_summary)
                    shared_result = append_create_excel(self.shared_excel_path, med_records, "MED")
                    med_ready, med_pgx = _split_pg_x_records(med_records)
                    files = [xlsx]
                    if med_ready:
                        export_word(docx, template, med_ready, stem=stem)
                        files.append(docx)
                    if med_pgx:
                        pgx_docx = docx.with_name(f"{docx.stem}-pg-x{docx.suffix}")
                        export_word(pgx_docx, template, med_pgx, stem=f"{stem}-PG-X")
                        files.append(pgx_docx)
                    word_text = ", ".join(path.name for path in files[1:]) or "no Word tickets"
                    created_items.append(f"• MED: {len(med_records)} tickets ({word_text})\n   ↳ Saved to: {xlsx}\n   ↳ Shared Excel: +{shared_result['added']} rows, {shared_result['skipped']} duplicate Partita Col skipped")
                    self._last_client_files["med"] = files
                if self.filato_enabled.get() and self.filato_output_dir and raw:
                    out_dir = Path(self.filato_output_dir); out_dir.mkdir(parents=True, exist_ok=True)
                    filato_file = out_dir / "Filato x Tinturia.xlsx"
                    matches = [
                        RawYarnMatch(
                            articolo=r["Articolo"], titolo=r["Titolo"], partita=r["Partita"],
                            rocce=r["Rocche"], peso=r["Peso"], label=r["تحضير خام"],
                        )
                        for r in _filato_rows(records, raw, magazino_summary)
                    ]
                    n = export_filato_full(filato_file, matches)
                    created_items.append(f"• Raw Yarn (Filato): {filato_file.name} ({n} rows)\n   ↳ Saved to: {filato_file}")
            summary_text = "\n\n".join(created_items)
            self._set_status("Completed — All order workbooks and dyeing tickets generated successfully.")
            self.after(0, lambda: messagebox.showinfo("Conversion Complete", f"Conversion completed successfully!\n\n{summary_text}"))
        except Exception as exc:
            self._logger.exception("Biglietti conversion failed")
            self._set_status(f"Error: {exc}")
            self.after(0, lambda exc=exc: messagebox.showerror("Conversion Error", str(exc)))
        finally:
            self.after(0, lambda: self.convert_btn.config(state="normal"))
