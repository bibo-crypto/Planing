"""Tkinter page for extracting Biglietti, customer workbook and Filato."""

from __future__ import annotations

import threading
import tkinter as tk
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
        self.columnconfigure(0, weight=1)
        style = ttk.Style(self)
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=4, pady=(0, 10))
        ttk.Label(header_frame, text="Order Extraction & Dyeing Tickets", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(header_frame, text="Single Order Data + Dispo-Bagno for all clients (ELVY, MED, EL KAMAL) — format is automatically recognized.", foreground="#666666", wraplength=950).pack(anchor="w", pady=(2, 0))

        input_box = ttk.LabelFrame(self, text=" 📁 Main Input Files (Required) ", padding=10)
        input_box.pack(fill="x", padx=4, pady=(0, 10))
        input_box.columnconfigure(1, weight=1)
        ttk.Button(input_box, text="📂  Select Order Data", command=self._pick_data, width=24).grid(row=0, column=0, sticky="w", pady=4)
        self.biglietti_data_path_label = ttk.Label(input_box, text="No file selected", foreground="grey", anchor="w")
        self.biglietti_data_path_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)
        ttk.Button(input_box, text="📂  Select Dispo-Bagno", command=self._pick_dispo, width=24).grid(row=1, column=0, sticky="w", pady=4)
        self.biglietti_dispo_path_label = ttk.Label(input_box, text="No file selected (optional if embedded)", foreground="grey", anchor="w")
        self.biglietti_dispo_path_label.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)

        template_box = ttk.LabelFrame(self, text=" 📝 Select Forma Biglietti ", padding=10)
        template_box.pack(fill="x", padx=4, pady=(0, 10))
        template_box.columnconfigure(1, weight=1)
        ttk.Button(template_box, text="📄  Select Forma Biglietti", command=self._pick_template, width=24).grid(row=0, column=0, sticky="w", pady=4)
        self.template_path_label = ttk.Label(template_box, text="Searching for Biglietti.docx...", foreground="grey", anchor="w")
        self.template_path_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)

        opt_box = ttk.LabelFrame(self, text=" ⚙️ Optional Data Sources (Shared across all clients) ", padding=10)
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

        export_box = ttk.LabelFrame(self, text=" 📂 Output Destinations & Settings ", padding=10)
        export_box.pack(fill="x", padx=4, pady=(0, 10))
        export_box.columnconfigure(1, weight=1)
        self._build_folder_row(export_box, 0, "elvy", "📁  ELVY Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 1, "med", "📁  MED Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 2, "el_kamal", "📁  EL KAMAL Output Folder", show_email_buttons=True)
        self._build_folder_row(export_box, 3, "filato", "📁  Filato Output Folder", with_checkbox=True)

        action_box = ttk.Frame(self)
        action_box.pack(fill="x", padx=4, pady=(8, 4))
        action_box.columnconfigure(0, weight=1)
        self.convert_btn = ttk.Button(action_box, text="⚡  Convert & Generate Tickets (Excel + Word)", command=self._run_convert, width=42, style="Bold.TButton")
        self.convert_btn.pack(side="top", anchor="w", pady=(0, 8))
        self.status = ttk.Label(action_box, text="● Ready", font=("Segoe UI", 9, "bold"), foreground="#2E7D32", anchor="w", wraplength=950)
        self.status.pack(fill="x")

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
            email_frame = ttk.Frame(parent)
            email_frame.grid(row=row_idx, column=2, sticky="e", padx=(10, 0), pady=4)
            ttk.Button(
                email_frame, text="✉ Template", width=11,
                command=lambda: self._open_email_template_editor(kind),
            ).pack(side="left", padx=(0, 4))
            ttk.Button(
                email_frame, text="📧 Prepare Email", width=16,
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
                        "When finished, open Change sender email again to select it.",
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
                    export_word(docx, template, elvy_records, stem=stem)
                    created_items.append(f"• ELVY: {len(elvy_records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")
                    self._last_client_files["elvy"] = [xlsx, docx]
                med_records = [r for r in records if r.customer_code == "3004"]
                if med_records:
                    out_dir = Path(self.med_output_dir); stem = build_output_stem(med_records, "MED")
                    xlsx = out_dir / f"{stem}_MED.xlsx"; docx = out_dir / f"{stem}_MED_Biglietti.docx"
                    enrich_records(med_records, "MED", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                    export_workbook(xlsx, med_records, raw, include_filato=True, stem=stem, customer="MED", magazino_summary=magazino_summary)
                    export_word(docx, template, med_records, stem=stem)
                    created_items.append(f"• MED: {len(med_records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")
                    self._last_client_files["med"] = [xlsx, docx]
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
