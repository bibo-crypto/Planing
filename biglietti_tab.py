"""Tkinter page for extracting Biglietti, customer workbook and Filato."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from biglietti_exporter import (
    build_el_kamal_stem,
    build_output_stem,
    detect_order_format,
    enrich_records,
    export_filato_workbook,
    export_word,
    export_workbook,
    load_articoli_marca_lookup,
    load_articoli_titolo_map,
    load_densita_query,
    load_el_kamal_order,
    load_order,
    load_prezzo_lookup,
    load_vmm22_ratio_from_magazino,
)
from articoli_cache import save_articoli_cache
from densita_cache import load_densita_cache, save_densita_cache
from magazino_cache import load_magazino_cache


class BigliettiTab(ttk.Frame):
    def __init__(self, parent, prefs: dict, save_prefs, logger):
        super().__init__(parent, padding=12)
        self._prefs = prefs
        self._save_prefs = save_prefs
        self._logger = logger
        self.data_path: Path | None = None
        self.dispo_path: Path | None = None
        self.elvy_output_dir: Path | None = None
        self.med_output_dir: Path | None = None
        self.el_kamal_output_dir: Path | None = None
        self.filato_output_dir: Path | None = None
        self.articoli_path: Path | None = None
        self.densita_path: Path | None = None
        self.filato_enabled = tk.BooleanVar(value=False)
        self._build()
        self._restore()

    def _build(self):
        self.columnconfigure(0, weight=1)

        style = ttk.Style(self)
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))

        # ── Header banner ──
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=4, pady=(0, 10))
        ttk.Label(header_frame, text="Order Extraction & Dyeing Tickets", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            header_frame,
            text="Single Order Data + Dispo-Bagno for all clients (ELVY, MED, EL KAMAL) — format is automatically recognized.",
            foreground="#666666",
            wraplength=950,
        ).pack(anchor="w", pady=(2, 0))

        # ── 1. Main Input Files (Required) ──
        input_box = ttk.LabelFrame(self, text=" 📁 Main Input Files (Required) ", padding=10)
        input_box.pack(fill="x", padx=4, pady=(0, 10))
        input_box.columnconfigure(1, weight=1)

        # Data Ordine
        ttk.Button(input_box, text="📂  Select Order Data", command=self._pick_data, width=24).grid(row=0, column=0, sticky="w", pady=4)
        self.biglietti_data_path_label = ttk.Label(input_box, text="No file selected", foreground="grey", anchor="w")
        self.biglietti_data_path_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)

        # Dispo-Bagno
        ttk.Button(input_box, text="📂  Select Dispo-Bagno", command=self._pick_dispo, width=24).grid(row=1, column=0, sticky="w", pady=4)
        self.biglietti_dispo_path_label = ttk.Label(input_box, text="No file selected (optional if embedded)", foreground="grey", anchor="w")
        self.biglietti_dispo_path_label.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)

        # ── 2. Optional Reference Sources (Shared) ──
        opt_box = ttk.LabelFrame(self, text=" ⚙️ Optional Data Sources (Shared across all clients) ", padding=10)
        opt_box.pack(fill="x", padx=4, pady=(0, 10))
        opt_box.columnconfigure(1, weight=1)

        # Articoli
        self.articoli_btn = ttk.Button(opt_box, text="📂  Articles (Titolo)", command=self._pick_articoli, width=24)
        self.articoli_btn.grid(row=0, column=0, sticky="w", pady=4)
        self.articoli_label = ttk.Label(opt_box, text="Shared with Situation tab", foreground="grey", anchor="w")
        self.articoli_label.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)

        # Densita
        self.densita_btn = ttk.Button(opt_box, text="📂  Density Query", command=self._pick_densita, width=24)
        self.densita_btn.grid(row=1, column=0, sticky="w", pady=4)
        self.densita_label = ttk.Label(opt_box, text="Not selected — KG remains raw yarn weight", foreground="grey", anchor="w")
        self.densita_label.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)

        # Shared info row
        info_frame = ttk.Frame(opt_box)
        info_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        info_frame.columnconfigure(0, weight=1)
        info_frame.columnconfigure(1, weight=1)

        self.vmm_label = ttk.Label(info_frame, text="• VMM22: Uses Yarn Warehouse (Magazino Filato) loaded in app", foreground="#555555", font=("Segoe UI", 8))
        self.vmm_label.grid(row=0, column=0, sticky="w")

        self.prezzi_label = ttk.Label(info_frame, text="• Price: Uses Price List (Listini) loaded in Prices / Situation", foreground="#555555", font=("Segoe UI", 8))
        self.prezzi_label.grid(row=0, column=1, sticky="w")

        # ── 3. Output Folders & Export Settings ──
        export_box = ttk.LabelFrame(self, text=" 📂 Output Destinations & Settings ", padding=10)
        export_box.pack(fill="x", padx=4, pady=(0, 10))
        export_box.columnconfigure(1, weight=1)

        self._build_folder_row(export_box, 0, "elvy", "📁  ELVY Output Folder")
        self._build_folder_row(export_box, 1, "med", "📁  MED Output Folder")
        self._build_folder_row(export_box, 2, "el_kamal", "📁  EL KAMAL Output Folder")
        self._build_folder_row(export_box, 3, "filato", "📁  Filato Output Folder", with_checkbox=True)

        # ── 4. Convert Action & Status Footer ──
        action_box = ttk.Frame(self)
        action_box.pack(fill="x", padx=4, pady=(8, 4))
        action_box.columnconfigure(0, weight=1)

        self.convert_btn = ttk.Button(
            action_box,
            text="⚡  Convert & Generate Tickets (Excel + Word)",
            command=self._run_convert,
            width=42,
            style="Bold.TButton",
        )
        self.convert_btn.pack(side="top", anchor="w", pady=(0, 8))

        self.status = ttk.Label(action_box, text="● Ready", font=("Segoe UI", 9, "bold"), foreground="#2E7D32", anchor="w", wraplength=950)
        self.status.pack(fill="x")

    def _build_folder_row(self, parent, row_idx: int, kind: str, btn_text: str, with_checkbox: bool = False):
        """Builds one clean folder picker row per destination."""
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=row_idx, column=0, sticky="w", pady=4)

        folder_btn = ttk.Button(btn_frame, text=btn_text, command=lambda: self._pick_output_dir(kind), width=24)
        folder_btn.pack(side="left")

        if with_checkbox:
            cb = ttk.Checkbutton(btn_frame, text="Generate Raw Yarn (Filato)", variable=self.filato_enabled)
            cb.pack(side="left", padx=(10, 0))

        lab = ttk.Label(parent, text="No folder selected", foreground="grey", anchor="w")
        lab.grid(row=row_idx, column=1, sticky="ew", padx=(10, 0), pady=4)
        setattr(self, f"{kind}_dir_label", lab)

    def _restore(self):
        for key, attr in [
            ("biglietti_data_path", "data_path"),
            ("biglietti_dispo_path", "dispo_path"),
        ]:
            p = self._prefs.get(key)
            if p and Path(p).is_file():
                setattr(self, attr, Path(p))
                getattr(self, f"{key}_label").config(text=p, foreground="#111827")

        for kind in ("elvy", "med", "el_kamal", "filato"):
            p = self._prefs.get(f"biglietti_{kind}_output_dir")
            if p and Path(p).is_dir():
                setattr(self, f"{kind}_output_dir", Path(p))
                getattr(self, f"{kind}_dir_label").config(text=p, foreground="#111827")

        # Articoli: prefer the Marca-based cache (this tab's own upload),
        # fall back to Situazione's shared TITOLO-based table.
        try:
            n = len(load_articoli_marca_lookup())
            if n:
                self.articoli_label.config(text=f"{n} articles (Brand/Marca) available", foreground="#111827")
            else:
                n = len(load_articoli_titolo_map())
                if n:
                    self.articoli_label.config(text=f"{n} articles available (loaded in Situation)", foreground="#111827")
        except Exception:
            pass

        cache = load_densita_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.densita_path = Path(cache["source_path"])
            self.densita_label.config(text=cache["source_path"], foreground="#111827")

        cache = load_magazino_cache()
        if cache.get("source_path"):
            self.vmm_label.config(text=f"• VMM22: {Path(cache['source_path']).name} (Yarn Warehouse)", foreground="#111827")

        price_lookup, price_src = load_prezzo_lookup()
        if price_src:
            self.prezzi_label.config(text=f"• Price: {price_src} (Price List)", foreground="#111827")

    def _pick_file(self, title, filetypes=None):
        filetypes = filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")]
        return filedialog.askopenfilename(title=title, filetypes=filetypes)

    def _pick_data(self):
        p = self._pick_file("Select Order Data File")
        if p:
            self.data_path = Path(p)
            self.biglietti_data_path_label.config(text=p, foreground="#111827")
            self._save_prefs(biglietti_data_path=p)

    def _pick_dispo(self):
        p = self._pick_file("Select Dispo-Bagno File", filetypes=[("CSV/Excel", "*.csv;*.xlsx;*.xlsm"), ("All files", "*.*")])
        if p:
            self.dispo_path = Path(p)
            self.biglietti_dispo_path_label.config(text=p, foreground="#111827")
            self._save_prefs(biglietti_dispo_path=p)

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

        import biglietti_exporter
        try:
            marca_map, marca_errors = biglietti_exporter.load_articoli_marca_map(Path(p))
        except Exception as exc:  # noqa: BLE001
            self._logger.exception("Articoli (Marca) read failed")
            marca_map, marca_errors = {}, [str(exc)]

        try:
            import situazione_db
            import situazione_loaders
            titolo_df, titolo_errors = situazione_loaders.load_codes(p)
        except Exception as exc:  # noqa: BLE001
            self._logger.exception("Articoli (TITOLO) read failed")
            titolo_df, titolo_errors = None, [str(exc)]

        # A file only needs to match ONE of the two shapes (Articolo Filato
        # + Marca, or Articolo Filato + TITOLO) to be usable -- a file
        # missing 'Marca' is not an error, it just won't feed the richer
        # Titolo lookup. Only fail if NEITHER shape matched.
        if not marca_map and not titolo_df:
            errors = marca_errors + (titolo_errors or [])
            messagebox.showerror("Articles Error", "; ".join(errors) if errors else "Invalid file: no recognizable Articolo Filato column found.")
            return

        self.articoli_path = Path(p)
        parts = []
        if marca_map:
            save_articoli_cache(p)
            parts.append(f"{len(marca_map)} articles (Brand/Marca)")
        if titolo_df is not None and not titolo_df.empty:
            situazione_db.save_codes(titolo_df)
            parts.append(f"{len(titolo_df)} rows shared with Situation")
        msg = " + ".join(parts) + f" loaded from {Path(p).name}"
        self.articoli_label.config(text=msg, foreground="#111827")

    def _pick_densita(self):
        p = self._pick_file("Select Densita' Query.xlsx")
        if p:
            self.densita_path = Path(p)
            self.densita_label.config(text=p, foreground="#111827")
            save_densita_cache(p)

    def _set_status(self, msg: str):
        color = "#C62828" if "Error" in msg else ("#1565C0" if "progress" in msg or "Converting" in msg else "#2E7D32")
        prefix = "✖ " if "Error" in msg else ("⏳ " if "progress" in msg or "Converting" in msg else "✔ ")
        self.after(0, lambda: self.status.config(text=f"{prefix}{msg}", foreground=color))

    def _load_common_sources(self):
        """Titolo/KG-Densita/Color Tube-VMM22/Prezzo lookups, shared by every customer's export."""
        codes_map = load_articoli_marca_lookup() or load_articoli_titolo_map()

        densita_map = {}
        if self.densita_path and self.densita_path.is_file():
            densita_map, _errors = load_densita_query(self.densita_path)

        vmm_ratio_map = {}
        magazino_cache = load_magazino_cache()
        magazino_path = magazino_cache.get("source_path")
        if magazino_path and Path(magazino_path).is_file():
            vmm_ratio_map, _errors = load_vmm22_ratio_from_magazino(Path(magazino_path))

        price_lookup, _price_source = load_prezzo_lookup()

        return codes_map, densita_map, vmm_ratio_map, price_lookup

    def _run_convert(self):
        """Validate files and start conversion for detected customers in the background."""
        if not self.data_path or not self.data_path.is_file():
            return messagebox.showwarning("Missing Input", "Please select the Order Data file first.")

        try:
            order_format = detect_order_format(self.data_path)
        except Exception as exc:
            return messagebox.showerror("File Error", f"Could not read Order Data file:\n{exc}")

        # Check required output folders before starting
        if order_format == "EL_KAMAL":
            if not self.el_kamal_output_dir or not Path(self.el_kamal_output_dir).is_dir():
                self._pick_output_dir("el_kamal")
                if not self.el_kamal_output_dir or not Path(self.el_kamal_output_dir).is_dir():
                    return
        else:
            # ELVY_MED: peek records to verify customer codes present
            try:
                records, _raw = load_order(self.data_path, self.dispo_path)
                has_elvy = any(r.customer_code == "3009" for r in records)
                has_med = any(r.customer_code == "3004" for r in records)
            except Exception as exc:
                return messagebox.showerror("Read Error", f"Could not parse order records:\n{exc}")

            if not has_elvy and not has_med:
                return messagebox.showwarning(
                    "No Customer Rows",
                    "No customer rows found for 3009 (ELVY) or 3004 (MED) in the selected Order Data.",
                )

            if has_elvy and (not self.elvy_output_dir or not Path(self.elvy_output_dir).is_dir()):
                self._pick_output_dir("elvy")
                if not self.elvy_output_dir or not Path(self.elvy_output_dir).is_dir():
                    return

            if has_med and (not self.med_output_dir or not Path(self.med_output_dir).is_dir()):
                self._pick_output_dir("med")
                if not self.med_output_dir or not Path(self.med_output_dir).is_dir():
                    return

            if self.filato_enabled.get() and (not self.filato_output_dir or not Path(self.filato_output_dir).is_dir()):
                self._pick_output_dir("filato")
                if not self.filato_output_dir or not Path(self.filato_output_dir).is_dir():
                    return

        self.convert_btn.config(state="disabled")
        self._set_status("Converting orders and generating tickets in progress...")
        threading.Thread(target=self._worker_convert, daemon=True).start()

    def _worker_convert(self):
        created_items: list[str] = []
        try:
            order_format = detect_order_format(self.data_path)
            codes_map, densita_map, vmm_ratio_map, price_lookup = self._load_common_sources()

            template = Path.home() / "OneDrive" / "Desktop" / "Biglietti.docx"
            if not template.is_file():
                template = Path.home() / "Desktop" / "Biglietti.docx"
            if not template.is_file():
                local_tmpl = Path(__file__).resolve().parent / "Biglietti.docx"
                if local_tmpl.is_file():
                    template = local_tmpl
            if not template.is_file():
                raise ValueError("Template 'Biglietti.docx' not found on Desktop.")

            if order_format == "EL_KAMAL":
                records, raw = load_el_kamal_order(self.data_path, self.dispo_path)
                stem = build_el_kamal_stem(records)
                out_dir = Path(self.el_kamal_output_dir)
                xlsx = out_dir / f"{stem}_EL_KAMAL.xlsx"
                docx = out_dir / f"{stem}_EL_KAMAL_Biglietti.docx"

                enrich_records(records, "EL_KAMAL", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                export_workbook(xlsx, records, raw, include_filato=False, stem=stem, customer="EL_KAMAL")
                export_word(docx, template, records, stem=stem)
                created_items.append(f"• EL KAMAL: {len(records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")

                if self.filato_enabled.get():
                    created_items.append("• Raw Yarn (Filato): EL KAMAL's order data doesn't include a raw-yarn sheet, so nothing was generated for it.")

            else:
                records, raw = load_order(self.data_path, self.dispo_path)

                # Process ELVY
                elvy_records = [r for r in records if r.customer_code == "3009"]
                if elvy_records:
                    out_dir = Path(self.elvy_output_dir)
                    stem = build_output_stem(elvy_records, "ELVY")
                    xlsx = out_dir / f"{stem}_ELVY.xlsx"
                    docx = out_dir / f"{stem}_ELVY_Biglietti.docx"

                    enrich_records(elvy_records, "ELVY", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                    export_workbook(xlsx, elvy_records, raw, include_filato=False, stem=stem, customer="ELVY")
                    export_word(docx, template, elvy_records, stem=stem)
                    created_items.append(f"• ELVY: {len(elvy_records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")

                # Process MED
                med_records = [r for r in records if r.customer_code == "3004"]
                if med_records:
                    out_dir = Path(self.med_output_dir)
                    stem = build_output_stem(med_records, "MED")
                    xlsx = out_dir / f"{stem}_MED.xlsx"
                    docx = out_dir / f"{stem}_MED_Biglietti.docx"

                    enrich_records(med_records, "MED", codes_map=codes_map, densita_map=densita_map, vmm_ratio_map=vmm_ratio_map, price_lookup=price_lookup)
                    export_workbook(xlsx, med_records, raw, include_filato=True, stem=stem, customer="MED")
                    export_word(docx, template, med_records, stem=stem)
                    created_items.append(f"• MED: {len(med_records)} tickets ({docx.name})\n   ↳ Saved to: {xlsx}")

                # Process Raw Yarn (Filato x Tinturia) if enabled
                if self.filato_enabled.get():
                    if not self.filato_output_dir:
                        created_items.append("• Raw Yarn (Filato): no output folder was set, so it was skipped.")
                    elif not raw:
                        created_items.append("• Raw Yarn (Filato): the selected Order Data has no raw-yarn sheet (تحضير خيط خام), so nothing was generated.")
                    else:
                        out_dir = Path(self.filato_output_dir)
                        out_dir.mkdir(parents=True, exist_ok=True)
                        filato_file = out_dir / f"{self.data_path.stem}_Filato.xlsx"
                        export_filato_workbook(filato_file, raw)
                        created_items.append(f"• Raw Yarn (Filato): {filato_file.name}\n   ↳ Saved to: {filato_file}")

            summary_text = "\n\n".join(created_items)
            self._set_status("Completed — All order workbooks and dyeing tickets generated successfully.")
            self.after(0, lambda: messagebox.showinfo(
                "Conversion Complete",
                f"Conversion completed successfully!\n\n{summary_text}",
            ))

        except Exception as exc:
            self._logger.exception("Biglietti conversion failed")
            self._set_status(f"Error: {exc}")
            self.after(0, lambda: messagebox.showerror("Conversion Error", str(exc)))
        finally:
            self.after(0, lambda: self.convert_btn.config(state="normal"))
