"""Tkinter page for extracting Biglietti, customer workbook and Filato."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from biglietti_exporter import (
    build_el_kamal_stem,
    build_output_stem,
    enrich_records,
    export_filato_workbook,
    export_word,
    export_workbook,
    load_articoli_marca_lookup,
    load_articoli_titolo_map,
    load_color_tube_magazino,
    load_densita_query,
    load_el_kamal_order,
    load_order,
    load_prezzo_lookup,
)
from articoli_cache import save_articoli_cache
from densita_cache import load_densita_cache, save_densita_cache
from vmm_magazino_cache import load_vmm_magazino_cache, save_vmm_magazino_cache


class BigliettiTab(ttk.Frame):
    def __init__(self, parent, prefs: dict, save_prefs, logger):
        super().__init__(parent, padding=12)
        self._prefs = prefs
        self._save_prefs = save_prefs
        self._logger = logger
        self.data_path: Path | None = None
        self.dispo_path: Path | None = None
        self.el_kamal_data_path: Path | None = None
        self.el_kamal_dispo_path: Path | None = None
        self.elvy_output_dir: Path | None = None
        self.med_output_dir: Path | None = None
        self.el_kamal_output_dir: Path | None = None
        self.filato_output_dir: Path | None = None
        self.articoli_path: Path | None = None
        self.densita_path: Path | None = None
        self.vmm_path: Path | None = None
        self._build()
        self._restore()

    def _build(self):
        self.columnconfigure(1, weight=1)
        row = 0
        ttk.Label(self, text="Estrazione Ordine", font=("Segoe UI", 16, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 4)); row += 1
        ttk.Label(self, text="Cliente 3009 = ELVY, cliente 3004 = MED, entrambi dallo stesso Data Ordine.", foreground="grey", wraplength=900).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 8)); row += 1
        row = self._row(row, "Data Ordine (ELVY/MED)", "biglietti_data_path", self._pick_data)
        row = self._row(row, "Dispo-Bagno (ELVY/MED)", "biglietti_dispo_path", self._pick_dispo)

        row += 1
        ttk.Separator(self, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 8)); row += 1
        ttk.Label(self, text="EL KAMAL — file separati: Data Ordine (Sheet1 già completo) + Dispo-Bagno (CSV)", foreground="grey", wraplength=900).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 8)); row += 1
        row = self._row(row, "Data Ordine (EL KAMAL)", "biglietti_el_kamal_data_path", self._pick_el_kamal_data)
        row = self._row(row, "Dispo-Bagno (EL KAMAL, CSV)", "biglietti_el_kamal_dispo_path", self._pick_el_kamal_dispo)

        row += 1
        ttk.Separator(self, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 8)); row += 1
        ttk.Label(self, text="Fonti opzionali per Titolo / KG / Densita / Color Tube / VMM22 / Prezzo (condivise fra tutti i clienti)", foreground="grey", wraplength=900).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1

        self.articoli_btn = ttk.Button(self, text="📂  Articoli (Titolo) — opzionale", command=self._pick_articoli, width=32)
        self.articoli_btn.grid(row=row, column=0, sticky="w", pady=3)
        self.articoli_label = ttk.Label(self, text="Condiviso con il tab Situazione", foreground="grey")
        self.articoli_label.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10); row += 1

        self.densita_btn = ttk.Button(self, text="📂  Densita' Query (KG/Densita) — opzionale", command=self._pick_densita, width=32)
        self.densita_btn.grid(row=row, column=0, sticky="w", pady=3)
        self.densita_label = ttk.Label(self, text="Non selezionato — KG resta il peso grezzo del filato", foreground="grey")
        self.densita_label.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10); row += 1

        self.vmm_btn = ttk.Button(self, text="📂  Magazino Color Tube (VMM22) — opzionale", command=self._pick_vmm, width=32)
        self.vmm_btn.grid(row=row, column=0, sticky="w", pady=3)
        self.vmm_label = ttk.Label(self, text="Non selezionato — Color Tube/VMM22 resteranno vuoti", foreground="grey")
        self.vmm_label.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10); row += 1

        self.prezzi_label = ttk.Label(self, text="Prezzo: userà l'ultimo Listini caricato nel tab Prezzi (se presente)", foreground="grey")
        self.prezzi_label.grid(row=row, column=0, columnspan=3, sticky="w", pady=(4, 0)); row += 1

        row += 1
        ttk.Separator(self, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 8)); row += 1
        ttk.Label(self, text="Estrazione — ogni pulsante salva nella propria cartella (ricordata finché non la cambi)", foreground="grey").grid(row=row, column=0, columnspan=3, sticky="w"); row += 1

        row = self._export_row(row, "elvy", "ELVY", "🧵  Estrai ELVY (Excel + Word)")
        row = self._export_row(row, "med", "MED", "🧵  Estrai MED (Excel + Word)")
        row = self._export_row(row, "el_kamal", "EL_KAMAL", "🧵  Estrai EL KAMAL (Excel + Word)")
        row = self._export_row(row, "filato", None, "🧶  Estrai Filato x Tinturia (condiviso)")

        self.status = ttk.Label(self, text="Pronto", anchor="w")
        self.status.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(10, 4)); row += 1
        self.log = tk.Text(self, height=7, state="disabled", wrap="word")
        self.log.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
        self.rowconfigure(row, weight=1)

    def _row(self, row, label, key, command):
        ttk.Button(self, text=f"📂  {label}", command=command, width=32).grid(row=row, column=0, sticky="w", pady=4)
        lab = ttk.Label(self, text="Non selezionato", foreground="grey", anchor="w")
        lab.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10)
        setattr(self, f"{key}_label", lab)
        return row + 1

    def _export_row(self, row, kind: str, customer: str | None, run_text: str):
        """kind: 'elvy' | 'med' | 'el_kamal' | 'filato'. customer is None
        for filato (customer-agnostic). Builds a folder label + Estrai
        button pair, both reading/writing self.<kind>_output_dir."""
        lab = ttk.Label(self, text="Cartella non ancora scelta", foreground="grey", anchor="w")
        lab.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10)
        setattr(self, f"{kind}_dir_label", lab)
        run_btn = ttk.Button(self, text=run_text, command=lambda: self._run(kind, customer))
        run_btn.grid(row=row, column=0, sticky="w", pady=(4, 0))
        setattr(self, f"{kind}_run_btn", run_btn)
        change_btn = ttk.Button(self, text="Cambia cartella...", command=lambda: self._pick_output_dir(kind))
        change_btn.grid(row=row + 1, column=0, sticky="w", pady=(0, 6))
        return row + 2

    def _restore(self):
        for key, attr in [
            ("biglietti_data_path", "data_path"),
            ("biglietti_dispo_path", "dispo_path"),
            ("biglietti_el_kamal_data_path", "el_kamal_data_path"),
            ("biglietti_el_kamal_dispo_path", "el_kamal_dispo_path"),
        ]:
            p = self._prefs.get(key)
            if p and Path(p).is_file():
                setattr(self, attr, Path(p)); getattr(self, f"{key}_label").config(text=p, foreground="black")

        for kind in ("elvy", "med", "el_kamal", "filato"):
            p = self._prefs.get(f"biglietti_{kind}_output_dir")
            if p and Path(p).is_dir():
                setattr(self, f"{kind}_output_dir", Path(p))
                getattr(self, f"{kind}_dir_label").config(text=p, foreground="black")

        # Articoli: prefer the Marca-based cache (this tab's own upload),
        # fall back to Situazione's shared TITOLO-based table.
        try:
            n = len(load_articoli_marca_lookup())
            if n:
                self.articoli_label.config(text=f"{n} articoli (Marca) disponibili", foreground="black")
            else:
                n = len(load_articoli_titolo_map())
                if n:
                    self.articoli_label.config(text=f"{n} articoli disponibili (caricati in Situazione)", foreground="black")
        except Exception:
            pass

        cache = load_densita_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.densita_path = Path(cache["source_path"])
            self.densita_label.config(text=cache["source_path"], foreground="black")

        cache = load_vmm_magazino_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.vmm_path = Path(cache["source_path"])
            self.vmm_label.config(text=cache["source_path"], foreground="black")

    def _pick_file(self, title, filetypes=None):
        filetypes = filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")]
        return filedialog.askopenfilename(title=title, filetypes=filetypes)

    def _pick_data(self):
        p = self._pick_file("Seleziona Data Ordine")
        if p: self.data_path = Path(p); self.biglietti_data_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_data_path=p)

    def _pick_dispo(self):
        p = self._pick_file("Seleziona Dispo-Bagno")
        if p: self.dispo_path = Path(p); self.biglietti_dispo_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_dispo_path=p)

    def _pick_el_kamal_data(self):
        p = self._pick_file("Seleziona Data Ordine EL KAMAL")
        if p: self.el_kamal_data_path = Path(p); self.biglietti_el_kamal_data_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_el_kamal_data_path=p)

    def _pick_el_kamal_dispo(self):
        p = self._pick_file("Seleziona Dispo-Bagno EL KAMAL", filetypes=[("CSV/Excel", "*.csv;*.xlsx;*.xlsm"), ("All files", "*.*")])
        if p: self.el_kamal_dispo_path = Path(p); self.biglietti_el_kamal_dispo_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_el_kamal_dispo_path=p)

    def _pick_output_dir(self, kind: str):
        p = filedialog.askdirectory(title=f"Seleziona cartella output ({kind.upper()})")
        if p:
            setattr(self, f"{kind}_output_dir", Path(p))
            getattr(self, f"{kind}_dir_label").config(text=p, foreground="black")
            self._save_prefs(**{f"biglietti_{kind}_output_dir": p})

    def _pick_articoli(self):
        p = self._pick_file("Seleziona Articoli.xlsx")
        if not p:
            return
        # Titolo for Biglietti's own tickets: Marca column (found on
        # 'Sheet2' for files like the EL KAMAL bundle, or the active sheet
        # otherwise) -- this is the primary, always attempted.
        try:
            import biglietti_exporter
            marca_map, marca_errors = biglietti_exporter.load_articoli_marca_map(Path(p))
        except Exception as exc:  # noqa: BLE001
            self._logger.exception("Articoli (Marca) upload failed")
            messagebox.showerror("Errore Articoli", str(exc))
            return
        if not marca_map:
            messagebox.showerror("Errore Articoli", "; ".join(marca_errors) if marca_errors else "File non valido.")
            return
        save_articoli_cache(p)
        self.articoli_path = Path(p)
        msg = f"{len(marca_map)} articoli (Marca) caricati da {Path(p).name}"

        # Also feed Situazione's shared TITOLO-based table when this file's
        # first sheet happens to match that shape -- best-effort, never
        # blocks the Marca upload above if it doesn't.
        try:
            import situazione_db
            import situazione_loaders
            df, _errors = situazione_loaders.load_codes(p)
            if df is not None:
                situazione_db.save_codes(df)
                msg += f" (e {len(df)} righe condivise con Situazione)"
        except Exception:
            pass

        self.articoli_label.config(text=msg, foreground="black")

    def _pick_densita(self):
        p = self._pick_file("Seleziona Densita' Query.xlsx")
        if p:
            self.densita_path = Path(p)
            self.densita_label.config(text=p, foreground="black")
            save_densita_cache(p)

    def _pick_vmm(self):
        p = self._pick_file("Seleziona il file Magazino (Color Tube)")
        if p:
            self.vmm_path = Path(p)
            self.vmm_label.config(text=p, foreground="black")
            save_vmm_magazino_cache(p)

    def _write(self, msg):
        self.after(0, lambda: (self.log.config(state="normal"), self.log.insert("end", msg + "\n"), self.log.see("end"), self.log.config(state="disabled")))

    def _run(self, kind: str, customer: str | None):
        if kind == "el_kamal":
            if not self.el_kamal_data_path or not self.el_kamal_data_path.is_file():
                return messagebox.showwarning("Input mancante", "Seleziona Data Ordine EL KAMAL.")
        elif not self.data_path or not self.data_path.is_file():
            return messagebox.showwarning("Input mancante", "Seleziona Data Ordine.")
        out_dir = getattr(self, f"{kind}_output_dir")
        if not out_dir or not Path(out_dir).is_dir():
            self._pick_output_dir(kind)
            out_dir = getattr(self, f"{kind}_output_dir")
            if not out_dir or not Path(out_dir).is_dir():
                return
        btn = getattr(self, f"{kind}_run_btn")
        btn.config(state="disabled")
        self.status.config(text=f"Estrazione {kind.upper()} in corso...")
        threading.Thread(target=self._worker, args=(kind, customer, Path(out_dir)), daemon=True).start()

    def _load_common_sources(self):
        """Titolo/KG-Densita/Color Tube-VMM22/Prezzo lookups, shared by
        every customer's export. Logs what it found (or didn't) either way."""
        codes_map = load_articoli_marca_lookup() or load_articoli_titolo_map()
        self._write(f"Titolo: {len(codes_map)} articoli disponibili" if codes_map else "Titolo: nessun Articoli.xlsx caricato")

        densita_map, densita_errors = ({}, [])
        if self.densita_path and self.densita_path.is_file():
            densita_map, densita_errors = load_densita_query(self.densita_path)
            for e in densita_errors: self._write(f"Densita' Query: {e}")
            self._write(f"KG/Densita: {len(densita_map)} partite caricate da {self.densita_path.name}")
        else:
            self._write("KG/Densita: nessun file caricato")

        vmm_map, vmm_errors = ({}, [])
        if self.vmm_path and self.vmm_path.is_file():
            vmm_map, vmm_errors = load_color_tube_magazino(self.vmm_path)
            for e in vmm_errors: self._write(f"Magazino Color Tube: {e}")
            self._write(f"Color Tube/VMM22: {len(vmm_map)} partite caricate da {self.vmm_path.name}")
        else:
            self._write("Color Tube/VMM22: nessun file caricato, colonne vuote")

        price_lookup, price_source = load_prezzo_lookup()
        self._write(f"Prezzo: {len(price_lookup)} righe da {price_source}" if price_lookup else "Prezzo: nessun Listini caricato nel tab Prezzi")

        return codes_map, densita_map, vmm_map, price_lookup

    def _worker(self, kind: str, customer: str | None, out_dir: Path):
        btn = getattr(self, f"{kind}_run_btn")
        try:
            if kind == "filato":
                records, raw = load_order(self.data_path, self.dispo_path)
                export_filato_workbook(out_dir / f"{self.data_path.stem}_Filato.xlsx", raw)
                self._write(f"Creato workbook Filato x Tinturia in {out_dir}")
                self.after(0, lambda: self.status.config(text="Completato"))
                return

            if kind == "el_kamal":
                records, raw = load_el_kamal_order(self.el_kamal_data_path, self.el_kamal_dispo_path)
                stem = build_el_kamal_stem(records)
            else:
                records, raw = load_order(self.data_path, self.dispo_path)
                code = "3009" if customer == "ELVY" else "3004"
                records = [r for r in records if r.customer_code == code]
                if not records:
                    raise ValueError(f"Nel Data Ordine selezionato non ci sono righe per il cliente {customer}.")
                stem = build_output_stem(records, customer)

            xlsx = out_dir / f"{stem}_{customer}.xlsx"
            template = Path.home() / "OneDrive" / "Desktop" / "Biglietti.docx"
            if not template.is_file(): template = Path.home() / "Desktop" / "Biglietti.docx"
            if not template.is_file(): raise ValueError("Modello Biglietti.docx non trovato sul Desktop.")
            docx = out_dir / f"{stem}_{customer}_Biglietti.docx"

            codes_map, densita_map, vmm_map, price_lookup = self._load_common_sources()
            enrich_records(records, customer, codes_map=codes_map, densita_map=densita_map, vmm_map=vmm_map, price_lookup=price_lookup)

            # MED's own workbook gets the Filato x Tinturia sheet embedded
            # again (separate from the standalone Filato button); other
            # customers keep it out to avoid duplicating that data.
            export_workbook(xlsx, records, raw, include_filato=(customer == "MED"), stem=stem, customer=customer)
            export_word(docx, template, records, stem=stem)
            self._write(f"Creati {len(records)} biglietti: {docx.name}")
            self._write(f"Creato workbook: {xlsx.name}")
            self.after(0, lambda: self.status.config(text="Completato"))
        except Exception as exc:
            self._logger.exception("Biglietti extraction failed")
            self._write(f"Errore: {exc}")
            self.after(0, lambda: messagebox.showerror("Errore estrazione", str(exc)))
            self.after(0, lambda: self.status.config(text="Errore"))
        finally:
            self.after(0, lambda: btn.config(state="normal"))
