"""ui/tabs/ordine_med_tab.py — "Ordine Med" tab: extracts the "Ordine da
creare" ERP-import workbook plus a raw-yarn availability/shortage check
("Filato X Tinturia") from a raw ORDINE export, the same way the
reference ORDINE_MED-MACRO.xlsm's Power Query does -- plus automatic
Consegna scheduling per machine queue (see ordine_med.py for the full
reverse-engineering notes).
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import ordine_med
from biglietti_exporter import load_articoli_marca_lookup, load_articoli_titolo_map, load_densita_query, load_prezzo_lookup
from densita_cache import load_densita_cache, save_densita_cache
from filato_disponibile_cache import load_filato_disponibile_cache, save_filato_disponibile_cache


class OrdineMedTab(ttk.Frame):
    def __init__(self, parent, situazione_tab=None, prefs: dict | None = None, save_prefs=None, logger=None):
        super().__init__(parent, padding=12)
        self._situazione_tab = situazione_tab
        self._prefs = prefs or {}
        self._save_prefs = save_prefs or (lambda **_kw: None)
        self._logger = logger
        self.ordine_path: Path | None = None
        self.filato_disp_path: Path | None = None
        self.densita_path: Path | None = None
        self.output_dir: Path | None = None
        self._build()
        self._restore()

    def _build(self):
        self.columnconfigure(1, weight=1)
        style = ttk.Style(self)
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))

        row = 0
        ttk.Label(self, text="Ordine MED", font=("Segoe UI", 15, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 4)); row += 1
        ttk.Label(
            self,
            text="Legge il foglio ORDINE ed estrae 'Ordine da creare' (dati per il sistema) + 'Filato X Tinturia' "
                 "(disponibilità/scoperto del filato) e la Consegna automatica per macchina.",
            foreground="grey", wraplength=950,
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 10)); row += 1

        row = self._row(row, "📂  Ordine (input)", "ordine_path", self._pick_ordine)
        row = self._row(row, "📂  Filato Disponibile (opzionale)", "filato_disp_path", self._pick_filato_disp)
        row = self._row(row, "📂  Densita' Query (opzionale, KG del Filato)", "densita_path", self._pick_densita)

        ttk.Label(
            self, text="Titolo e Prezzo usano gli stessi Articoli (Marca) / Listini già caricati altrove nel programma.",
            foreground="grey",
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(4, 10)); row += 1

        ttk.Separator(self, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(4, 10)); row += 1

        ttk.Button(self, text="📁  Cartella output", command=self._pick_output_dir, width=24).grid(row=row, column=0, sticky="w", pady=4)
        self.output_dir_label = ttk.Label(self, text="Cartella non ancora scelta", foreground="grey", anchor="w")
        self.output_dir_label.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10); row += 1

        self.run_btn = ttk.Button(self, text="⚡  Estrai Ordine MED (Excel)", command=self._run, style="Bold.TButton")
        self.run_btn.grid(row=row, column=0, sticky="w", pady=(10, 4)); row += 1

        self.status = ttk.Label(self, text="● Pronto", font=("Segoe UI", 9, "bold"), foreground="#2E7D32", anchor="w", wraplength=950)
        self.status.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(6, 4))

    def _row(self, row, label, key, command):
        ttk.Button(self, text=label, command=command, width=32).grid(row=row, column=0, sticky="w", pady=4)
        lab = ttk.Label(self, text="Non selezionato", foreground="grey", anchor="w")
        lab.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10)
        setattr(self, f"{key}_label", lab)
        return row + 1

    def _restore(self):
        p = self._prefs.get("ordine_med_output_dir")
        if p and Path(p).is_dir():
            self.output_dir = Path(p)
            self.output_dir_label.config(text=p, foreground="#111827")

        cache = load_filato_disponibile_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.filato_disp_path = Path(cache["source_path"])
            self.filato_disp_path_label.config(text=cache["source_path"], foreground="#111827")

        cache = load_densita_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.densita_path = Path(cache["source_path"])
            self.densita_path_label.config(text=cache["source_path"], foreground="#111827")

    def _pick_file(self, title, filetypes=None):
        filetypes = filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")]
        return filedialog.askopenfilename(title=title, filetypes=filetypes)

    def _pick_ordine(self):
        p = self._pick_file("Seleziona il file Ordine")
        if p:
            self.ordine_path = Path(p)
            self.ordine_path_label.config(text=p, foreground="#111827")

    def _pick_filato_disp(self):
        p = self._pick_file("Seleziona Filato Disponibile")
        if p:
            self.filato_disp_path = Path(p)
            self.filato_disp_path_label.config(text=p, foreground="#111827")
            save_filato_disponibile_cache(p)

    def _pick_densita(self):
        p = self._pick_file("Seleziona Densita' Query.xlsx")
        if p:
            self.densita_path = Path(p)
            self.densita_path_label.config(text=p, foreground="#111827")
            save_densita_cache(p)

    def _pick_output_dir(self):
        p = filedialog.askdirectory(title="Seleziona cartella output")
        if p:
            self.output_dir = Path(p)
            self.output_dir_label.config(text=p, foreground="#111827")
            self._save_prefs(ordine_med_output_dir=p)

    def _set_status(self, msg: str):
        color = "#C62828" if "Errore" in msg else ("#1565C0" if "corso" in msg else "#2E7D32")
        prefix = "✖ " if "Errore" in msg else ("⏳ " if "corso" in msg else "✔ ")
        self.after(0, lambda: self.status.config(text=f"{prefix}{msg}", foreground=color))

    def _run(self):
        if not self.ordine_path or not self.ordine_path.is_file():
            return messagebox.showwarning("Input mancante", "Seleziona il file Ordine.")
        if not self.output_dir or not Path(self.output_dir).is_dir():
            self._pick_output_dir()
            if not self.output_dir or not Path(self.output_dir).is_dir():
                return
        self.run_btn.config(state="disabled")
        self._set_status("Estrazione in corso...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            records = ordine_med.load_ordine(self.ordine_path)
            ordine_med.compute_mc_and_gruppo(records)

            codes_map = load_articoli_marca_lookup() or load_articoli_titolo_map()
            ordine_med.compute_titolo(records, codes_map)
            ordine_med.compute_cliente(records)
            ordine_med.compute_commento(records)

            price_lookup, _src = load_prezzo_lookup()
            ordine_med.compute_prezzo(records, price_lookup)
            ordine_med.compute_prezzo_plus2(records)

            dfm_pairs = set()
            try:
                import situazione_db
                dfm_info = situazione_db.get_all_uploads().get("dfm", {})
                dfm_path = dfm_info.get("file_path")
                if dfm_path and Path(dfm_path).is_file():
                    dfm_pairs = ordine_med.load_dfm_articolo_colore(Path(dfm_path))
            except Exception:
                dfm_pairs = set()
            ordine_med.compute_check_articolo(records, dfm_pairs)

            machine_totals = {}
            try:
                st = self._situazione_tab
                situation_df = getattr(st, "current_df", None)
                copertura_df = getattr(st, "loaded_frames", {}).get("copertura") if st is not None else None
                if situation_df is not None and copertura_df is not None:
                    import situazione_logic
                    machine_totals = situazione_logic.compute_machine_totals(situation_df, copertura_df)
            except Exception:
                machine_totals = {}
            ordine_med.assign_consegna(records, machine_totals)
            ordine_med.compute_data_riconsegna(records)

            densita_map = {}
            if self.densita_path and self.densita_path.is_file():
                densita_map, _errors = load_densita_query(self.densita_path)

            stock_map = {}
            if self.filato_disp_path and self.filato_disp_path.is_file():
                stock_map = ordine_med.load_filato_disponibile(self.filato_disp_path)
            availability = ordine_med.compute_filato_availability(records, densita_map, stock_map)

            out_path = Path(self.output_dir) / f"{self.ordine_path.stem}_Ordine_da_creare.xlsx"
            ordine_med.export_ordine_med_workbook(out_path, records, availability)

            shortage_count = sum(1 for a in availability if a.disponibilita == "NO")
            extra = f" — {shortage_count} filati in scoperto" if shortage_count else ""
            self._set_status(f"Completato — {len(records)} righe, {out_path.name}{extra}")
            self.after(0, lambda: messagebox.showinfo("Ordine MED", f"Creato:\n{out_path}{extra}"))
        except Exception as exc:
            if self._logger:
                self._logger.exception("Ordine MED extraction failed")
            self._set_status(f"Errore: {exc}")
            self.after(0, lambda: messagebox.showerror("Errore estrazione", str(exc)))
        finally:
            self.after(0, lambda: self.run_btn.config(state="normal"))
