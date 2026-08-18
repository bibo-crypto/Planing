"""Tkinter page for extracting Biglietti, customer workbook and Filato."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from biglietti_exporter import export_filato_workbook, export_word, export_workbook, load_order


class BigliettiTab(ttk.Frame):
    def __init__(self, parent, prefs: dict, save_prefs, logger):
        super().__init__(parent, padding=12)
        self._prefs = prefs
        self._save_prefs = save_prefs
        self._logger = logger
        self.data_path: Path | None = None
        self.dispo_path: Path | None = None
        self.output_dir: Path | None = None
        self.filato_dir: Path | None = None
        self.include_filato = tk.BooleanVar(value=True)
        self._build()
        self._restore()

    def _build(self):
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text="Estrazione Ordine", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
        ttk.Label(self, text="Seleziona Data Ordine e Dispo-Bagno. Il cliente 3009 genera ELVY; il cliente 3004 genera MED.", foreground="grey", wraplength=900).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 12))
        self._row(2, "Data Ordine", "biglietti_data_path", self._pick_data)
        self._row(3, "Dispo-Bagno", "biglietti_dispo_path", self._pick_dispo)
        self._row(4, "Cartella output Word + Excel", "biglietti_output_dir", self._pick_output)
        ttk.Checkbutton(self, text="Estrai anche un file separato Filato x Tinturia", variable=self.include_filato).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 2))
        self.filato_btn = ttk.Button(self, text="Seleziona cartella Filato (opzionale)", command=self._pick_filato)
        self.filato_btn.grid(row=6, column=0, sticky="w", pady=3)
        self.filato_label = ttk.Label(self, text="Stessa cartella output se non selezionata", foreground="grey")
        self.filato_label.grid(row=6, column=1, sticky="w", padx=8)
        self.run_btn = ttk.Button(self, text="▶  Estrai file", command=self._run)
        self.run_btn.grid(row=7, column=0, sticky="w", pady=(14, 4))
        self.status = ttk.Label(self, text="Pronto", anchor="w")
        self.status.grid(row=7, column=1, columnspan=2, sticky="ew", padx=10)
        self.log = tk.Text(self, height=7, state="disabled", wrap="word")
        self.log.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        self.rowconfigure(8, weight=1)

    def _row(self, row, label, key, command):
        ttk.Button(self, text=f"📂  {label}", command=command, width=30).grid(row=row, column=0, sticky="w", pady=4)
        lab = ttk.Label(self, text="Non selezionato", foreground="grey", anchor="w")
        lab.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10)
        setattr(self, f"{key}_label", lab)

    def _restore(self):
        for key, attr in [("biglietti_data_path", "data_path"), ("biglietti_dispo_path", "dispo_path")]:
            p = self._prefs.get(key)
            if p and Path(p).is_file():
                setattr(self, attr, Path(p)); getattr(self, f"{key}_label").config(text=p, foreground="black")
        p = self._prefs.get("biglietti_output_dir")
        if p and Path(p).is_dir(): self.output_dir = Path(p); self.biglietti_output_dir_label.config(text=p, foreground="black")
        p = self._prefs.get("biglietti_filato_dir")
        if p and Path(p).is_dir(): self.filato_dir = Path(p); self.filato_label.config(text=p, foreground="black")

    def _pick_file(self, title):
        return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])

    def _pick_data(self):
        p = self._pick_file("Seleziona Data Ordine")
        if p: self.data_path = Path(p); self.biglietti_data_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_data_path=p)

    def _pick_dispo(self):
        p = self._pick_file("Seleziona Dispo-Bagno")
        if p: self.dispo_path = Path(p); self.biglietti_dispo_path_label.config(text=p, foreground="black"); self._save_prefs(biglietti_dispo_path=p)

    def _pick_output(self):
        p = filedialog.askdirectory(title="Seleziona cartella output")
        if p: self.output_dir = Path(p); self.biglietti_output_dir_label.config(text=p, foreground="black"); self._save_prefs(biglietti_output_dir=p)

    def _pick_filato(self):
        p = filedialog.askdirectory(title="Seleziona cartella Filato")
        if p: self.filato_dir = Path(p); self.filato_label.config(text=p, foreground="black"); self._save_prefs(biglietti_filato_dir=p)

    def _write(self, msg):
        self.after(0, lambda: (self.log.config(state="normal"), self.log.insert("end", msg + "\n"), self.log.see("end"), self.log.config(state="disabled")))

    def _run(self):
        if not self.data_path or not self.data_path.is_file(): return messagebox.showwarning("Input mancante", "Seleziona Data Ordine.")
        if not self.output_dir or not self.output_dir.is_dir(): return messagebox.showwarning("Output mancante", "Seleziona la cartella di output.")
        self.run_btn.config(state="disabled"); self.status.config(text="Estrazione in corso...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            records, raw = load_order(self.data_path, self.dispo_path)
            customer = "ELVY" if records[0].customer_code == "3009" else "MED"
            stem = self.data_path.stem
            xlsx = self.output_dir / f"{stem}_{customer}.xlsx"
            template = Path.home() / "OneDrive" / "Desktop" / "Biglietti.docx"
            if not template.is_file(): template = Path.home() / "Desktop" / "Biglietti.docx"
            if not template.is_file(): raise ValueError("Modello Biglietti.docx non trovato sul Desktop.")
            docx = self.output_dir / f"{stem}_{customer}_Biglietti.docx"
            export_workbook(xlsx, records, raw, include_filato=True)
            export_word(docx, template, records)
            if self.include_filato.get():
                target = self.filato_dir or self.output_dir
                export_filato_workbook(target / f"{stem}_{customer}_Filato.xlsx", raw)
            self._write(f"Creati {len(records)} biglietti: {docx.name}")
            self._write(f"Creato workbook: {xlsx.name}")
            self.after(0, lambda: self.status.config(text="Completato"))
        except Exception as exc:
            self._logger.exception("Biglietti extraction failed")
            self._write(f"Errore: {exc}")
            self.after(0, lambda: messagebox.showerror("Errore estrazione", str(exc)))
            self.after(0, lambda: self.status.config(text="Errore"))
        finally:
            self.after(0, lambda: self.run_btn.config(state="normal"))
