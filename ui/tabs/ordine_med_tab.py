"""ui/tabs/ordine_med_tab.py — "Ordine Med" tab: same layout/button naming
as Ordine Kamal (Input / Raw Yarn Matching / ▶ Convert), reading the raw
ORDINE export and producing 'Ordine da creare' + Filato X Tinturia. See
ordine_med.py for the full reverse-engineering notes.
"""

from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ordine_med
from biglietti_exporter import load_articoli_marca_lookup, load_articoli_titolo_map, load_densita_query, load_prezzo_lookup
from densita_cache import load_densita_cache, save_densita_cache
from magazino_cache import load_magazino_cache, save_magazino_cache
from path_manager import save_source, source_path


class OrdineMedTab(ttk.Frame):
    def __init__(self, parent, situazione_tab=None, prefs: dict | None = None, save_prefs=None, logger=None, on_shared_cache_changed=None):
        super().__init__(parent)
        self._situazione_tab = situazione_tab
        self._prefs = prefs or {}
        self._save_prefs = save_prefs or (lambda **_kw: None)
        self._logger = logger
        self._on_shared_cache_changed = on_shared_cache_changed
        self.ordine_path: Path | None = None
        self.output_path: Path | None = None
        self.erp_folder: Path | None = None
        self.filato_folder: Path | None = None
        self.erp_enabled = tk.BooleanVar(value=False)
        self.filato_enabled = tk.BooleanVar(value=False)
        self.magazino_path: Path | None = None
        self.densita_path: Path | None = None
        self._build_ui()
        self._restore_saved_paths()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.columnconfigure(0, weight=1)

        sel_frame = ttk.LabelFrame(self, text="Input", padding=6)
        sel_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        sel_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(sel_frame, text="📄 Select Ordine File", command=self._on_select_ordine, width=18).grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        ttk.Button(sel_frame, text="💾 Output Folder", command=self._on_select_output, width=16).grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        ttk.Button(sel_frame, text="📁 ERP File Folder", command=self._on_select_erp_folder, width=18).grid(row=0, column=2, padx=4, pady=4, sticky="ew")
        self._lbl_ordine = ttk.Label(sel_frame, text="No Ordine file selected", foreground="grey", anchor="w")
        self._lbl_ordine.grid(row=1, column=0, sticky="ew", padx=4)
        self._lbl_output = ttk.Label(sel_frame, text="No output folder selected", foreground="grey", anchor="w")
        self._lbl_output.grid(row=1, column=1, sticky="ew", padx=4)
        self._lbl_erp = ttk.Label(sel_frame, text="No ERP folder selected", foreground="grey", anchor="w")
        self._lbl_erp.grid(row=1, column=2, sticky="ew", padx=4)

        self._lbl_shared = ttk.Label(
            self,
            text="Articoli (Titolo) and Listini (Prezzo): shared automatically with whatever was last uploaded elsewhere.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        self._lbl_shared.grid(row=1, column=0, sticky="ew", padx=4, pady=(2, 0))

        yarn_frame = ttk.LabelFrame(self, text="Raw Yarn Matching", padding=6)
        yarn_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 2))
        yarn_frame.columnconfigure(1, weight=1)
        ttk.Button(yarn_frame, text="📦 Select Magazino File", command=self._on_select_magazino, width=20).grid(row=0, column=0, padx=4, pady=3, sticky="w")
        self._lbl_magazino = ttk.Label(yarn_frame, text="No Magazino file selected", foreground="grey", anchor="w")
        self._lbl_magazino.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(yarn_frame, text="📦 Select Densita Query", command=self._on_select_densita, width=20).grid(row=1, column=0, padx=4, pady=3, sticky="w")
        self._lbl_densita = ttk.Label(yarn_frame, text="No Densita Query file selected", foreground="grey", anchor="w")
        self._lbl_densita.grid(row=1, column=1, sticky="ew", padx=4)

        erp_frame = ttk.LabelFrame(self, text="Also Extract ERP Files", padding=6)
        erp_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(3, 2))
        erp_frame.columnconfigure(1, weight=1)
        ttk.Button(
            erp_frame, text="📁 Select ERP Files Folder…", command=self._on_select_erp_folder, width=22
        ).grid(row=0, column=0, padx=(0, 6), pady=(2, 4), sticky="w")
        ttk.Checkbutton(
            erp_frame, text="Extract ERP order file (Ordine_MED_ERP.xlsx)", variable=self.erp_enabled
        ).grid(row=0, column=1, sticky="w", pady=(2, 4))
        self._lbl_erp = ttk.Label(erp_frame, text="No folder selected", foreground="grey", anchor="w")
        self._lbl_erp.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Button(
            erp_frame, text="📁 Select Filato Folder…", command=self._on_select_filato_folder, width=22
        ).grid(row=2, column=0, padx=(0, 6), pady=(4, 2), sticky="w")
        ttk.Checkbutton(
            erp_frame, text="Extract Filato X Tinturia.xlsx", variable=self.filato_enabled
        ).grid(row=2, column=1, sticky="w", pady=(4, 2))
        self._lbl_filato = ttk.Label(erp_frame, text="No folder selected", foreground="grey", anchor="w")
        self._lbl_filato.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        self.erp_enabled.trace_add("write", lambda *_: self._save_prefs(ordine_med_erp_enabled=self.erp_enabled.get()))
        self.filato_enabled.trace_add("write", lambda *_: self._save_prefs(ordine_med_filato_enabled=self.filato_enabled.get()))

        action_frame = ttk.Frame(self, padding=(6, 4))
        action_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=2)
        self._btn_convert = ttk.Button(
            action_frame, text="▶  Convert", command=self._on_convert, style="Accent.TButton"
        )
        self._btn_convert.pack(side="left")
        self._lbl_status = ttk.Label(action_frame, text="", foreground="grey")
        self._lbl_status.pack(side="left", padx=12)

    # --------------------------------------------------------------- events
    def _on_select_ordine(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self.ordine_path = Path(path)
        self._lbl_ordine.config(text=str(self.ordine_path), foreground="black")
        save_source("data_ordine", self.ordine_path)

    def _on_select_output(self):
        path = filedialog.askdirectory(title="Select Ordine MED output folder")
        if not path:
            return
        self.output_path = Path(path)
        self._lbl_output.config(text=str(self.output_path), foreground="black")
        self._save_prefs(ordine_med_output_dir=str(self.output_path))

    def _on_select_erp_folder(self):
        path = filedialog.askdirectory(title="Select ERP file folder")
        if not path:
            return
        self.erp_folder = Path(path)
        self._lbl_erp.config(text=str(self.erp_folder), foreground="black")
        self._save_prefs(ordine_med_erp_folder=str(self.erp_folder))

    def _on_select_filato_folder(self):
        path = filedialog.askdirectory(title="Select Filato X Tinturia output folder")
        if not path:
            return
        self.filato_folder = Path(path)
        self._lbl_filato.config(text=str(self.filato_folder), foreground="black")
        self._save_prefs(ordine_med_filato_folder=str(self.filato_folder))

    def _on_select_magazino(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not path:
            return
        self.magazino_path = Path(path)
        self._lbl_magazino.config(text=str(self.magazino_path), foreground="black")
        # Shared with Ordine Elvy / Ordine Kamal / Situazione Generale /
        # Magazino Filato -- same cache, same key, read by all of them.
        save_magazino_cache(path)
        if self._on_shared_cache_changed:
            self._on_shared_cache_changed()

    def _on_select_densita(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self.densita_path = Path(path)
        self._lbl_densita.config(text=str(self.densita_path), foreground="black")
        save_densita_cache(path)

    def _restore_saved_paths(self):
        output_str = self._prefs.get("ordine_med_output_dir") or self._prefs.get("ordine_med_output_path")
        if output_str:
            candidate = Path(output_str)
            if candidate.is_dir() or self._prefs.get("ordine_med_output_path"):
                self.output_path = candidate if candidate.is_dir() else candidate.parent
                self._lbl_output.config(text=str(self.output_path), foreground="black")
        erp_str = self._prefs.get("ordine_med_erp_folder")
        if erp_str and Path(erp_str).is_dir():
            self.erp_folder = Path(erp_str)
            self._lbl_erp.config(text=erp_str, foreground="black")
        filato_str = self._prefs.get("ordine_med_filato_folder")
        if filato_str and Path(filato_str).is_dir():
            self.filato_folder = Path(filato_str)
            self._lbl_filato.config(text=filato_str, foreground="black")
        self.erp_enabled.set(bool(self._prefs.get("ordine_med_erp_enabled", False)))
        self.filato_enabled.set(bool(self._prefs.get("ordine_med_filato_enabled", False)))

        ordine = source_path("data_ordine")
        if ordine and ordine.is_file():
            self.ordine_path = ordine
            self._lbl_ordine.config(text=str(ordine), foreground="black")

        cache = load_magazino_cache()
        if cache.get("source_path"):
            self.magazino_path = Path(cache["source_path"])
            self._lbl_magazino.config(text=f"✅ Shared: {cache['source_path']}", foreground="black")

        cache = load_densita_cache()
        if cache.get("source_path") and Path(cache["source_path"]).is_file():
            self.densita_path = Path(cache["source_path"])
            self._lbl_densita.config(text=cache["source_path"], foreground="black")

    def sync_shared_magazino(self):
        """Public hook (same shape as Kamal's) so switching to this tab
        after a Magazino upload elsewhere reflects it without re-browsing."""
        cache = load_magazino_cache()
        if cache.get("source_path"):
            self.magazino_path = Path(cache["source_path"])
            self._lbl_magazino.config(text=f"✅ Shared: {cache['source_path']}", foreground="black")

    def _set_status(self, msg: str):
        color = "#b00020" if msg.startswith("Error") else ("grey" if "..." in msg else "green")
        self.after(0, lambda: self._lbl_status.config(text=msg, foreground=color))

    def _on_convert(self):
        if not self.ordine_path or not self.ordine_path.is_file():
            return messagebox.showwarning("Missing input", "Select an Ordine file.")
        if not self.output_path or not self.output_path.is_dir():
            return self._on_select_output() or messagebox.showwarning("Missing output", "Select an output folder.")
        self._btn_convert.config(state="disabled")
        self._set_status("Converting...")
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
            if self.magazino_path and self.magazino_path.is_file():
                stock_map = ordine_med.load_filato_disponibile(self.magazino_path)
            availability = ordine_med.compute_filato_availability(records, densita_map, stock_map)

            output_dir = self.output_path
            output_file = output_dir / "Ordine_MED.xlsx"
            ordine_med.export_ordine_med_workbook(output_file, records, availability)
            if self.erp_enabled.get():
                if not self.erp_folder or not self.erp_folder.is_dir():
                    raise ValueError("ERP extraction is enabled but no ERP Files Folder is selected.")
                ordine_med.export_erp_order_workbook(
                    self.erp_folder / "Ordine_MED_ERP.xlsx", records
                )
            if self.filato_enabled.get():
                if not self.filato_folder or not self.filato_folder.is_dir():
                    raise ValueError("Filato extraction is enabled but no Filato folder is selected.")
                ordine_med.export_filato_availability_workbook(
                    self.filato_folder / "Filato X Tinturia.xlsx", availability
                )

            shortage_count = sum(1 for a in availability if a.disponibilita == "NO")
            extra = f" — {shortage_count} short" if shortage_count else ""
            self._set_status(f"✅ Done — {len(records)} row(s){extra}")
            self.after(0, lambda: messagebox.showinfo("Ordine MED", f"Created:\n{output_file}{extra}"))
        except Exception as exc:
            if self._logger:
                self._logger.exception("Ordine MED conversion failed")
            self._set_status(f"Error: {exc}")
            self.after(0, lambda: messagebox.showerror("Conversion error", str(exc)))
        finally:
            self.after(0, lambda: self._btn_convert.config(state="normal"))
