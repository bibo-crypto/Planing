"""
kamal_tab.py
"Ordine Kamal" tab: upload one or more Kamal order-letter PDFs, match each
colour's Articolo via DFM (shared with Data Elvy) and optionally match raw
yarn stock (Magazino), then export an Excel workbook.
"""
import os
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from dfm_lookup import load_dfm_cache, load_dfm_entries_by_prefix
from kamal_excel_exporter import KamalExcelExporter
from kamal_parser import KamalParser
from magazino_cache import load_magazino_cache, save_magazino_cache
from lotti_cache import load_lotti_cache, save_lotti_cache
from ordine_kamal import (
    KAMAL_ARTICLE_PREFIX,
    assign_ordine_kamal_machines,
    build_ordine_kamal_rows,
    match_by_lotto,
)
from ordini_elvy import export_filato_full, export_ordini_full, match_raw_yarn, read_filato_tinturia_sheet
from utils import logger, load_settings, save_settings
from typing import Callable

import magazino_logic
import lotti_logic


class KamalTab(ttk.Frame):
    """Embeddable 'Ordine Kamal' tab."""

    def __init__(self, master, on_shared_cache_changed: Callable[[], None] | None = None):
        super().__init__(master)
        self._pdf_paths: list[Path] = []
        self._raw_yarn_path: Path | None = None
        self._lotti_path: Path | None = None
        self._output_path: Path | None = None
        self._kamal_erp_export_dir: Path | None = None
        self._kamal_update_erp_file = tk.BooleanVar(value=False)
        self._shared_dfm_path = ""
        self._on_shared_cache_changed = on_shared_cache_changed
        self._prefs = load_settings()

        self._build_ui()
        self._restore_saved_paths()
        # DFM sync is cheap (reads a small cached path, not the DFM Excel
        # itself) and has no competing per-tab preference to conflict with,
        # unlike Magazino/LOTTI above -- so it's safe to do proactively here,
        # fixing the gap where a returning session wouldn't reflect an
        # already-uploaded DFM until the user re-uploads it this session.
        self.sync_shared_dfm()
        self.sync_shared_magazino()
        self.sync_shared_lotti()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        sel_frame = ttk.LabelFrame(self, text="Input", padding=6)
        sel_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        sel_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(sel_frame, text="📄 Select PDF(s)", command=self._on_select_pdfs, width=16).grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        ttk.Button(sel_frame, text="💾 Output File", command=self._on_select_output, width=16).grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        self._lbl_pdfs = ttk.Label(sel_frame, text="No PDFs selected", foreground="grey", anchor="w")
        self._lbl_pdfs.grid(row=1, column=0, sticky="ew", padx=4)
        self._lbl_output = ttk.Label(sel_frame, text="No output file selected", foreground="grey", anchor="w")
        self._lbl_output.grid(row=1, column=1, columnspan=2, sticky="ew", padx=4)

        self._lbl_dfm = ttk.Label(
            self,
            text="DFM: upload a reference on the Data Elvy tab — it is shared automatically.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        self._lbl_dfm.grid(row=1, column=0, sticky="ew", padx=4, pady=(2, 0))

        lotti_frame = ttk.LabelFrame(self, text="Raw Yarn Matching", padding=6)
        lotti_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 2))
        lotti_frame.columnconfigure(1, weight=1)
        ttk.Button(lotti_frame, text="📦 Select LOTTI File", command=self._on_select_lotti, width=20).grid(row=0, column=0, padx=4, pady=3, sticky="w")
        self._lbl_lotti = ttk.Label(lotti_frame, text="No LOTTI file selected", foreground="grey", anchor="w")
        self._lbl_lotti.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(lotti_frame, text="📦 Select Magazino File", command=self._on_select_raw_yarn, width=20).grid(row=1, column=0, padx=4, pady=3, sticky="w")
        self._lbl_raw_yarn = ttk.Label(lotti_frame, text="No Magazino file selected", foreground="grey", anchor="w")
        self._lbl_raw_yarn.grid(row=1, column=1, sticky="ew", padx=4)

        erp_frame = ttk.LabelFrame(self, text="Also Extract ERP Files", padding=6)
        erp_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(3, 2))
        erp_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            erp_frame,
            text="After converting, extract \"EXCEL PER ORDINE VENDITA EGITTO\" and "
                 "\"Filato x Tinturia\" into the folder below — each file is (re)written "
                 "fresh, fully formatted, every Convert",
            variable=self._kamal_update_erp_file,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Button(
            erp_frame, text="📁 Select ERP Files Folder…", command=self._on_select_erp_folder, width=22
        ).grid(row=1, column=0, padx=(0, 6), pady=(4, 0), sticky="w")

        self._lbl_erp_dir = ttk.Label(
            erp_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._lbl_erp_dir.grid(row=1, column=1, sticky="ew", pady=(4, 0))

        self._kamal_update_erp_file.trace_add(
            "write",
            lambda *_: self._save_prefs(kamal_update_erp_file=self._kamal_update_erp_file.get()),
        )

        action_frame = ttk.Frame(self, padding=(6, 4))
        action_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=2)
        self._btn_convert = ttk.Button(
            action_frame, text="▶  Convert", command=self._on_convert, style="Accent.TButton"
        )
        self._btn_convert.pack(side="left")
        self._lbl_status = ttk.Label(action_frame, text="", foreground="grey")
        self._lbl_status.pack(side="left", padx=12)

    # --------------------------------------------------------------- events
    def _on_select_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        self._pdf_paths = [Path(p) for p in paths]
        self._lbl_pdfs.config(
            text=f"{len(self._pdf_paths)} file(s) selected", foreground="black"
        )

    def _on_select_lotti(self):
        path = filedialog.askopenfilename(
            title="Select the LOTTI (raw yarn lot reference) Excel export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self._prefs.get("kamal_last_dir") or None,
        )
        if path:
            self._lotti_path = Path(path)
            self._lbl_lotti.config(text=str(self._lotti_path), foreground="black")
            self._save_prefs(kamal_lotti_path=str(self._lotti_path), kamal_last_dir=str(Path(path).parent))
            save_lotti_cache(self._lotti_path)
            if self._on_shared_cache_changed:
                self._on_shared_cache_changed()

    def _on_select_raw_yarn(self):
        path = filedialog.askopenfilename(
            title="Select the Magazino (raw yarn) Excel export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self._prefs.get("kamal_last_dir") or None,
        )
        if path:
            self._raw_yarn_path = Path(path)
            self._lbl_raw_yarn.config(text=str(self._raw_yarn_path), foreground="black")
            self._save_prefs(kamal_raw_yarn_path=str(self._raw_yarn_path), kamal_last_dir=str(Path(path).parent))
            save_magazino_cache(self._raw_yarn_path)
            if self._on_shared_cache_changed:
                self._on_shared_cache_changed()

    def _on_select_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Ordine_Kamal.xlsx",
            initialdir=self._prefs.get("kamal_last_dir") or None,
        )
        if path:
            self._output_path = Path(path)
            self._lbl_output.config(text=str(self._output_path), foreground="black")
            self._save_prefs(kamal_output_path=str(self._output_path), kamal_last_dir=str(Path(path).parent))

    def _on_select_erp_folder(self):
        path = filedialog.askdirectory(
            title="Select the folder for the ERP export files",
            initialdir=self._prefs.get("kamal_erp_export_dir") or self._prefs.get("kamal_last_dir") or None,
        )
        if path:
            self._kamal_erp_export_dir = Path(path)
            self._lbl_erp_dir.config(text=str(self._kamal_erp_export_dir), foreground="black")
            self._save_prefs(kamal_erp_export_dir=str(self._kamal_erp_export_dir), kamal_last_dir=str(path))

    def _save_prefs(self, **kwargs: object) -> None:
        self._prefs.update(kwargs)
        save_settings(self._prefs)

    def _restore_saved_paths(self) -> None:
        # Restore the ERP update checkbox together with the saved file path.
        # The variable trace persists user changes; this restores the value
        # on the next application launch.
        if self._prefs.get("kamal_update_erp_file"):
            self._kamal_update_erp_file.set(True)

        raw_yarn_str = self._prefs.get("kamal_raw_yarn_path")
        if raw_yarn_str and Path(raw_yarn_str).is_file():
            self._raw_yarn_path = Path(raw_yarn_str)
            self._lbl_raw_yarn.config(text=raw_yarn_str, foreground="black")

        lotti_str = self._prefs.get("kamal_lotti_path")
        if lotti_str and Path(lotti_str).is_file():
            self._lotti_path = Path(lotti_str)
            self._lbl_lotti.config(text=lotti_str, foreground="black")

        cache = load_lotti_cache()
        cache_path = cache.get("source_path", "")
        if not self._lotti_path and cache_path and Path(cache_path).is_file():
            self._lotti_path = Path(cache_path)
            self._lbl_lotti.config(text=str(self._lotti_path), foreground="black")

        output_str = self._prefs.get("kamal_output_path")
        if output_str and Path(output_str).parent.is_dir():
            self._output_path = Path(output_str)
            self._lbl_output.config(text=output_str, foreground="black")

        erp_dir_str = self._prefs.get("kamal_erp_export_dir")
        if erp_dir_str and Path(erp_dir_str).is_dir():
            self._kamal_erp_export_dir = Path(erp_dir_str)
            self._lbl_erp_dir.config(text=erp_dir_str, foreground="black")

    def sync_shared_dfm(self):
        """Reflect the DFM file uploaded on Data Elvy/Situazione, if any."""
        cache = load_dfm_cache()
        source_path = cache.get("source_path", "")
        if source_path and source_path != self._shared_dfm_path:
            self._shared_dfm_path = source_path
        if self._shared_dfm_path:
            self._lbl_dfm.config(
                text=f"✅ Using: {Path(self._shared_dfm_path).name}", foreground="black"
            )
        else:
            self._lbl_dfm.config(
                text="Upload a DFM file on the Data Elvy tab — it's shared automatically.",
                foreground="grey",
            )

    def sync_shared_magazino(self):
        """Reflect the shared Magazino selection from Magazino Filato or Purchase Orders."""
        cache = load_magazino_cache()
        source_path = cache.get("source_path", "")
        if not source_path or not Path(source_path).is_file():
            return
        if self._raw_yarn_path is not None and str(self._raw_yarn_path) == str(source_path):
            return
        self._raw_yarn_path = Path(source_path)
        self._lbl_raw_yarn.config(text=str(self._raw_yarn_path), foreground="black")
        self._save_prefs(kamal_raw_yarn_path=str(self._raw_yarn_path), kamal_last_dir=str(self._raw_yarn_path.parent))

    def sync_shared_lotti(self):
        """Reuse the last selected LOTTI file if it was uploaded elsewhere."""
        cache = load_lotti_cache()
        source_path = cache.get("source_path", "")
        if not source_path or not Path(source_path).is_file():
            return
        if self._lotti_path is not None and str(self._lotti_path) == str(source_path):
            return
        self._lotti_path = Path(source_path)
        self._lbl_lotti.config(text=str(self._lotti_path), foreground="black")
        self._save_prefs(kamal_lotti_path=str(self._lotti_path), kamal_last_dir=str(self._lotti_path.parent))

    def _on_convert(self):
        if not self._pdf_paths:
            messagebox.showwarning("Missing input", "Select at least one Kamal order PDF first.")
            return
        if not self._output_path:
            messagebox.showwarning("Missing output", "Select an output file first.")
            return
        if not self._shared_dfm_path:
            messagebox.showwarning(
                "Missing DFM",
                "No DFM reference is loaded yet. Upload one on the Data Elvy tab first — "
                "it's needed to match each colour's Articolo.",
            )
            return

        self._btn_convert.config(state="disabled")
        self._lbl_status.config(text="Converting…", foreground="grey")
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _set_status(self, text: str):
        """Thread-safe status label update — safe to call from the background conversion thread."""
        self.after(0, lambda: self._lbl_status.config(text=text, foreground="grey"))

    def _run_conversion(self):
        errors: list[str] = []
        all_rows = []
        self._set_status("Reading PDF(s)…")
        try:
            for pdf_path in self._pdf_paths:
                all_rows.extend(KamalParser(pdf_path).parse())
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Error parsing PDF(s): {exc}")

        self._set_status("Loading DFM reference (C170)…")
        dfm_c170_entries = load_dfm_entries_by_prefix(KAMAL_ARTICLE_PREFIX)
        if not dfm_c170_entries:
            errors.append(
                f"No {KAMAL_ARTICLE_PREFIX} entries found in the shared DFM reference — "
                "Articolo matching will be blank for every row."
            )

        magazino_summary = None
        codes_map = None
        if self._raw_yarn_path is not None:
            self._set_status("Loading raw yarn stock (Magazino)…")
            try:
                magazino_df, magazino_errors = magazino_logic.load_magazino(
                    str(self._raw_yarn_path), articolo_prefix="G170"
                )
                if magazino_errors or magazino_df is None or magazino_df.empty:
                    errors.append(
                        f"Raw yarn file could not be read: "
                        f"{'; '.join(magazino_errors) if magazino_errors else 'empty after filtering'}"
                    )
                else:
                    magazino_summary = magazino_logic.summarize_by_partita(magazino_df)
            except Exception as exc:  # noqa: BLE001
                errors.append(f"Error loading raw yarn file: {exc}")
            codes_map = {e["articolo"]: e.get("titolo", "") for e in dfm_c170_entries if e.get("titolo")}

        lotti_summary = None
        if self._lotti_path is not None:
            self._set_status("Loading LOTTI reference…")
            try:
                lotti_df, lotti_errors = lotti_logic.load_lotti(str(self._lotti_path), articolo_prefix=lotti_logic.RAW_ARTICOLO_PREFIXES)
                if lotti_errors or lotti_df is None or lotti_df.empty:
                    errors.append(
                        f"LOTTI file could not be read: "
                        f"{'; '.join(lotti_errors) if lotti_errors else 'empty after filtering'}"
                    )
                else:
                    lotti_summary = lotti_logic.summarize_by_partita(lotti_df)
            except Exception as exc:  # noqa: BLE001
                errors.append(f"Error loading LOTTI file: {exc}")

        self._set_status("Matching raw yarn and exporting…")
        if all_rows:
            try:
                KamalExcelExporter(
                    all_rows, self._output_path,
                    dfm_c170_entries=dfm_c170_entries,
                    magazino_summary=magazino_summary,
                    codes_map=codes_map,
                    lotti_summary=lotti_summary,
                ).export()
            except Exception as exc:  # noqa: BLE001
                errors.append(f"Error exporting: {exc}")

            if self._kamal_update_erp_file.get():
                if self._kamal_erp_export_dir is None:
                    errors.append("ERP file extraction was enabled but no folder is selected.")
                else:
                    ordini_path = self._kamal_erp_export_dir / "EXCEL PER ORDINE VENDITA EGITTO.xlsx"
                    filato_path = self._kamal_erp_export_dir / "Filato x Tinturia.xlsx"
                    try:
                        ordini_rows = build_ordine_kamal_rows(all_rows, dfm_c170_entries)
                        assign_ordine_kamal_machines(ordini_rows)
                        if lotti_summary is not None and not lotti_summary.empty:
                            match_by_lotto(ordini_rows, lotti_summary)
                        if magazino_summary is not None and not magazino_summary.empty:
                            match_raw_yarn(ordini_rows, magazino_summary, codes_map, quantity_attr="peso_kg")
                        n = export_ordini_full(ordini_path, ordini_rows)
                        logger.info("Ordine Kamal: extracted ERP file %s (%d rows)", ordini_path.name, n)
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"Error extracting {ordini_path.name}: {exc}")

                    try:
                        matches = read_filato_tinturia_sheet(self._output_path)
                        n2 = export_filato_full(filato_path, matches)
                        logger.info("Filato x Tinturia: extracted %s (%d rows)", filato_path.name, n2)
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"Error extracting {filato_path.name}: {exc}")

        self.after(0, self._on_conversion_done, errors, len(all_rows))

    def _on_conversion_done(self, errors: list[str], row_count: int):
        self._btn_convert.config(state="normal")
        if errors:
            self._lbl_status.config(text="Completed with errors", foreground="#b00020")
            messagebox.showerror("Errors", "\n".join(errors))
        else:
            self._lbl_status.config(text=f"✅ Done — {row_count} row(s)", foreground="green")
            logger.info("Ordine Kamal: exported %d row(s) to %s", row_count, self._output_path)
            messagebox.showinfo("Completed", f"Export completed successfully:\n{self._output_path}")
