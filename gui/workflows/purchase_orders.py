"""Tkinter workflow for the Purchase Orders tab."""

from __future__ import annotations

import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utility.magazino_cache import save_magazino_cache
from utility.utils import find_pdfs, logger, make_output_path


class PurchaseOrderWorkflowMixin:
    def _build_po_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=6)
        sel_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._po_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_po_select_pdf, width=16
        )
        self._po_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._po_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_po_select_folder, width=16
        )
        self._po_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._po_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_po_select_output, width=16
        )
        self._po_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._po_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._po_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 2))

        self._po_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 2))

        self._po_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._po_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4, pady=(0, 2))

        abbina_info = ttk.Label(
            parent,
            text="Abbina uses the PDF Machine annotation when available. Rows without it are filled automatically.",
            foreground="grey", anchor="w", wraplength=620, justify="left",
        )
        abbina_info.grid(row=1, column=0, sticky="ew", padx=4, pady=(2, 0))

        # ── Raw yarn (Magazino) matching ────────────────────────────
        raw_yarn_frame = ttk.LabelFrame(parent, text="Raw Yarn Matching", padding=6)
        raw_yarn_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 2))
        raw_yarn_frame.columnconfigure(1, weight=1)

        ttk.Button(
            raw_yarn_frame, text="📦 Select Magazino File…", command=self._on_po_select_raw_yarn, width=20
        ).grid(row=0, column=0, padx=4, pady=3, sticky="w")

        self._po_lbl_raw_yarn = ttk.Label(
            raw_yarn_frame, text="No Magazino file selected", foreground="grey", anchor="w"
        )
        self._po_lbl_raw_yarn.grid(row=0, column=1, sticky="ew", padx=4)

        # ── Extract "EXCEL PER ORDINE VENDITA EGITTO" + "Filato x Tinturia" ──
        erp_frame = ttk.LabelFrame(parent, text="Also Extract ERP Files", padding=6)
        erp_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(3, 2))
        erp_frame.columnconfigure(1, weight=1)

        ttk.Button(
            erp_frame, text="📁 Select ERP Files Folder…", command=self._on_po_select_erp_folder, width=22
        ).grid(row=0, column=0, padx=(0, 6), pady=(2, 4), sticky="w")
        ttk.Checkbutton(
            erp_frame, text="Extract ERP order file", variable=self._po_update_erp_file
        ).grid(row=0, column=1, sticky="w", pady=(2, 4))
        self._po_lbl_erp_dir = ttk.Label(erp_frame, text="No folder selected", foreground="grey", anchor="w")
        self._po_lbl_erp_dir.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))

        ttk.Button(
            erp_frame, text="📁 Select Filato Folder…", command=self._on_po_select_filato_folder, width=22
        ).grid(row=2, column=0, padx=(0, 6), pady=(4, 2), sticky="w")
        ttk.Checkbutton(
            erp_frame, text="Extract Filato x Tinturia.xlsx", variable=self._po_update_filato_file
        ).grid(row=2, column=1, sticky="w", pady=(4, 2))
        self._po_lbl_filato_dir = ttk.Label(erp_frame, text="No folder selected", foreground="grey", anchor="w")
        self._po_lbl_filato_dir.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 2))

        erp_warning = ttk.Label(
            erp_frame,
            text="⚠ Close these files in Excel before converting, or saving will fail.",
            foreground="#8a6d00", anchor="w",
        )
        erp_warning.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2, 0))

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(6, 3))
        opt_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=1)
        opt_frame.columnconfigure(0, weight=1)

        self._po_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_po_convert,
            style="Accent.TButton",
            width=12,
        )
        self._po_btn_convert.grid(row=0, column=0, sticky="w", pady=(0, 1))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(4, 2))
        prog_frame.grid(row=5, column=0, sticky="ew", padx=4, pady=1)
        prog_frame.columnconfigure(0, weight=1)

        self._po_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._po_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._po_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._po_lbl_status.grid(row=1, column=0, sticky="ew")
    def _on_po_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("po_last_dir") or None,
        )
        if path:
            self._po_pdf_path = Path(path)
            self._po_folder_path = None
            self._po_lbl_pdf_path.config(text=str(self._po_pdf_path), foreground="black")
            self._po_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(po_pdf_path=str(self._po_pdf_path), po_folder_path=None,
                              po_last_dir=str(self._po_pdf_path.parent))
    def _on_po_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of PDFs",
            initialdir=self._prefs.get("po_last_dir") or None,
        )
        if path:
            self._po_folder_path = Path(path)
            self._po_pdf_path = None
            self._po_lbl_folder_path.config(text=str(self._po_folder_path), foreground="black")
            self._po_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(po_folder_path=str(self._po_folder_path), po_pdf_path=None,
                              po_last_dir=str(self._po_folder_path))
    def _on_po_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("po_output_dir") or None,
        )
        if path:
            self._po_output_dir = Path(path)
            self._po_lbl_output_path.config(text=str(self._po_output_dir), foreground="black")
            self._save_prefs(po_output_dir=str(self._po_output_dir))
    def _on_po_select_erp_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select ERP file folder",
            initialdir=self._prefs.get("po_erp_export_dir") or None,
        )
        if path:
            self._po_erp_export_dir = Path(path)
            self._po_lbl_erp_dir.config(text=str(self._po_erp_export_dir), foreground="black")
            self._save_prefs(po_erp_export_dir=str(self._po_erp_export_dir))
    def _on_po_select_filato_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select Filato x Tinturia output folder",
            initialdir=self._prefs.get("po_filato_export_dir") or None,
        )
        if path:
            self._po_filato_export_dir = Path(path)
            self._po_lbl_filato_dir.config(text=str(self._po_filato_export_dir), foreground="black")
            self._save_prefs(po_filato_export_dir=str(self._po_filato_export_dir))
    def _on_po_select_raw_yarn(self) -> None:
        path = filedialog.askopenfilename(
            title="Select the Magazino (raw yarn) Excel export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._po_raw_yarn_path = Path(path)
            self._po_lbl_raw_yarn.config(text=str(self._po_raw_yarn_path), foreground="black")
            save_magazino_cache(self._po_raw_yarn_path)
            self._save_prefs(po_raw_yarn_path=str(self._po_raw_yarn_path))
            self._on_shared_cache_changed()
    def _on_po_convert(self) -> None:
        """Validate inputs then kick off conversion in a background thread."""
        input_source: Path | None = self._po_pdf_path or self._po_folder_path
        if input_source is None:
            messagebox.showwarning(
                "No Input", "Please select a PDF file or a folder first."
            )
            return

        if self._po_output_dir is None:
            messagebox.showwarning(
                "No Output Folder", "Please select an output folder first."
            )
            return

        pdf_list = find_pdfs(input_source)
        if not pdf_list:
            messagebox.showwarning(
                "No PDFs Found", f"No PDF files found in:\n{input_source}"
            )
            return

        self._set_po_ui_enabled(False)
        self._po_progress["value"] = 0
        self._po_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_po_conversion,
            args=(
                pdf_list,
                self._po_output_dir,
                False,
                self._po_update_erp_file.get(),
                self._po_update_filato_file.get(),
                self._po_erp_export_dir,
                self._po_filato_export_dir,
                self._po_raw_yarn_path,
            ),
            daemon=True,
        )
        thread.start()
    def _run_po_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
        update_erp_file: bool,
        update_filato_file: bool,
        erp_export_dir: Path | None,
        filato_export_dir: Path | None,
        raw_yarn_path: Path | None = None,
    ) -> None:
        """Run the Purchase Order conversion and optional ERP extracts."""
        from calculate.abbina_calculator import AbbinaCalculator
        from calculate import prezzi as prezzi_logic
        from exporters.excel_exporter import ExcelExporter
        from parsers.dfm_lookup import (
            is_first_time_dyeing,
            load_dfm_cache,
            lookup_dfm_color,
            raw_to_finished_articolo,
        )
        from parsers.elvy_mapping import load_elvy_mapping, lookup_articolo_delta
        from parsers.pdf_parser import OrderRow, PDFParser
        from pipelines.ordini_elvy import (
            build_ordini_elvy_rows,
            export_filato_full,
            export_ordini_full,
            match_raw_yarn,
            read_filato_tinturia_sheet,
        )
        errors: list[str] = []
        merged_rows: list[OrderRow] = []
        all_rows: list[OrderRow] = []
        last_export_path: Path | None = None
        total = len(pdf_list)
        calculator = AbbinaCalculator()
        elvy_mapping = load_elvy_mapping()
        dfm_entries = load_dfm_cache().get("entries", [])
        prezzi_df = getattr(self._prezzi_tab, "prezzi_df", None)
        price_lookup = prezzi_logic.build_price_lookup(prezzi_df)

        magazino_summary = None
        codes_map = None
        if raw_yarn_path is not None:
            try:
                from calculate import magazino as magazino_logic
                magazino_df, magazino_errors = magazino_logic.load_magazino(str(raw_yarn_path))
                if magazino_errors or magazino_df is None or magazino_df.empty:
                    msg = f"Raw yarn file could not be read: {'; '.join(magazino_errors) if magazino_errors else 'empty after filtering'}"
                    logger.error(msg)
                    errors.append(msg)
                else:
                    magazino_summary = magazino_logic.summarize_by_partita(magazino_df)
                    logger.info("Raw yarn stock loaded: %d batches from %s",
                                len(magazino_summary), raw_yarn_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error loading raw yarn file: {exc}"
                logger.error(msg)
                errors.append(msg)
            try:
                # Titolo lookup for the Filato x Tinturia sheet comes from the
                # DFM reference already loaded on the Data Elvy tab (parsed
                # from DESCRIZARTICOLOLI), keyed by Articolo -- no separate
                # upload needed.
                codes_map = {e["articolo"]: e.get("titolo", "") for e in dfm_entries if e.get("titolo")}
            except Exception:  # noqa: BLE001
                codes_map = None

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_po_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows = PDFParser(pdf_path).parse()

                for row in rows:
                    row.articolo_delta = lookup_articolo_delta(row.article_no, elvy_mapping)
                    row.coloredfm, row.cldescr = lookup_dfm_color(
                        row.articolo_delta, row.colour, row.ne, row.dye_type, row.yarn,
                        dfm_entries,
                    )
                    if is_first_time_dyeing(row.articolo_delta, row.coloredfm, dfm_entries):
                        row.check_articolo = "prima volta tint."
                    finished_articolo = raw_to_finished_articolo(row.articolo_delta)
                    if finished_articolo:
                        row.livello, row.prezzo = price_lookup.get(
                            (finished_articolo, row.coloredfm), (None, None)
                        )

                calculator.calculate(rows, only_if_missing=True)
                all_rows.extend(rows)

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    ExcelExporter(rows, out_path, magazino_summary, codes_map).export()
                    last_export_path = out_path
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_po_progress()

        # Final export for merged mode
        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "Ordine_Elvy.xlsx"
                ExcelExporter(merged_rows, out_path, magazino_summary, codes_map).export()
                last_export_path = out_path
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        # Extract both ERP files into their saved folders, if configured --
        # each write replaces whatever was already at that path (see
        # export_ordini_full/export_filato_full) so a shared destination
        # always reflects only the order that was just extracted, never a
        # mix of an old order's rows with a new one's.
        if update_erp_file and all_rows:
            if erp_export_dir is None:
                msg = "ERP file extraction was enabled but no ERP folder is selected."
                logger.error(msg)
                errors.append(msg)
            else:
                ordini_path = erp_export_dir / "EXCEL PER ORDINE VENDITA EGITTO.xlsx"
                try:
                    ordini_rows = build_ordini_elvy_rows(all_rows)
                    if magazino_summary is not None and not magazino_summary.empty:
                        match_raw_yarn(ordini_rows, magazino_summary, codes_map)
                    n = export_ordini_full(ordini_path, ordini_rows)
                    logger.info("Extracted ERP file: %s (%d rows)", ordini_path.name, n)
                except Exception as exc:  # noqa: BLE001
                    msg = f"Error extracting {ordini_path.name}: {exc}"
                    logger.error(msg)
                    errors.append(msg)

        if update_filato_file and all_rows:
            if filato_export_dir is None:
                msg = "Filato x Tinturia extraction was enabled but no Filato folder is selected."
                logger.error(msg)
                errors.append(msg)
            elif last_export_path is not None:
                filato_path = filato_export_dir / "Filato x Tinturia.xlsx"
                try:
                    matches = read_filato_tinturia_sheet(last_export_path)
                    n2 = export_filato_full(filato_path, matches)
                    logger.info("Extracted Filato x Tinturia file: %s (%d rows)", filato_path.name, n2)
                except Exception as exc:  # noqa: BLE001
                    msg = f"Error extracting {filato_path.name}: {exc}"
                    logger.error(msg)
                    errors.append(msg)

        self._po_last_export_path = last_export_path
        self.after(0, self._on_po_conversion_done, errors, total)
    def _on_po_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_po_ui_enabled(True)
        if self._po_last_export_path and self._po_last_export_path.is_file():
            self._save_prefs(po_last_export_path=str(self._po_last_export_path))

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n"
                f"{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_po_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} PDF file(s).\n\n"
                f"Output saved to:\n{self._po_output_dir}",
            )
            self._set_po_status(f"Done — {total} file(s) converted successfully.")
    def _set_po_status(self, text: str) -> None:
        self.after(0, self._po_lbl_status.config, {"text": text})
    def _advance_po_progress(self) -> None:
        def _inc() -> None:
            self._po_progress["value"] += 1
        self.after(0, _inc)
    def _set_po_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._po_btn_pdf,
                self._po_btn_folder,
                self._po_btn_output,
                self._po_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)
