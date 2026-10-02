"""Tkinter workflow for the Elvy Invoice tab."""

from __future__ import annotations

import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utility.utils import find_pdfs, logger, make_output_path


class ElvyInvoiceWorkflowMixin:
    def _build_elvy_invoice_tab(self, parent: ttk.Frame) -> None:
        """
        Converts Elvy's raw-yarn "Invoice to Delta" PDFs (Proforma invoice +
        Packing list, 2 pages). Extracts Inv No / Date, one row per Pos
        (Yarn Code, NM, Ne, Yarn Type, LOT, Price/Net/Gross/Value), looks
        up each Yarn Code's Articolo Delta from the same Elvy mapping used
        on the Purchase Orders tab, and reads each Pos's total No. of
        cones from the packing-list page into a Rocche column.
        """
        parent.columnconfigure(0, weight=1)

        info = ttk.Label(
            parent,
            text="Elvy-specific data. Converts Elvy's raw-yarn \"Invoice to Delta\" "
                 "PDFs (Proforma invoice + Packing list). Each row's Yarn Code is "
                 "looked up against the same Elvy mapping from the Elvy tab to fill "
                 "in Articolo Delta, and Rocche is read from the packing-list page's "
                 "total No. of cones per Pos.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        info.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=8)
        sel_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 4))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._einv_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_einv_select_pdf, width=20
        )
        self._einv_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._einv_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_einv_select_folder, width=20
        )
        self._einv_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._einv_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_einv_select_output, width=20
        )
        self._einv_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._einv_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)

        self._einv_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)

        self._einv_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._einv_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4)

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(8, 4))
        opt_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=2)
        opt_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            opt_frame,
            text="One Excel file per invoice  (otherwise merge all into one workbook)",
            variable=self._einv_one_per_file,
        ).grid(row=0, column=0, sticky="w")

        self._einv_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_einv_convert,
            style="Accent.TButton",
            width=14,
        )
        self._einv_btn_convert.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(8, 4))
        prog_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=2)
        prog_frame.columnconfigure(0, weight=1)

        self._einv_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._einv_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._einv_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._einv_lbl_status.grid(row=1, column=0, sticky="ew")
    def _on_einv_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select an Elvy Invoice PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("einv_last_dir") or None,
        )
        if path:
            self._einv_pdf_path = Path(path)
            self._einv_folder_path = None
            self._einv_lbl_pdf_path.config(text=str(self._einv_pdf_path), foreground="black")
            self._einv_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(einv_pdf_path=str(self._einv_pdf_path), einv_folder_path=None,
                              einv_last_dir=str(self._einv_pdf_path.parent))
    def _on_einv_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of Elvy Invoice PDFs",
            initialdir=self._prefs.get("einv_last_dir") or None,
        )
        if path:
            self._einv_folder_path = Path(path)
            self._einv_pdf_path = None
            self._einv_lbl_folder_path.config(text=str(self._einv_folder_path), foreground="black")
            self._einv_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(einv_folder_path=str(self._einv_folder_path), einv_pdf_path=None,
                              einv_last_dir=str(self._einv_folder_path))
    def _on_einv_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("einv_output_dir") or None,
        )
        if path:
            self._einv_output_dir = Path(path)
            self._einv_lbl_output_path.config(text=str(self._einv_output_dir), foreground="black")
            self._save_prefs(einv_output_dir=str(self._einv_output_dir))
    def _on_einv_convert(self) -> None:
        input_source: Path | None = self._einv_pdf_path or self._einv_folder_path
        if input_source is None:
            messagebox.showwarning("No Input", "Please select a PDF file or a folder first.")
            return

        if self._einv_output_dir is None:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return

        pdf_list = find_pdfs(input_source)
        if not pdf_list:
            messagebox.showwarning("No PDFs Found", f"No PDF files found in:\n{input_source}")
            return

        self._set_einv_ui_enabled(False)
        self._einv_progress["value"] = 0
        self._einv_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_einv_conversion,
            args=(pdf_list, self._einv_output_dir, self._einv_one_per_file.get()),
            daemon=True,
        )
        thread.start()
    def _run_einv_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
    ) -> None:
        from exporters.elvy_invoice_exporter import ElvyInvoiceExporter
        from parsers.elvy_invoice_parser import ElvyInvoiceParser, ElvyInvoiceRow
        from parsers.elvy_mapping import load_elvy_mapping, lookup_articolo_delta
        errors: list[str] = []
        merged_rows: list[ElvyInvoiceRow] = []
        total = len(pdf_list)
        elvy_mapping = load_elvy_mapping()

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_einv_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows = ElvyInvoiceParser(pdf_path).parse()
                for row in rows:
                    row.articolo_delta = lookup_articolo_delta(row.yarn_code, elvy_mapping)

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    ElvyInvoiceExporter(rows, out_path).export()
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_einv_progress()

        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "Invoice_Elvy.xlsx"
                ElvyInvoiceExporter(merged_rows, out_path).export()
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        self.after(0, self._on_einv_conversion_done, errors, total)
    def _on_einv_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_einv_ui_enabled(True)

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_einv_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} Elvy Invoice PDF file(s).\n\n"
                f"Output saved to:\n{self._einv_output_dir}",
            )
            self._set_einv_status(f"Done — {total} file(s) converted successfully.")
    def _set_einv_status(self, text: str) -> None:
        self.after(0, self._einv_lbl_status.config, {"text": text})
    def _advance_einv_progress(self) -> None:
        def _inc() -> None:
            self._einv_progress["value"] += 1
        self.after(0, _inc)
    def _set_einv_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._einv_btn_pdf,
                self._einv_btn_folder,
                self._einv_btn_output,
                self._einv_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)
