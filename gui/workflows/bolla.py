"""Tkinter workflow for the Bolla tab."""

from __future__ import annotations

import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utility.utils import find_pdfs, logger, make_output_path


class BollaWorkflowMixin:
    def _build_bolla_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        info = ttk.Label(
            parent,
            text="Extracts Bolla No / Del, and every Pallet row (Scatola, Disposizione, "
                 "Articolo, Descrizione, Colore, Partita, Rocche, KgNetto, KgLordo, "
                 "Famiglia) plus the Totale summary line.",
            foreground="grey", anchor="w", wraplength=720, justify="left",
        )
        info.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))

        # ── Input ────────────────────────────────────────────────────
        sel_frame = ttk.LabelFrame(parent, text="Input", padding=8)
        sel_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 4))
        sel_frame.columnconfigure((0, 1, 2), weight=1)

        self._bolla_btn_pdf = ttk.Button(
            sel_frame, text="📄 Select PDF", command=self._on_bolla_select_pdf, width=20
        )
        self._bolla_btn_pdf.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._bolla_btn_folder = ttk.Button(
            sel_frame, text="📁 Select Folder", command=self._on_bolla_select_folder, width=20
        )
        self._bolla_btn_folder.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._bolla_btn_output = ttk.Button(
            sel_frame, text="💾 Output Folder", command=self._on_bolla_select_output, width=20
        )
        self._bolla_btn_output.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        self._bolla_lbl_pdf_path = ttk.Label(
            sel_frame, text="No PDF selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_pdf_path.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)

        self._bolla_lbl_folder_path = ttk.Label(
            sel_frame, text="No folder selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_folder_path.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)

        self._bolla_lbl_output_path = ttk.Label(
            sel_frame, text="No output folder selected", foreground="grey", anchor="w"
        )
        self._bolla_lbl_output_path.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4)

        # ── Options + Convert ────────────────────────────────────────
        opt_frame = ttk.Frame(parent, padding=(8, 4))
        opt_frame.grid(row=2, column=0, sticky="ew", padx=4, pady=2)
        opt_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            opt_frame,
            text="One Excel file per Bolla  (otherwise merge all into one workbook)",
            variable=self._bolla_one_per_file,
        ).grid(row=0, column=0, sticky="w")

        self._bolla_btn_convert = ttk.Button(
            opt_frame,
            text="▶  Convert",
            command=self._on_bolla_convert,
            style="Accent.TButton",
            width=14,
        )
        self._bolla_btn_convert.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # ── Progress + status ────────────────────────────────────────
        prog_frame = ttk.Frame(parent, padding=(8, 4))
        prog_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=2)
        prog_frame.columnconfigure(0, weight=1)

        self._bolla_progress = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self._bolla_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._bolla_lbl_status = ttk.Label(prog_frame, text="Ready", anchor="w")
        self._bolla_lbl_status.grid(row=1, column=0, sticky="ew")
    def _on_bolla_select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a Bolla PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self._prefs.get("bolla_last_dir") or None,
        )
        if path:
            self._bolla_pdf_path = Path(path)
            self._bolla_folder_path = None
            self._bolla_lbl_pdf_path.config(text=str(self._bolla_pdf_path), foreground="black")
            self._bolla_lbl_folder_path.config(text="No folder selected", foreground="grey")
            self._save_prefs(bolla_pdf_path=str(self._bolla_pdf_path), bolla_folder_path=None,
                              bolla_last_dir=str(self._bolla_pdf_path.parent))
    def _on_bolla_select_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select a folder of Bolla PDFs",
            initialdir=self._prefs.get("bolla_last_dir") or None,
        )
        if path:
            self._bolla_folder_path = Path(path)
            self._bolla_pdf_path = None
            self._bolla_lbl_folder_path.config(text=str(self._bolla_folder_path), foreground="black")
            self._bolla_lbl_pdf_path.config(text="No PDF selected", foreground="grey")
            self._save_prefs(bolla_folder_path=str(self._bolla_folder_path), bolla_pdf_path=None,
                              bolla_last_dir=str(self._bolla_folder_path))
    def _on_bolla_select_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._prefs.get("bolla_output_dir") or None,
        )
        if path:
            self._bolla_output_dir = Path(path)
            self._bolla_lbl_output_path.config(text=str(self._bolla_output_dir), foreground="black")
            self._save_prefs(bolla_output_dir=str(self._bolla_output_dir))
    def _on_bolla_convert(self) -> None:
        input_source: Path | None = self._bolla_pdf_path or self._bolla_folder_path
        if input_source is None:
            messagebox.showwarning(
                "No Input", "Please select a PDF file or a folder first."
            )
            return

        if self._bolla_output_dir is None:
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

        self._set_bolla_ui_enabled(False)
        self._bolla_progress["value"] = 0
        self._bolla_progress["maximum"] = len(pdf_list)

        thread = threading.Thread(
            target=self._run_bolla_conversion,
            args=(pdf_list, self._bolla_output_dir, self._bolla_one_per_file.get()),
            daemon=True,
        )
        thread.start()
    def _run_bolla_conversion(
        self,
        pdf_list: list[Path],
        output_dir: Path,
        one_per_file: bool,
    ) -> None:
        from exporters.bolla_exporter import BollaExporter
        from parsers.bolla_parser import BollaParser, BollaRow, BollaTotals
        errors: list[str] = []
        merged_rows: list[BollaRow] = []
        merged_totals: list[BollaTotals] = []
        total = len(pdf_list)

        for idx, pdf_path in enumerate(pdf_list, start=1):
            self._set_bolla_status(f"Processing {idx}/{total}: {pdf_path.name} …")
            try:
                rows, totals = BollaParser(pdf_path).parse()

                if one_per_file:
                    out_path = make_output_path(pdf_path, output_dir)
                    BollaExporter(rows, [totals] if totals else [], out_path).export()
                    logger.info("Saved: %s", out_path.name)
                else:
                    merged_rows.extend(rows)
                    if totals:
                        merged_totals.append(totals)

            except Exception as exc:  # noqa: BLE001
                msg = f"Error processing {pdf_path.name}: {exc}"
                logger.error(msg)
                errors.append(msg)

            self._advance_bolla_progress()

        if not one_per_file and merged_rows:
            try:
                out_path = output_dir / "Bolla_Med.xlsx"
                BollaExporter(merged_rows, merged_totals, out_path).export()
                logger.info("Merged export saved: %s", out_path.name)
            except Exception as exc:  # noqa: BLE001
                msg = f"Error saving merged Excel: {exc}"
                logger.error(msg)
                errors.append(msg)

        self.after(0, self._on_bolla_conversion_done, errors, total)
    def _on_bolla_conversion_done(self, errors: list[str], total: int) -> None:
        self._set_bolla_ui_enabled(True)

        if errors:
            summary = "\n".join(errors)
            messagebox.showerror(
                "Conversion Finished With Errors",
                f"Processed {total} file(s).\n\n"
                f"{len(errors)} error(s) occurred:\n\n{summary}",
            )
            self._set_bolla_status(f"Done — {len(errors)} error(s). Check the log.")
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"✅  Successfully converted {total} Bolla PDF file(s).\n\n"
                f"Output saved to:\n{self._bolla_output_dir}",
            )
            self._set_bolla_status(f"Done — {total} file(s) converted successfully.")
    def _set_bolla_status(self, text: str) -> None:
        self.after(0, self._bolla_lbl_status.config, {"text": text})
    def _advance_bolla_progress(self) -> None:
        def _inc() -> None:
            self._bolla_progress["value"] += 1
        self.after(0, _inc)
    def _set_bolla_ui_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"

        def _apply() -> None:
            for widget in (
                self._bolla_btn_pdf,
                self._bolla_btn_folder,
                self._bolla_btn_output,
                self._bolla_btn_convert,
            ):
                widget.config(state=state)
        self.after(0, _apply)
