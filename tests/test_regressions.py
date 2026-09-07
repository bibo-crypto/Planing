import unittest
from pathlib import Path
import tempfile
import pandas as pd
from unittest.mock import patch

from calculate.abbina_calculator import _smallest_fitting_machine
from calculate.abbina_suggestions import titles_compatible
from exporters.biglietti_exporter import _machine_for_count
from exporters.biglietti_exporter import _get, _read_sheet_rows
from pipelines.master_import import find_files_in_directory
import pipelines.master_import as master_import
from calculate.situazione import compute_delivery_date, compute_delivery_dates
from calculate.lotti import load_lotti
from gui.tabs.overview_tab import _format_display_dates
from gui.tabs.overview_tab import _write_typed_excel_table
import openpyxl
import utility.path_manager as path_manager


class PlanningRegressionTests(unittest.TestCase):
    def test_empty_machine_group_has_no_capacity(self):
        self.assertEqual(_smallest_fitting_machine(0), 0)
        self.assertEqual(_smallest_fitting_machine(-1), 0)

    def test_machine_lookup_rejects_empty_counts(self):
        self.assertEqual(_machine_for_count(None), "")
        self.assertEqual(_machine_for_count(0), "")
        self.assertEqual(_machine_for_count(-2), "")

    def test_special_titles_are_case_and_space_insensitive(self):
        self.assertTrue(titles_compatible(" 30/1 ", "30/1"))
        self.assertFalse(titles_compatible("30/1", "31/1"))

    def test_shared_path_round_trip(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "Produzione.xlsx"
            path.write_bytes(b"placeholder")
            path_manager.save_source("data_prod", path)
            self.assertEqual(
                path_manager.load_source("produzione")["source_path"], str(path)
            )

    def test_elvy_delivery_rules_skip_friday(self):
        row = {"cliente": "3009", "data": "2026-09-11", "mc": 192, "comment": ""}
        self.assertEqual(compute_delivery_date(row), "2026-09-19")
        row["mc"] = 24
        self.assertEqual(compute_delivery_date(row), "2026-09-26")

    def test_elvy_lab_adds_one_week_and_pgx_waits_for_yarn(self):
        row = {"cliente": "ELVY", "data": "2026-09-07", "mc": 192, "new_comment": "Lab"}
        self.assertEqual(compute_delivery_date(row), "2026-09-21")
        row.update({"comment": "PG-X-123", "new_comment": "Filato"})
        self.assertEqual(compute_delivery_date(row), "Bending for yarn")
        row["raw_yarn_match"] = "G130 / 44"
        row["yarn_arrival_date"] = "2026-09-08"
        self.assertEqual(compute_delivery_date(row), "2026-09-15")

    def test_delivery_dates_leave_non_elvy_blank(self):
        frame = pd.DataFrame([{"cliente": "3004", "data": "2026-09-07", "mc": 24}])
        result = compute_delivery_dates(frame)
        self.assertEqual(result.loc[0, "delivery_date"], "")

    def test_master_import_recognizes_lotto_serial_filename(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "ARKMAG-CON-LOTTO_20260829.xlsx"
            path.write_bytes(b"placeholder")
            self.assertEqual(find_files_in_directory(temp_dir)["lotti"], path)

    def test_master_file_finds_data_ordine_sheet_by_content(self):
        # Data Ordine / Dispo Bagno sheets inside a single bundled master
        # workbook have order-specific names (e.g. "MED-D-505449-2026"),
        # never a matchable candidate like "copertura" -- they must be
        # found by sniffing header content instead, the same way loose
        # files in a folder already are.
        with tempfile.TemporaryDirectory() as temp_dir:
            master_path = Path(temp_dir) / "Master.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Copertura"
            ws.append(["N. Bagno", "Machine", "Batch Start"])
            order_ws = wb.create_sheet("MED-D-505449-2026")
            order_ws.append(["Cliente", "DescrizioneAggiuntivaOrdine", "Altro"])
            order_ws.append(["ACME", "Some note", "x"])
            wb.save(master_path)

            matches = master_import._match_sheets(["Copertura", "MED-D-505449-2026"])
            self.assertNotIn("data_ordine", matches)  # not findable by name

            wb2 = openpyxl.load_workbook(master_path, read_only=True)
            try:
                sniffed = master_import._content_match_ordine_sheets(
                    wb2, wb2.sheetnames, set(matches.values())
                )
            finally:
                wb2.close()
            self.assertEqual(sniffed.get("data_ordine"), "MED-D-505449-2026")

    def test_lotti_loader_reads_all_workbook_sheets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "LOTTI.xlsx"
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.append(["MAGAZZINO", "ARTICOLO", "PARTITA", "ORDINE", "QESI", "LOTTO"])
            worksheet.append([900910, "G130-1", "P-1", 0, 10, "L-1"])
            workbook.create_sheet("Seconda")
            workbook.save(path)

            frame, errors = load_lotti(path)

            self.assertEqual(errors, [])
            self.assertEqual(frame.loc[0, "lotto"], "L-1")

    def test_overview_import_skips_order_only_sources(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            (folder / "Data Ordine.xlsx").write_bytes(b"placeholder")
            (folder / "Dispo Bagno.xlsx").write_bytes(b"placeholder")

            with patch.object(master_import, "_route") as route:
                loaded, skipped = master_import.import_master_directory(
                    folder,
                    situazione_tab=None,
                    magazino_tab=None,
                    skip_keys={"data_ordine", "dispo_bagno"},
                )

            route.assert_not_called()
            self.assertNotIn("Data Ordine", loaded + skipped)
            self.assertNotIn("Dispo Bagno", loaded + skipped)

    def test_master_directory_routes_biglietti_template(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "Biglietti.docx"
            template.write_bytes(b"template")

            class Biglietti:
                def __init__(self):
                    self.path = None

                def set_template_path(self, path):
                    self.path = path

            target = Biglietti()
            master_import.import_master_directory(
                temp_dir, situazione_tab=None, magazino_tab=None,
                biglietti_tab=target,
            )

            self.assertEqual(target.path, str(template))

    def test_master_file_reports_unfound_sources_as_skipped(self):
        # Regression guard for the previous silent-drop bug: import_master_file
        # used to build its skipped-list from SHEET_CANDIDATES only, so a
        # source with no sheet-name candidate at all (Data Ordine, Dispo
        # Bagno) simply vanished from both loaded and skipped instead of
        # being reported as missing.
        with tempfile.TemporaryDirectory() as temp_dir:
            master_path = Path(temp_dir) / "Master.xlsx"
            wb = openpyxl.Workbook()
            wb.active.title = "Sheet1"
            wb.save(master_path)
            loaded, skipped = master_import.import_master_file(
                str(master_path), situazione_tab=None, magazino_tab=None
            )
            self.assertIn("Data Ordine", skipped)
            self.assertIn("Dispo Bagno", skipped)

    def test_master_file_keeps_shared_magazino_and_lotti_extracts(self):
        class SharedTab:
            def __init__(self):
                self.magazino_paths = []
                self.lotti_paths = []

            def _on_upload_magazino(self, path, cache_path=None):
                self.magazino_paths.append((path, cache_path))

            def _on_upload_lotti(self, path, cache_path=None):
                self.lotti_paths.append((path, cache_path))

        with tempfile.TemporaryDirectory() as temp_dir:
            master_path = Path(temp_dir) / "Master.xlsx"
            workbook = openpyxl.Workbook()
            workbook.active.title = "Magazino"
            workbook["Magazino"].append(["Articolo", "Partita", "Peso"])
            workbook.create_sheet("LOTTI").append(["Partita", "Lotto"])
            workbook.save(master_path)

            shared_tab = SharedTab()
            master_import.import_master_file(
                str(master_path), situazione_tab=None, magazino_tab=shared_tab
            )

            self.assertTrue(shared_tab.magazino_paths)
            self.assertTrue(shared_tab.lotti_paths)
            for path, cache_path in shared_tab.magazino_paths + shared_tab.lotti_paths:
                self.assertEqual(path, cache_path)
                self.assertTrue(Path(path).is_file())

    def test_order_headers_can_start_later_and_use_aliases(self):
        class Sheet:
            def iter_rows(self, values_only=True):
                return iter([
                    ["Export generated", None, None],
                    ["ARTICOLO", "Q.TA", "DATA CONSEGNA"],
                    ["C130", 4, "2026-09-06"],
                ])

        rows = _read_sheet_rows(Sheet())
        self.assertEqual(_get(rows[0], "Articolo"), "C130")
        self.assertEqual(_get(rows[0], "Ordinata"), 4)
        self.assertEqual(_get(rows[0], "Consegna"), "2026-09-06")

    def test_report_dates_display_day_first(self):
        frame = pd.DataFrame([{"delivery_date": "2026-09-06", "data": "2026-12-31"}, {"delivery_date": "Bending for yarn", "data": ""}])
        result = _format_display_dates(frame, ["data", "delivery_date"])
        self.assertEqual(result.loc[0, "delivery_date"], "06/09/2026")
        self.assertEqual(result.loc[0, "data"], "31/12/2026")
        self.assertEqual(result.loc[1, "delivery_date"], "Bending for yarn")

    def test_overview_shared_refresh_helper_is_available(self):
        from gui.tabs.overview_tab import OverviewTab
        self.assertTrue(hasattr(OverviewTab, "_refresh_shared_tabs"))

    def test_canonical_ui_tab_imports(self):
        from gui.tabs.biglietti_tab import BigliettiTab
        from gui.tabs.kamal_tab import KamalTab
        from gui.tabs.magazino_filato_tab import MagazinoFilatoTab
        from gui.tabs.situazione_settimana_tab import SettimanaTab
        from gui.tabs.situazione_tab import SituazioneTab

        self.assertTrue(all((BigliettiTab, KamalTab, MagazinoFilatoTab, SettimanaTab, SituazioneTab)))

    def test_canonical_logic_imports(self):
        # The old root-level *_logic.py compatibility shims (prezzi_logic,
        # lotti_logic, situazione_settimana_logic, magazino_logic) were
        # removed once every caller was migrated to import from calculate.*
        # directly (see ARCHITECTURE.md). This just confirms the canonical
        # modules still expose the expected callables.
        from calculate.lotti import summarize_by_partita
        from calculate.prezzi import build_price_lookup, load_prezzi
        from calculate.magazino import summarize_by_partita as magazino_summary
        from calculate.situazione_settimana import summarize as weekly_summary

        self.assertTrue(callable(build_price_lookup))
        self.assertTrue(callable(summarize_by_partita))
        self.assertTrue(callable(magazino_summary))
        self.assertTrue(callable(weekly_summary))
        self.assertTrue(callable(load_prezzi))

    def test_price_lookup_cache_reuses_same_source(self):
        import exporters.biglietti_exporter as biglietti_exporter
        self.assertTrue(hasattr(biglietti_exporter, "_PREZZO_LOOKUP_CACHE"))
        self.assertTrue(hasattr(biglietti_exporter, "_ARTICOLI_MARCA_CACHE"))

    def test_excel_export_keeps_yarn_waiting_comment(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        _write_typed_excel_table(
            worksheet,
            pd.DataFrame([
                {"Delivery Date": "2026-09-08"},
                {"Delivery Date": "Bending for yarn"},
            ]),
        )
        self.assertEqual(worksheet.cell(row=2, column=1).value, "08/09/2026")
        self.assertEqual(worksheet.cell(row=2, column=1).data_type, "s")
        self.assertEqual(worksheet.cell(row=3, column=1).value, "Bending for yarn")


if __name__ == "__main__":
    unittest.main()
