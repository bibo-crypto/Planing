import unittest
from pathlib import Path
import tempfile
import pandas as pd

from abbina_calculator import _smallest_fitting_machine
from abbina_suggestions import titles_compatible
from biglietti_exporter import _machine_for_count
from biglietti_exporter import _get, _read_sheet_rows
from master_import import find_files_in_directory
from situazione_logic import compute_delivery_date, compute_delivery_dates
from ui.tabs.overview_tab import _format_display_dates
from ui.tabs.overview_tab import _write_typed_excel_table
import openpyxl
import path_manager


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
        from ui.tabs.overview_tab import OverviewTab
        self.assertTrue(hasattr(OverviewTab, "_refresh_shared_tabs"))

    def test_canonical_ui_tab_imports(self):
        from ui.tabs.biglietti_tab import BigliettiTab
        from ui.tabs.kamal_tab import KamalTab
        from ui.tabs.magazino_filato_tab import MagazinoFilatoTab
        from ui.tabs.situazione_settimana_tab import SettimanaTab
        from ui.tabs.situazione_tab import SituazioneTab

        self.assertTrue(all((BigliettiTab, KamalTab, MagazinoFilatoTab, SettimanaTab, SituazioneTab)))

    def test_canonical_logic_imports(self):
        from logic.lotti import summarize_by_partita
        from logic.prezzi import build_price_lookup, load_prezzi
        from logic.magazino import summarize_by_partita as magazino_summary
        from logic.situazione_settimana import summarize as weekly_summary
        from prezzi_logic import build_price_lookup as legacy_lookup
        from lotti_logic import summarize_by_partita as legacy_lotti_summary
        from situazione_settimana_logic import summarize as legacy_weekly_summary
        from magazino_logic import summarize_by_partita as legacy_magazino_summary

        self.assertIs(build_price_lookup, legacy_lookup)
        self.assertIs(summarize_by_partita, legacy_lotti_summary)
        self.assertIs(weekly_summary, legacy_weekly_summary)
        self.assertIs(magazino_summary, legacy_magazino_summary)
        self.assertTrue(callable(load_prezzi))

    def test_price_lookup_cache_reuses_same_source(self):
        import biglietti_exporter
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
