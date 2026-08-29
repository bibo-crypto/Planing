import unittest
from pathlib import Path
import tempfile

from abbina_calculator import _smallest_fitting_machine
from abbina_suggestions import titles_compatible
from biglietti_exporter import _machine_for_count
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


if __name__ == "__main__":
    unittest.main()
