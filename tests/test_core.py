import tempfile
import unittest
from pathlib import Path

from app import extract_employee_id, unique_output_path


class CoreWorkflowTests(unittest.TestCase):
    def test_extracts_and_normalizes_employee_id(self):
        self.assertEqual(extract_employee_id("Employee ID: emp12345"), "EMP12345")

    def test_returns_none_when_id_is_missing(self):
        self.assertIsNone(extract_employee_id("No identifier on this page"))

    def test_unique_path_does_not_overwrite_existing_file(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "EMP12345.pdf").touch()
            self.assertEqual(
                unique_output_path(root, "EMP12345.pdf").name,
                "EMP12345_1.pdf",
            )


if __name__ == "__main__":
    unittest.main()


