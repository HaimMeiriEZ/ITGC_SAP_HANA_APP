from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from src.pipeline import process_file
from src.validators.engine import ValidationEngine


class TestSmoke(unittest.TestCase):
    def test_text_file_is_loaded_and_validated(self) -> None:
        with TemporaryDirectory() as temp_dir:
            file_path = Path(temp_dir) / "users.txt"
            file_path.write_text("user_id;name\n1;Dana\n2;Noam\n", encoding="utf-8")

            result = process_file(file_path, required_columns=["user_id", "name"])

            self.assertEqual(result.summary.total_rows, 2)
            self.assertTrue(result.summary.is_valid)

    def test_excel_file_missing_required_value_is_reported(self) -> None:
        with TemporaryDirectory() as temp_dir:
            file_path = Path(temp_dir) / "users.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["user_id", "name"])
            sheet.append([1, "Dana"])
            sheet.append([2, None])
            workbook.save(file_path)

            result = process_file(file_path, required_columns=["user_id", "name"])

            self.assertEqual(result.summary.total_rows, 2)
            self.assertFalse(result.summary.is_valid)
            self.assertEqual(len(result.issues), 1)
            self.assertEqual(result.issues[0].column_name, "name")

    def test_validation_engine_detects_missing_columns(self) -> None:
        engine = ValidationEngine(required_columns=["user_id", "name", "email"])

        result = engine.validate([
            {"user_id": 1, "name": "Dana"},
        ])

        self.assertFalse(result.summary.is_valid)
        self.assertEqual(result.issues[0].column_name, "email")

    def test_excel_report_is_created_with_summary_and_issues(self) -> None:
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "users.txt"
            output_dir = Path(temp_dir) / "output"
            input_path.write_text(
                "user_id;name;email\n1;Dana;dana@example.com\n2;Noam;\n",
                encoding="utf-8",
            )

            result = process_file(
                input_path,
                required_columns=["user_id", "name", "email"],
                output_dir=output_dir,
            )

            self.assertIsNotNone(result.report_path)
            self.assertTrue(result.report_path.exists())

            workbook = load_workbook(result.report_path)
            self.assertIn("Summary", workbook.sheetnames)
            self.assertIn("Issues", workbook.sheetnames)
            self.assertEqual(workbook["Summary"]["B2"].value, 2)
            self.assertEqual(workbook["Summary"]["B5"].value, False)
            self.assertEqual(workbook["Issues"]["A2"].value, 2)
            self.assertEqual(workbook["Issues"]["B2"].value, "email")


if __name__ == "__main__":
    unittest.main()
