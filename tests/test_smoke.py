from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import unittest

from openpyxl import Workbook, load_workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QSizePolicy

from src.models.validation_result import ValidationIssue, ValidationResult
from src.pipeline import process_file
from src.ui.desktop_app import ValidationDesktopApp, get_qt_app
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
            self.assertIn("סיכום", workbook.sheetnames)
            self.assertIn("שגיאות", workbook.sheetnames)
            self.assertEqual(workbook["סיכום"]["A2"].value, "שורות שנבדקו")
            self.assertEqual(workbook["סיכום"]["B2"].value, 2)
            self.assertEqual(workbook["סיכום"]["A5"].value, "הקובץ תקין")
            self.assertEqual(workbook["סיכום"]["B5"].value, False)
            self.assertEqual(workbook["שגיאות"]["A1"].value, "מספר שורה")
            self.assertEqual(workbook["שגיאות"]["A2"].value, 2)
            self.assertEqual(workbook["שגיאות"]["B2"].value, "email")
            self.assertEqual(workbook["שגיאות"]["C2"].value, "ערך חובה חסר")

    def test_process_file_accepts_multiple_files_for_slot(self) -> None:
        with TemporaryDirectory() as temp_dir:
            first_path = Path(temp_dir) / "usr02_part1.txt"
            second_path = Path(temp_dir) / "usr02_part2.txt"
            first_path.write_text("BNAME;UFLAG;TRDAT;LTIME\nUSER_A;0;20260101;080000\n", encoding="utf-8")
            second_path.write_text("BNAME;UFLAG;TRDAT;LTIME\nUSER_B;0;20260102;090000\n", encoding="utf-8")

            result = process_file(
                [first_path, second_path],
                required_columns=["BNAME", "UFLAG"],
                source_name_override="USR02",
            )

            self.assertEqual(result.summary.total_rows, 2)
            self.assertTrue(result.summary.is_valid)
            self.assertEqual(sorted(result.source_files), sorted([first_path.name, second_path.name]))

    def test_sap_usr02_export_with_metadata_and_hash_is_parsed(self) -> None:
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "usr02.txt"
            output_dir = Path(temp_dir) / "output"
            input_path.write_text(
                "Table:\t\t\tUSR02\n"
                "Displayed Fields:\t\t\t44\tof\t44\n\n"
                "\tMANDT\tBNAME\tUFLAG\tTRDAT\tLTIME\tPWDSALTEDHASH\n"
                "\t100\tAROMI\t0\t17.11.2025\t11:10:44\t{x-isSHA512, 15000}ABCDEF\n",
                encoding="utf-8",
            )

            result = process_file(
                input_path,
                required_columns=["BNAME", "UFLAG", "TRDAT", "LTIME"],
                output_dir=output_dir,
                source_name_override="USR02",
            )

            self.assertEqual(result.summary.total_rows, 1)
            self.assertIsNotNone(result.report_path)
            self.assertTrue(result.report_path.exists())

    def test_sap_e070_export_with_windows_encoding_is_parsed(self) -> None:
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "e070.txt"
            content = (
                "Table:\t\tE070\n"
                "Displayed Fields:\t\t\t10\tof\t\t10\n\n"
                "\tTRKORR\tTRFUNCTION\tTRSTATUS\tTARSYSTEM\tKORRDEV\tAS4USER\tAS4DATE\tAS4TIME\tSTRKORR\tAS4TEXT\n\n"
                "\tFPDK901838\tW\tR\tFPQ\tCUST\tPICCOLOG\t17.01.2025\t13:42:40\t\tDER “Customizing”\n"
            )
            input_path.write_bytes(content.encode("cp1252"))

            result = process_file(
                input_path,
                required_columns=["TRKORR", "AS4USER", "TRFUNCTION"],
                source_name_override="E070",
            )

            self.assertEqual(result.summary.total_rows, 1)
            self.assertTrue(result.summary.is_valid)

    def test_desktop_gui_initializes_with_hebrew_labels(self) -> None:
        qt_app = get_qt_app()
        self.assertIsInstance(qt_app, QApplication)

        window = ValidationDesktopApp()
        try:
            self.assertIn("הרץ בדיקה", window.run_button.text())
            self.assertIn("מקורות קלט לבדיקת SAP HANA APP", window.slots_group.title())
            self.assertTrue(window.slots_group.alignment() & Qt.AlignRight)
            self.assertEqual(window.header_label.layoutDirection(), Qt.RightToLeft)
            self.assertEqual(window.hint_label.layoutDirection(), Qt.RightToLeft)
            self.assertEqual(window.actions_row.layoutDirection(), Qt.RightToLeft)
            self.assertEqual(window.actions_row.sizePolicy().horizontalPolicy(), QSizePolicy.Expanding)
            self.assertTrue(window.header_label.alignment() & Qt.AlignRight)
            self.assertTrue(window.hint_label.alignment() & Qt.AlignRight)
            self.assertFalse(window.required_columns_group.isVisible())
            self.assertFalse(window.summary_group.isVisible())
            self.assertFalse(window.results_group.isVisible())
            self.assertIn("לחיצה כפולה", window.run_log_table.toolTip())
            self.assertEqual(window.run_log_table.columnCount(), 10)
            self.assertEqual(window.run_log_table.horizontalHeaderItem(1).text(), "קבוצת דוחות")
            self.assertEqual(window.run_log_table.horizontalHeaderItem(3).text(), "תאריך הפקה")
            self.assertEqual(window.run_log_table.horizontalHeaderItem(4).text(), "רשומות שנקלטו")
            self.assertEqual(window.run_log_table.horizontalHeaderItem(7).text(), "תיאור שגיאה")
            self.assertEqual(window.run_log_table.horizontalHeaderItem(8).text(), "תאריך בדיקה")
            self.assertEqual(window.run_log_table.horizontalHeaderItem(9).text(), "שעת בדיקה")
            self.assertIn("USR02", window.slot_widgets)
            self.assertIn("extraction_date_edit", window.slot_widgets["USR02"])
            self.assertIn("extraction_date_label", window.slot_widgets["USR02"])
            self.assertTrue(window.slot_widgets["USR02"]["extraction_date_label"].alignment() & Qt.AlignRight)
            self.assertEqual(window.slot_widgets["USR02"]["path_label"].layoutDirection(), Qt.RightToLeft)
            self.assertIn("טרם נבחר קובץ", window.slot_widgets["USR02"]["path_label"].text())
            self.assertIn("AGR_USERS", window.slot_widgets)
            self.assertIn("RSPARAM", window.slot_widgets)
            self.assertIn("טבלאות משתמשים", window.category_run_buttons)
            self.assertIn("טבלאות משתמשים", window.category_sections)
            self.assertEqual(window.category_run_buttons["טבלאות משתמשים"].text(), "הרץ בדיקה")
            self.assertNotEqual(window.category_run_buttons["טבלאות משתמשים"].styleSheet(), "")
        finally:
            window.close()

    def test_slot_controls_are_visibly_rendered(self) -> None:
        qt_app = get_qt_app()
        window = ValidationDesktopApp()
        try:
            window.show()
            qt_app.processEvents()
            self.assertGreater(window.slot_widgets["USR02"]["button"].height(), 20)
            self.assertGreater(window.slot_widgets["USR02"]["path_label"].height(), 20)
            self.assertGreater(window.slot_widgets["USR02"]["extraction_date_edit"].height(), 20)
            self.assertGreater(
                window.slots_scroll.widget().minimumSizeHint().height(),
                window.slots_scroll.viewport().height(),
            )
        finally:
            window.close()

    def test_category_run_button_validates_selected_group(self) -> None:
        with TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            usr02_path = base_dir / "usr02_100.txt"
            adr6_path = base_dir / "adr6.txt"
            usr02_path.write_text(
                "BNAME;UFLAG;TRDAT;LTIME\nUSER_A;0;20260101;080000\n",
                encoding="utf-8",
            )
            adr6_path.write_text(
                "ADDRNUMBER;PERSNUMBER;SMTP_ADDR\n1001;2001;user@example.com\n",
                encoding="utf-8",
            )

            get_qt_app()
            window = ValidationDesktopApp(base_dir=base_dir)
            try:
                window.slot_widgets["USR02"]["selected_paths"] = [str(usr02_path)]
                window.slot_widgets["ADR6_USR21"]["selected_paths"] = [str(adr6_path)]

                with patch("src.ui.desktop_app.QMessageBox.information") as information_mock, patch(
                    "src.ui.desktop_app.QMessageBox.warning"
                ) as warning_mock:
                    window.run_category_validation("טבלאות משתמשים")

                self.assertEqual(window.run_log_table.rowCount(), 2)
                self.assertTrue(window.report_button.isEnabled())
                self.assertTrue(information_mock.called)
                self.assertFalse(warning_mock.called)
            finally:
                window.close()

    def test_category_run_warns_when_required_slot_is_missing(self) -> None:
        with TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            adr6_path = base_dir / "adr6_only.txt"
            adr6_path.write_text(
                "ADDRNUMBER;PERSNUMBER;SMTP_ADDR\n1001;2001;user@example.com\n",
                encoding="utf-8",
            )

            get_qt_app()
            window = ValidationDesktopApp(base_dir=base_dir)
            try:
                window.slot_widgets["ADR6_USR21"]["selected_paths"] = [str(adr6_path)]

                with patch("src.ui.desktop_app.QMessageBox.information") as information_mock, patch(
                    "src.ui.desktop_app.QMessageBox.warning"
                ) as warning_mock:
                    window.run_category_validation("טבלאות משתמשים")

                self.assertEqual(window.run_log_table.rowCount(), 1)
                self.assertTrue(warning_mock.called)
                self.assertTrue(information_mock.called)
            finally:
                window.close()

    def test_run_log_is_recorded_per_file_and_exposes_details(self) -> None:
        window = ValidationDesktopApp()
        try:
            window.slot_widgets["USR02"]["extraction_date_edit"].setText("2026-04-15")
            result = ValidationResult(
                rows=[
                    {"BNAME": "USER_A", "__source_file": "usr02_a.txt"},
                    {"BNAME": "USER_B", "__source_file": "usr02_b.txt"},
                ],
                issues=[
                    ValidationIssue(
                        row_number=1,
                        column_name="BNAME",
                        message="ערך חריג",
                        source_file="usr02_a.txt",
                    )
                ],
                source_files=["usr02_a.txt", "usr02_b.txt"],
            )

            window._append_run_log_entries("USR02", ["C:/temp/usr02_a.txt", "C:/temp/usr02_b.txt"], result)

            self.assertEqual(window.run_log_table.rowCount(), 2)
            self.assertEqual(window.run_log_table.item(0, 0).text(), "USR02")
            self.assertEqual(window.run_log_table.item(0, 1).text(), "טבלאות משתמשים")
            self.assertEqual(window.run_log_table.item(0, 3).text(), "2026-04-15")
            self.assertEqual(window.run_log_table.item(0, 4).text(), "1")
            self.assertEqual(window.run_log_table.item(0, 5).text(), "שגוי")
            self.assertEqual(window.run_log_table.item(1, 5).text(), "תקין")
            self.assertIn("ערך חריג", window.run_log_table.item(0, 7).text())
            self.assertIn("ללא שגיאות", window.run_log_table.item(1, 7).text())
            self.assertRegex(window.run_log_table.item(0, 8).text(), r"\d{4}-\d{2}-\d{2}")
            self.assertRegex(window.run_log_table.item(0, 9).text(), r"\d{2}:\d{2}:\d{2}")
            invalid_details = window._build_log_details(0)
            valid_details = window._build_log_details(1)
            self.assertIn("usr02_a.txt", invalid_details)
            self.assertIn("טבלאות משתמשים", invalid_details)
            self.assertIn("2026-04-15", invalid_details)
            self.assertIn("מספר רשומות שנקלטו: 1", invalid_details)
            self.assertIn("ערך חריג", invalid_details)
            self.assertIn("usr02_b.txt", valid_details)
            self.assertIn("לא נמצאו שגיאות", valid_details)
        finally:
            window.close()

    def test_rtl_formatter_keeps_text_clean_without_control_markers(self) -> None:
        value = ValidationDesktopApp.format_rtl_text("קובץ Excel, users.txt (2026)")

        self.assertEqual(value, "קובץ Excel, users.txt (2026)")
        self.assertNotIn("\u2066", value)
        self.assertNotIn("\u2067", value)
        self.assertNotIn("\u2069", value)

    def test_hana_app_slot_catalog_contains_expected_sources(self) -> None:
        self.assertEqual(ValidationDesktopApp.SLOT_DEFINITIONS["USR02"]["category"], "טבלאות משתמשים")
        self.assertEqual(ValidationDesktopApp.SLOT_DEFINITIONS["AGR_USERS"]["category"], "טבלאות הרשאות כלליות")
        self.assertEqual(ValidationDesktopApp.SLOT_DEFINITIONS["E070"]["category"], "טבלאות שינויים")
        self.assertEqual(ValidationDesktopApp.SLOT_DEFINITIONS["RSPARAM"]["category"], "מדיניות סיסמאות")
        self.assertIn("משתמשים", ValidationDesktopApp.SLOT_DEFINITIONS["USR02"]["description"])

    def test_password_policy_profile_detects_security_gaps(self) -> None:
        rows = [
            {"PROPERTY": "minimal_password_length", "VALUE": "6"},
            {"PROPERTY": "force_first_password_change", "VALUE": "FALSE"},
            {"PROPERTY": "maximum_invalid_connect_attempts", "VALUE": "10"},
        ]

        result = ValidationEngine().validate(rows, source_name="password_policy.csv")
        messages = {issue.message for issue in result.issues}

        self.assertIn("אורך סיסמה מינימלי חייב להיות לפחות 8", messages)
        self.assertIn("חובת החלפת סיסמה ראשונית חייבת להיות פעילה", messages)
        self.assertIn("מספר ניסיונות התחברות שגויים חייב להיות מוגבל", messages)

    def test_audit_policies_profile_requires_active_policy(self) -> None:
        rows = [
            {"AUDIT_POLICY_NAME": "USER_CHANGES", "IS_AUDIT_POLICY_ACTIVE": "FALSE"},
        ]

        result = ValidationEngine().validate(rows, source_name="audit_policies.csv")

        self.assertFalse(result.summary.is_valid)
        self.assertTrue(any(issue.message == "לפחות מדיניות Audit אחת חייבת להיות פעילה" for issue in result.issues))

    def test_users_profile_requires_last_login_field(self) -> None:
        rows = [
            {"USER_NAME": "DANA"},
        ]

        result = ValidationEngine().validate(rows, source_name="users_export.csv")

        self.assertFalse(result.summary.is_valid)
        self.assertTrue(any(issue.message == "נדרשת לפחות אחת מהעמודות: LAST_SUCCESSFUL_CONNECT / LAST_SUCCESSFUL_CONNECT_DATE" for issue in result.issues))

    def test_adr6_usr21_slot_accepts_adr6_only_structure(self) -> None:
        rows = [
            {"ADDRNUMBER": "1001", "PERSNUMBER": "2001", "SMTP_ADDR": "user@example.com"},
        ]

        result = ValidationEngine().validate(rows, source_name="ADR6_USR21")

        self.assertFalse(any(issue.column_name == "BNAME" and issue.message == "עמודת חובה חסרה" for issue in result.issues))

    def test_agr_1251_allows_empty_high_value_in_normal_sap_rows(self) -> None:
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "agr_1251_100.txt"
            input_path.write_text(
                "MANDT;AGR_NAME;COUNTER;OBJECT;AUTH;VARIANT;FIELD;LOW;HIGH\n"
                "100;/AIF/ARC_CREATE;000001;S_TCODE;T_XX93001900;;TCD;SARA;\n"
                "100;/AIF/ARC_CREATE;000002;S_ADMI_FCD;T_XX93001900;;S_ADMI_FCD;;\n",
                encoding="utf-8",
            )

            result = process_file(
                input_path,
                required_columns=["AGR_NAME", "OBJECT", "FIELD", "LOW", "HIGH"],
                source_name_override="AGR_1251",
            )

            range_issues = [issue for issue in result.issues if issue.column_name in {"LOW", "HIGH"}]
            self.assertEqual(range_issues, [])

    def test_stms_accepts_formal_sap_headers(self) -> None:
        rows = [
            {
                "Number": "1",
                "Date": "17.01.25",
                "Time": "13:42:40",
                "Request": "FPDK901838",
                "Clt": "400",
                "Owner": "PICCOLOG",
                "User": "PICCOLOG",
                "Project": "PICCOLOG",
                "Short Text": "DER_Customizing_Finetuning_16012025",
                "RC": "0",
            }
        ]

        result = ValidationEngine().validate(rows, source_name="STMS")

        self.assertFalse(any(issue.message == "עמודת חובה חסרה" for issue in result.issues))
        self.assertFalse(any("אינו תואם למבנה המצופה עבור המשבצת STMS" in issue.message for issue in result.issues))

    def test_stms_blank_rc_value_is_not_treated_as_missing_status(self) -> None:
        rows = [
            {
                "Request": "FPDK901838",
                "RC": "",
                "Owner": "PICCOLOG",
            }
        ]

        result = ValidationEngine(required_columns=["TRKORR", "STATUS"]).validate(rows, source_name="STMS")

        self.assertFalse(any(issue.column_name == "STATUS" for issue in result.issues))

    def test_failed_slot_validation_is_logged_in_run_log(self) -> None:
        get_qt_app()
        window = ValidationDesktopApp()
        try:
            with patch("src.ui.desktop_app.process_file", side_effect=RuntimeError("boom")):
                summary = window._run_slot_validation("E070", ["C:/temp/e070.txt"], show_feedback=False)

            self.assertEqual(summary["status"], "error")
            self.assertEqual(window.run_log_table.rowCount(), 1)
            self.assertEqual(window.run_log_table.item(0, 0).text(), "E070")
            self.assertEqual(window.run_log_table.item(0, 1).text(), "טבלאות שינויים")
            self.assertEqual(window.run_log_table.item(0, 4).text(), "0")
            self.assertEqual(window.run_log_table.item(0, 5).text(), "שגיאה")
            self.assertIn("boom", window.run_log_table.item(0, 7).text())
            self.assertIn("boom", window._build_log_details(0))
        finally:
            window.close()

    def test_usr02_slot_blocks_wrong_rsparam_structure(self) -> None:
        rows = [
            {"PARAMETER": "login/min_password_lng", "VALUE": "8"},
        ]

        result = ValidationEngine().validate(rows, source_name="USR02")

        self.assertFalse(result.summary.is_valid)
        self.assertTrue(any("אינו תואם למבנה המצופה עבור המשבצת USR02" in issue.message for issue in result.issues))


if __name__ == "__main__":
    unittest.main()
