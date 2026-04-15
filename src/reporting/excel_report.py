from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook

from src.models.validation_result import ValidationIssue, ValidationResult


class ExcelReportWriter:
    def write(
        self,
        result: ValidationResult,
        source_file: Path,
        output_dir: Path,
    ) -> Path:
        output_dir.mkdir(parents=True, exist_ok=True)
        report_path = output_dir / f"{source_file.stem}_דוח_בדיקות.xlsx"

        workbook = Workbook()
        summary_sheet = workbook.active
        summary_sheet.title = "סיכום"
        self._write_summary(summary_sheet, result, source_file)

        issues_sheet = workbook.create_sheet("שגיאות")
        self._write_issues(issues_sheet, result.issues)

        data_sheet = workbook.create_sheet("נתונים")
        self._write_data(data_sheet, result)

        workbook.save(report_path)
        return report_path

    @staticmethod
    def _write_summary(sheet, result: ValidationResult, source_file: Path) -> None:
        sheet["A1"] = "מדד"
        sheet["B1"] = "ערך"
        sheet["A2"] = "שורות שנבדקו"
        sheet["B2"] = result.summary.total_rows
        sheet["A3"] = "שורות תקינות"
        sheet["B3"] = result.summary.valid_rows
        sheet["A4"] = "שורות שגויות"
        sheet["B4"] = result.summary.invalid_rows
        sheet["A5"] = "הקובץ תקין"
        sheet["B5"] = result.summary.is_valid
        sheet["A6"] = "קובץ מקור"
        if result.source_files:
            sheet["B6"] = result.source_files[0] if len(result.source_files) == 1 else f"{result.source_files[0]} ועוד {len(result.source_files) - 1}"
        else:
            sheet["B6"] = source_file.name
        sheet["A7"] = "פרופיל בדיקה"
        sheet["B7"] = result.detected_profile or "GENERIC"
        sheet["A8"] = "מספר קבצים שנבחרו"
        sheet["B8"] = max(len(result.source_files), 1)

    @staticmethod
    def _write_issues(sheet, issues: Iterable[ValidationIssue]) -> None:
        sheet.append(["מספר שורה", "שם עמודה", "הודעת שגיאה", "קובץ מקור"])
        for issue in issues:
            sheet.append([
                ExcelReportWriter._safe_excel_value(issue.row_number),
                ExcelReportWriter._safe_excel_value(issue.column_name),
                ExcelReportWriter._safe_excel_value(issue.message),
                ExcelReportWriter._safe_excel_value(issue.source_file),
            ])

    @staticmethod
    def _write_data(sheet, result: ValidationResult) -> None:
        all_columns: list[str] = []
        include_source_column = any(row.get("__source_file") for row in result.rows)
        for row in result.rows:
            for column in row.keys():
                if str(column).startswith("__"):
                    continue
                if column not in all_columns:
                    all_columns.append(column)

        headers = all_columns + (["קובץ מקור"] if include_source_column else []) + ["סטטוס בדיקה", "פירוט שגיאות"]
        sheet.append(headers)

        issues_by_row: dict[int, list[str]] = {}
        for issue in result.issues:
            if issue.row_number > 0:
                issues_by_row.setdefault(issue.row_number, []).append(
                    f"{issue.column_name}: {issue.message}"
                )

        for index, row in enumerate(result.rows, start=1):
            row_issues = issues_by_row.get(index, [])
            values = [ExcelReportWriter._safe_excel_value(row.get(column)) for column in all_columns]
            if include_source_column:
                values.append(ExcelReportWriter._safe_excel_value(row.get("__source_file", "")))
            values.append("תקין" if not row_issues else "שגוי")
            values.append(ExcelReportWriter._safe_excel_value(" | ".join(row_issues)))
            sheet.append(values)

    @staticmethod
    def _safe_excel_value(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, list):
            return " | ".join(str(item) for item in value)
        if isinstance(value, dict):
            return " | ".join(f"{key}={item}" for key, item in value.items())
        return value
