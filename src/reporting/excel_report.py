from datetime import datetime
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

        workbook.save(report_path)
        return report_path

    @staticmethod
    def _write_summary(sheet, result: ValidationResult, source_file: Path) -> None:
        generated_at = datetime.now()

        sheet.append(["מדד", "ערך"])
        sheet.append(["שורות שנבדקו", result.summary.total_rows])
        sheet.append(["שורות תקינות", result.summary.valid_rows])
        sheet.append(["שורות שגויות", result.summary.invalid_rows])
        sheet.append(["הקובץ תקין", result.summary.is_valid])
        if result.source_files:
            source_label = result.source_files[0] if len(result.source_files) == 1 else f"{result.source_files[0]} ועוד {len(result.source_files) - 1}"
        else:
            source_label = source_file.name
        sheet.append(["קובץ מקור", source_label])
        sheet.append(["פרופיל בדיקה", result.detected_profile or "GENERIC"])
        sheet.append(["מספר קבצים שנבחרו", max(len(result.source_files), 1)])
        sheet.append(["תאריך הפקה", generated_at.strftime("%Y-%m-%d")])
        sheet.append(["שעת הפקה", generated_at.strftime("%H:%M:%S")])

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
    def _safe_excel_value(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, list):
            return " | ".join(str(item) for item in value)
        if isinstance(value, dict):
            return " | ".join(f"{key}={item}" for key, item in value.items())
        return value
