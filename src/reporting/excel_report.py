from pathlib import Path
from typing import Iterable

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
        report_path = output_dir / f"{source_file.stem}_validation_report.xlsx"

        workbook = Workbook()
        summary_sheet = workbook.active
        summary_sheet.title = "Summary"
        self._write_summary(summary_sheet, result, source_file)

        issues_sheet = workbook.create_sheet("Issues")
        self._write_issues(issues_sheet, result.issues)

        data_sheet = workbook.create_sheet("Data")
        self._write_data(data_sheet, result)

        workbook.save(report_path)
        return report_path

    @staticmethod
    def _write_summary(sheet, result: ValidationResult, source_file: Path) -> None:
        sheet["A1"] = "Metric"
        sheet["B1"] = "Value"
        sheet["A2"] = "Rows checked"
        sheet["B2"] = result.summary.total_rows
        sheet["A3"] = "Valid rows"
        sheet["B3"] = result.summary.valid_rows
        sheet["A4"] = "Invalid rows"
        sheet["B4"] = result.summary.invalid_rows
        sheet["A5"] = "Overall valid"
        sheet["B5"] = result.summary.is_valid
        sheet["A6"] = "Source file"
        sheet["B6"] = source_file.name

    @staticmethod
    def _write_issues(sheet, issues: Iterable[ValidationIssue]) -> None:
        sheet.append(["row_number", "column_name", "message"])
        for issue in issues:
            sheet.append([issue.row_number, issue.column_name, issue.message])

    @staticmethod
    def _write_data(sheet, result: ValidationResult) -> None:
        all_columns: list[str] = []
        for row in result.rows:
            for column in row.keys():
                if column not in all_columns:
                    all_columns.append(column)

        headers = all_columns + ["validation_status", "issue_messages"]
        sheet.append(headers)

        issues_by_row: dict[int, list[str]] = {}
        for issue in result.issues:
            if issue.row_number > 0:
                issues_by_row.setdefault(issue.row_number, []).append(
                    f"{issue.column_name}: {issue.message}"
                )

        for index, row in enumerate(result.rows, start=1):
            row_issues = issues_by_row.get(index, [])
            values = [row.get(column) for column in all_columns]
            values.append("valid" if not row_issues else "invalid")
            values.append(" | ".join(row_issues))
            sheet.append(values)
