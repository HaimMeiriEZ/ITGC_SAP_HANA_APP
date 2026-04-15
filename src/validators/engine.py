from typing import Any

from src.models.validation_result import ValidationIssue, ValidationResult
from src.validators.spec_rules import build_profile_issues, detect_validation_profile, normalize_name


class ValidationEngine:
    def __init__(self, required_columns: list[str] | None = None) -> None:
        self.required_columns = required_columns or []

    def validate(self, rows: list[dict[str, Any]], source_name: str | None = None) -> ValidationResult:
        issues: list[ValidationIssue] = []
        available_columns = {
            normalize_name(column)
            for row in rows[:1]
            for column in row.keys()
            if not str(column).startswith("__")
        }

        for column in self.required_columns:
            normalized_column = normalize_name(column)
            if normalized_column not in available_columns:
                issues.append(
                    ValidationIssue(
                        row_number=0,
                        column_name=column,
                        message="עמודת חובה חסרה",
                    )
                )

        present_required_columns = [
            column for column in self.required_columns if normalize_name(column) in available_columns
        ]

        for row_number, row in enumerate(rows, start=1):
            normalized_row = {
                normalize_name(key): value
                for key, value in row.items()
                if not str(key).startswith("__")
            }
            for column in present_required_columns:
                value = normalized_row.get(normalize_name(column))
                if value is None or (isinstance(value, str) and not value.strip()):
                    issues.append(
                        ValidationIssue(
                            row_number=row_number,
                            column_name=column,
                            message="ערך חובה חסר",
                            source_file=str(row.get("__source_file", "")),
                        )
                    )

        detected_profile = detect_validation_profile(source_name, rows)
        for issue in build_profile_issues(detected_profile, rows):
            if issue.row_number > 0 and issue.row_number <= len(rows):
                issue.source_file = str(rows[issue.row_number - 1].get("__source_file", ""))
            issues.append(issue)

        return ValidationResult(rows=rows, issues=issues, detected_profile=detected_profile)
