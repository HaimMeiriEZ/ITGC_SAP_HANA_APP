from typing import Any

from src.models.validation_result import ValidationIssue, ValidationResult
from src.validators.spec_rules import (
    build_profile_issues,
    detect_validation_profile,
    filter_required_value_columns,
    get_column_aliases,
    matches_column_alias,
    normalize_name,
)


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
        detected_profile = detect_validation_profile(source_name, rows)
        required_value_columns = filter_required_value_columns(detected_profile, self.required_columns)

        for column in required_value_columns:
            if not matches_column_alias(available_columns, column):
                issues.append(
                    ValidationIssue(
                        row_number=0,
                        column_name=column,
                        message="עמודת חובה חסרה",
                    )
                )

        present_required_columns = [
            column for column in required_value_columns if matches_column_alias(available_columns, column)
        ]

        for row_number, row in enumerate(rows, start=1):
            normalized_row = {
                normalize_name(key): value
                for key, value in row.items()
                if not str(key).startswith("__")
            }
            for column in present_required_columns:
                value = None
                for candidate in get_column_aliases(column):
                    if candidate in normalized_row:
                        value = normalized_row.get(candidate)
                        break
                if value is None or (isinstance(value, str) and not value.strip()):
                    issues.append(
                        ValidationIssue(
                            row_number=row_number,
                            column_name=column,
                            message="ערך חובה חסר",
                            source_file=str(row.get("__source_file", "")),
                        )
                    )

        for issue in build_profile_issues(detected_profile, rows):
            if issue.row_number > 0 and issue.row_number <= len(rows):
                issue.source_file = str(rows[issue.row_number - 1].get("__source_file", ""))
            issues.append(issue)

        return ValidationResult(rows=rows, issues=issues, detected_profile=detected_profile)
