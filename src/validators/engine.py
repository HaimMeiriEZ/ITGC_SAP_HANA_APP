from typing import Any

from src.models.validation_result import ValidationIssue, ValidationResult


class ValidationEngine:
    def __init__(self, required_columns: list[str] | None = None) -> None:
        self.required_columns = required_columns or []

    def validate(self, rows: list[dict[str, Any]]) -> ValidationResult:
        issues: list[ValidationIssue] = []
        available_columns = set(rows[0].keys()) if rows else set()

        for column in self.required_columns:
            if column not in available_columns:
                issues.append(
                    ValidationIssue(
                        row_number=0,
                        column_name=column,
                        message="Missing required column",
                    )
                )

        for row_number, row in enumerate(rows, start=1):
            for column in self.required_columns:
                value = row.get(column)
                if value is None or (isinstance(value, str) and not value.strip()):
                    issues.append(
                        ValidationIssue(
                            row_number=row_number,
                            column_name=column,
                            message="Missing required value",
                        )
                    )

        return ValidationResult(rows=rows, issues=issues)
