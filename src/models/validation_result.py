from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class ValidationIssue:
    row_number: int
    column_name: str
    message: str


@dataclass
class ValidationSummary:
    total_rows: int
    valid_rows: int
    invalid_rows: int
    is_valid: bool


@dataclass
class ValidationResult:
    rows: list[dict[str, Any]]
    issues: list[ValidationIssue] = field(default_factory=list)
    report_path: Path | None = None

    @property
    def summary(self) -> ValidationSummary:
        total_rows = len(self.rows)
        row_level_issues = {issue.row_number for issue in self.issues if issue.row_number > 0}
        invalid_rows = len(row_level_issues)

        if any(issue.row_number == 0 for issue in self.issues):
            invalid_rows = total_rows if total_rows else invalid_rows

        valid_rows = max(total_rows - invalid_rows, 0)
        return ValidationSummary(
            total_rows=total_rows,
            valid_rows=valid_rows,
            invalid_rows=invalid_rows,
            is_valid=len(self.issues) == 0,
        )
