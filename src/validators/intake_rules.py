from __future__ import annotations

from typing import Iterable

from src.models.validation_result import ValidationIssue


# Intake issues are structural/file-level or mandatory-value violations discovered
# during ingestion. Centralizing this rule avoids drift between pipeline/reporting.
def is_intake_issue(issue: ValidationIssue) -> bool:
    msg = str(issue.message or "")
    if issue.row_number == 0:
        return (
            "עמודת חובה חסרה" in msg
            or "אינו תואם למבנה" in msg
            or "נדרשת לפחות" in msg
        )
    return issue.row_number > 0 and "ערך חובה חסר" in msg


def has_intake_issues(issues: Iterable[ValidationIssue]) -> bool:
    return any(is_intake_issue(issue) for issue in issues)


def intake_failure_reasons(issues: Iterable[ValidationIssue]) -> str:
    reasons: list[str] = []
    for issue in issues:
        if not is_intake_issue(issue):
            continue
        text = str(issue.message or "").strip()
        if text and text not in reasons:
            reasons.append(text)
    if not reasons:
        return "-"
    return " | ".join(reasons)
