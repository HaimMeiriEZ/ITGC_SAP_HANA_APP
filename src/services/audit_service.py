from __future__ import annotations

from typing import Any, Callable

from src.models.validation_result import ValidationIssue


def build_audit_detail_row(
    issue: ValidationIssue | None,
    control_id: str,
    source_file: str,
    extraction_date: str,
    control_meta: dict[str, str],
    work_environment_label: str,
    control_snapshot: dict[str, str] | None = None,
) -> dict[str, Any]:
    if issue is None:
        return {
            "control_id": control_id,
            "source_file": source_file,
            "extraction_date": extraction_date,
            "work_environment": work_environment_label,
            "category": control_meta.get("category", "-"),
            "risk_level": control_meta.get("risk_level", "-"),
            "description": control_meta.get("description", "-"),
            "check_type": control_meta.get("check_type", "-"),
            "actual_value": (control_snapshot or {}).get("actual_value", "-"),
            "expected_value": (control_snapshot or {}).get("expected_value", "-"),
            "status": (control_snapshot or {}).get("status", "תקין"),
            "full_description": (control_snapshot or {}).get("full_description", "לא נמצאו ממצאים עבור הבקרה."),
        }

    return {
        "control_id": control_id,
        "source_file": issue.source_file or source_file,
        "extraction_date": extraction_date,
        "work_environment": work_environment_label,
        "category": issue.category or control_meta.get("category", "-"),
        "risk_level": issue.risk_level or control_meta.get("risk_level", "-"),
        "description": issue.description or control_meta.get("description", "-"),
        "check_type": issue.check_type or control_meta.get("check_type", "-"),
        "actual_value": issue.actual_value or "-",
        "expected_value": issue.expected_value or "-",
        "status": issue.status or "עם ממצא",
        "full_description": issue.full_description or issue.message,
    }


def upsert_audit_control_data(
    audit_summary_records: dict[str, dict[str, Any]],
    audit_details_by_control: dict[str, list[dict[str, Any]]],
    slot_key: str,
    result: Any,
    audit_issues: list[ValidationIssue],
    extraction_date: str,
    work_environment_label: str,
    slot_display_name: str,
    get_audit_control_definition_cb: Callable[[str], dict[str, str]],
    get_profile_audit_controls_cb: Callable[[str], list[str]],
    count_stms_control_records_cb: Callable[[list[dict[str, Any]]], int],
    build_password_control_snapshots_cb: Callable[[list[dict[str, Any]]], dict[str, dict[str, str]]],
) -> None:
    control_ids = [issue.control_id for issue in audit_issues if issue.control_id]
    expected_controls = get_profile_audit_controls_cb(getattr(result, "detected_profile", slot_key))
    all_control_ids = sorted(set(control_ids + expected_controls))
    if not all_control_ids:
        return

    source_file_label = ", ".join(getattr(result, "source_files", []) or [slot_display_name])
    detected_profile = str(getattr(result, "detected_profile", slot_key) or slot_key).upper()
    rows = getattr(result, "rows", [])
    password_snapshots = build_password_control_snapshots_cb(rows) if detected_profile in {"RSPARAM", "TPFET"} else {}

    for control_id in all_control_ids:
        control_meta = get_audit_control_definition_cb(control_id)
        control_issues = [issue for issue in audit_issues if issue.control_id == control_id]
        finding_records = len(control_issues)

        if control_id == "MC7-25_AYALON_44":
            total_records = count_stms_control_records_cb(rows)
        else:
            total_records = 1

        if total_records <= 0:
            total_records = max(finding_records, 1)
        valid_records = max(total_records - finding_records, 0)

        audit_summary_records[control_id] = {
            "control_id": control_id,
            "check_type": control_meta.get("check_type", "-"),
            "source_file": source_file_label,
            "extraction_date": extraction_date,
            "work_environment": work_environment_label,
            "risk_level": control_meta.get("risk_level", "-"),
            "description": control_meta.get("description", "-"),
            "valid_records": valid_records,
            "finding_records": finding_records,
            "total_records": total_records,
        }

        detail_rows = [
            build_audit_detail_row(
                issue,
                control_id,
                source_file_label,
                extraction_date,
                control_meta,
                work_environment_label,
            )
            for issue in control_issues
        ]
        if not detail_rows:
            detail_rows = [
                build_audit_detail_row(
                    None,
                    control_id,
                    source_file_label,
                    extraction_date,
                    control_meta,
                    work_environment_label,
                    password_snapshots.get(control_id),
                )
            ]
        audit_details_by_control[control_id] = detail_rows


def sync_user_review_completion_finding(
    audit_summary_records: dict[str, dict[str, Any]],
    audit_details_by_control: dict[str, list[dict[str, Any]]],
    control_id: str,
    control_meta: dict[str, str],
    source_file_label: str,
    extraction_date: str,
    work_environment_label: str,
    reviewed_rows: int,
    preview_rows: list[dict[str, str]],
    incomplete_rows: list[dict[str, str]],
    build_incomplete_reason_cb: Callable[[dict[str, str]], str],
) -> None:
    audit_summary_records.pop(control_id, None)
    audit_details_by_control.pop(control_id, None)

    total_rows = len(preview_rows)
    if total_rows <= 0 or not incomplete_rows:
        return

    audit_summary_records[control_id] = {
        "control_id": control_id,
        "check_type": control_meta.get("check_type", "השלמת סקירת משתמשים"),
        "source_file": source_file_label,
        "extraction_date": extraction_date,
        "work_environment": work_environment_label,
        "risk_level": control_meta.get("risk_level", "בינוני"),
        "description": control_meta.get("description", "סקירת המשתמשים טרם הושלמה במלואה."),
        "valid_records": reviewed_rows,
        "finding_records": len(incomplete_rows),
        "total_records": total_rows,
    }

    audit_details_by_control[control_id] = [
        {
            "control_id": control_id,
            "source_file": source_file_label,
            "extraction_date": extraction_date,
            "work_environment": work_environment_label,
            "category": control_meta.get("category", "MA - ניהול גישה"),
            "risk_level": control_meta.get("risk_level", "בינוני"),
            "description": control_meta.get("description", "סקירת המשתמשים טרם הושלמה במלואה."),
            "check_type": control_meta.get("check_type", "השלמת סקירת משתמשים"),
            "actual_value": str(preview_row.get("BNAME", "-")) or "-",
            "expected_value": "השלמת סקירה בהתאם לכלל ההשלמה",
            "status": "עם ממצא",
            "full_description": build_incomplete_reason_cb(preview_row),
        }
        for preview_row in incomplete_rows
    ]


def sorted_audit_summary_rows(audit_summary_records: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(audit_summary_records.values(), key=lambda item: str(item.get("control_id", "")))


def build_audit_summary_values(row_data: dict[str, Any]) -> list[str]:
    return [
        str(row_data.get("control_id", "-")),
        str(row_data.get("check_type", "-")),
        str(row_data.get("risk_level", "-")),
        str(row_data.get("valid_records", 0)),
        str(row_data.get("finding_records", 0)),
        str(row_data.get("total_records", 0)),
        str(row_data.get("work_environment", "-")),
        str(row_data.get("source_file", "-")),
        str(row_data.get("extraction_date", "-")),
        str(row_data.get("description", "-")),
    ]


def build_audit_detail_values(detail: dict[str, Any]) -> list[str]:
    return [
        str(detail.get("source_file", "-")),
        str(detail.get("extraction_date", "-")),
        str(detail.get("work_environment", "-")),
        str(detail.get("category", "-")),
        str(detail.get("risk_level", "-")),
        str(detail.get("description", "-")),
        str(detail.get("check_type", "-")),
        str(detail.get("actual_value", "-")),
        str(detail.get("expected_value", "-")),
        str(detail.get("status", "-")),
        str(detail.get("full_description", "-")),
    ]
