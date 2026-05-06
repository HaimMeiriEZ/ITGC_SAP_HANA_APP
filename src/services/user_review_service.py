from __future__ import annotations

from typing import Iterable


ALLOWED_REVIEW_FIELDS = {"REVIEW_STATUS", "TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES"}


def reviewer_state_key(mandt: object, bname: object) -> str:
    mandt_value = "" if mandt is None else str(mandt).strip()
    bname_value = "" if bname is None else str(bname).strip()
    return f"{mandt_value}|{bname_value}"


def normalize_reviewer_status(value: object, review_status_options: Iterable[str], default_status: str) -> str:
    normalized_value = "" if value is None else str(value).strip()
    if normalized_value in review_status_options:
        return normalized_value
    return default_status


def default_reviewer_values(default_status: str) -> dict[str, str]:
    return {
        "REVIEW_STATUS": default_status,
        "TECH_REVIEW_NOTES": "",
        "BUS_REVIEW_NOTES": "",
    }


def normalize_review_field(field_name: str) -> str:
    return "TECH_REVIEW_NOTES" if field_name == "REVIEW_NOTES" else field_name


def has_review_note(technical_note: object, business_note: object) -> bool:
    return bool(str(technical_note or "").strip() or str(business_note or "").strip())


def is_user_review_complete(
    review_status: object,
    findings_description: object,
    technical_note: object,
    business_note: object,
    reviewed_statuses: set[str],
    review_status_options: Iterable[str],
    default_status: str,
) -> bool:
    normalized_status = normalize_reviewer_status(review_status, review_status_options, default_status)
    if normalized_status not in reviewed_statuses:
        return False

    if normalized_status == "נבדק - לא תקין":
        return has_review_note(technical_note, business_note)

    if str(findings_description or "").strip():
        return has_review_note(technical_note, business_note)

    return True


def build_user_review_incomplete_reason(
    review_status: object,
    findings_description: object,
    reviewed_statuses: set[str],
    review_status_options: Iterable[str],
    default_status: str,
) -> str:
    normalized_status = normalize_reviewer_status(review_status, review_status_options, default_status)
    findings_text = str(findings_description or "").strip()

    if normalized_status not in reviewed_statuses:
        return "סטטוס הסקירה עדיין אינו מסומן כמשתמש שנבדק."
    if normalized_status == "נבדק - לא תקין":
        return "המשתמש סומן כלא תקין אך לא הוזנה הערה טכנית או עסקית."
    if findings_text:
        return "המשתמש סומן כתקין למרות שקיים ממצא, אך לא הוזנה הערה טכנית או עסקית."
    return "הסקירה טרם הושלמה בהתאם לכלל ההשלמה שהוגדר."
