"""Business logic for compensating-controls tab and working-paper export."""
from __future__ import annotations

from typing import Any, Callable

_PASSING_STATUSES = {"", "תקין"}

_DEFAULT_COLUMN_WIDTHS = {
    "0": 190,
    "1": 150,
    "2": 170,
    "3": 220,
    "4": 170,
}
_DEFAULT_ROW_HEIGHT = 56


def control_has_findings(detail_rows: list[dict[str, Any]]) -> bool:
    return any(
        str(row.get("status", "")).strip() not in _PASSING_STATUSES
        for row in detail_rows
    )


def build_findings_description(
    detail_rows: list[dict[str, Any]],
    *,
    max_chars: int = 4000,
) -> str:
    parts: list[str] = []
    seen: set[str] = set()
    for row in detail_rows:
        if str(row.get("status", "")).strip() in _PASSING_STATUSES:
            continue
        text = str(row.get("full_description") or row.get("description") or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        parts.append(text)
    combined = "\n".join(parts)
    if len(combined) <= max_chars:
        return combined
    return combined[: max_chars - 3].rstrip() + "..."


def _non_passing_rows(detail_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        row
        for row in detail_rows
        if str(row.get("status", "")).strip() not in _PASSING_STATUSES
    ]


def _count_distinct_users(rows: list[dict[str, Any]]) -> int:
    users: set[str] = set()
    for row in rows:
        for key in ("user_name", "UNAME", "BNAME"):
            value = str(row.get(key, "") or "").strip().upper()
            if value and value != "-":
                users.add(value)
                break
    return len(users)


def build_findings_brief_summary(
    detail_rows: list[dict[str, Any]],
    *,
    control_id: str = "",
    control_meta: dict[str, str] | None = None,
    summary_record: dict[str, Any] | None = None,
) -> str:
    """Return a one-sentence Hebrew summary of findings for the compensating-controls tab."""
    rows = _non_passing_rows(detail_rows)
    if not rows:
        return "-"

    resolved_control_id = control_id or str((summary_record or {}).get("control_id", "")).strip()
    check_type = str((control_meta or {}).get("check_type", "")).strip()
    finding_rows = [row for row in rows if str(row.get("status", "")).strip() == "עם ממצא"]
    review_rows = [row for row in rows if str(row.get("status", "")).strip() == "לסקירה"]

    user_total = _count_distinct_users(rows)
    user_findings = _count_distinct_users(finding_rows)
    user_review = _count_distinct_users(review_rows)

    if resolved_control_id == "MA7-17_AYALON_30" or "סקירת הרשאות" in check_type:
        parts: list[str] = []
        if user_findings:
            parts.append(f"{user_findings} משתמשים עם פרופילים חזקים")
        if user_review:
            parts.append(f"{user_review} משתמשים לסקירת הרשאות")
        if parts:
            return f"נמצאו {' ו-'.join(parts)}."
        if user_total:
            return f"נמצאו {user_total} משתמשים פעילים הדורשים טיפול."

    if resolved_control_id == "MA3-3_AYALON_14" or "פרופילים חזקים" in check_type:
        count = user_total or len(finding_rows) or len(rows)
        return f"נמצאו {count} משתמשים עם פרופילים חזקים."

    if "מדיניות סיסמאות" in check_type or resolved_control_id == "MA2-2_AYALON_6":
        count = len(finding_rows) or len(rows)
        return f"נמצאו {count} חריגות במדיניות סיסמאות."

    if resolved_control_id == "MA1-1&MA7-17_AYALON_2" or "סקירת משתמשים" in check_type:
        count = len(finding_rows) or user_total or len(rows)
        return f"נמצאו {count} משתמשים עם חריגות בסקירה."

    if "משתמשים חדשים" in check_type or resolved_control_id == "MA5.1-13_AYALON_24":
        count = len(finding_rows) or user_total or len(rows)
        return f"נמצאו {count} משתמשים חדשים הדורשים בדיקה."

    if "הרשאות" in check_type or user_total:
        if review_rows and not finding_rows:
            count = user_review or len(review_rows)
            return f"נמצאו {count} משתמשים לסקירת הרשאות."
        count = user_total or len(finding_rows) or len(rows)
        return f"נמצאו {count} משתמשים בעלי הרשאות רגישות."

    if finding_rows and review_rows:
        return f"נמצאו {len(finding_rows)} ממצאים ו-{len(review_rows)} רשומות לסקירה."
    if review_rows:
        return f"נמצאו {len(review_rows)} רשומות לסקירה."
    count = len(finding_rows) or len(rows)
    if check_type:
        return f"נמצאו {count} ממצאים בבקרת {check_type}."
    return f"נמצאו {count} ממצאים."


def build_compensating_control_rows(
    summary_records: dict[str, dict[str, Any]],
    details_by_control: dict[str, list[dict[str, Any]]],
    compensating_state: dict[str, dict[str, Any]],
    get_meta_cb: Callable[[str], dict[str, str]],
    is_in_scope_cb: Callable[[str], bool],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for control_id in sorted(summary_records.keys()):
        if not is_in_scope_cb(control_id):
            continue
        summary_record = summary_records.get(control_id, {})
        detail_rows = details_by_control.get(control_id, [])
        if not control_has_findings(detail_rows):
            continue
        meta = get_meta_cb(control_id)
        rows.append(
            {
                "control_id": control_id,
                "risk_description": meta.get("risk_description", "-") or "-",
                "description": meta.get("description", "-") or "-",
                "findings_brief": build_findings_brief_summary(
                    detail_rows,
                    control_id=control_id,
                    control_meta=meta,
                    summary_record=summary_record,
                ),
                "findings_description": build_findings_description(detail_rows),
                "attachment": compensating_state.get(control_id),
            }
        )
    return rows


DEFAULT_COMPENSATING_COLUMN_WIDTHS = _DEFAULT_COLUMN_WIDTHS
DEFAULT_COMPENSATING_ROW_HEIGHT = _DEFAULT_ROW_HEIGHT
