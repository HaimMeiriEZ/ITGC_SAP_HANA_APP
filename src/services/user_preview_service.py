from __future__ import annotations

from datetime import datetime
from typing import Any, Callable


def parse_user_preview_date(raw_value: object) -> datetime | None:
    normalized_value = "" if raw_value is None else str(raw_value).strip()
    if not normalized_value:
        return None

    supported_patterns = [
        "%Y-%m-%d",
        "%Y%m%d",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%d.%m.%y",
        "%d/%m/%y",
    ]
    for pattern in supported_patterns:
        try:
            return datetime.strptime(normalized_value, pattern)
        except ValueError:
            continue
    return None


def format_user_preview_value_for_display(_field_name: str, value: object) -> str:
    return "" if value is None else str(value).strip()


def get_user_preview_sort_value(field_name: str, value: object, date_fields: set[str]) -> str:
    normalized_value = "" if value is None else str(value).strip()
    if field_name in date_fields:
        parsed_date = parse_user_preview_date(normalized_value)
        return parsed_date.strftime("%Y%m%d") if parsed_date is not None else ""
    return normalized_value.casefold()


def filter_user_preview_rows(
    preview_rows: list[dict[str, str]],
    filter_mode: str,
    start_text: str,
    end_text: str,
) -> tuple[list[dict[str, str]], str]:
    if filter_mode == "all":
        return preview_rows, ""

    if not start_text or not end_text:
        return preview_rows, "כדי לסנן לפי פעילות בתקופה יש להזין תאריך התחלה ותאריך סיום."

    start_date = parse_user_preview_date(start_text)
    end_date = parse_user_preview_date(end_text)
    if start_date is None or end_date is None:
        return preview_rows, "יש להזין את טווח התאריכים בפורמט YYYY-MM-DD."
    if start_date > end_date:
        return preview_rows, "תאריך ההתחלה חייב להיות מוקדם או שווה לתאריך הסיום."

    filtered_rows: list[dict[str, str]] = []
    for preview_row in preview_rows:
        last_login_date = parse_user_preview_date(preview_row.get("TRDAT", ""))
        was_active_in_period = last_login_date is not None and start_date <= last_login_date <= end_date
        if filter_mode == "active" and was_active_in_period:
            filtered_rows.append(preview_row)
        elif filter_mode == "inactive" and not was_active_in_period:
            filtered_rows.append(preview_row)

    return filtered_rows, ""


def build_user_preview_rows(
    usr02_rows: list[dict[str, Any]],
    combined_rows: list[dict[str, Any]],
    get_row_value: Callable[[dict[str, Any], str], str],
    format_user_status: Callable[[object], str],
    extraction_date_text: str,
    work_environment_label: str,
    get_reviewer_values: Callable[[object, object], dict[str, str]],
    build_user_findings_description: Callable[[dict[str, str], str], str],
    default_review_status: str,
) -> list[dict[str, str]]:
    usr02_map: dict[tuple[str, str], dict[str, str]] = {}
    addr_users_map: dict[tuple[str, str], dict[str, str]] = {}
    email_by_addr: dict[str, str] = {}
    email_by_pers: dict[str, str] = {}

    for row in usr02_rows:
        mandt = get_row_value(row, "MANDT")
        bname = get_row_value(row, "BNAME")
        if not bname:
            continue
        raw_uflag = get_row_value(row, "UFLAG")
        usr02_map[(mandt, bname)] = {
            "MANDT": mandt,
            "BNAME": bname,
            "UFLAG": raw_uflag,
            "STATUS": format_user_status(raw_uflag),
            "TRDAT": get_row_value(row, "TRDAT"),
            "LTIME": get_row_value(row, "LTIME"),
            "GLTGV": get_row_value(row, "GLTGV"),
            "GLTGB": get_row_value(row, "GLTGB"),
            "USTYP": get_row_value(row, "USTYP"),
            "LOCNT": get_row_value(row, "LOCNT"),
            "PWDINITIAL": get_row_value(row, "PWDINITIAL"),
            "PWDCHGDATE": get_row_value(row, "PWDCHGDATE"),
            "PWDSETDATE": get_row_value(row, "PWDSETDATE"),
            "OCOD1": get_row_value(row, "OCOD1"),
            "PASSCODE": get_row_value(row, "PASSCODE"),
            "PWDSALTEDHASH": get_row_value(row, "PWDSALTEDHASH"),
            "SECURITY_POLICY": get_row_value(row, "SECURITY_POLICY"),
        }

    for row in combined_rows:
        addrnumber = get_row_value(row, "ADDRNUMBER")
        persnumber = get_row_value(row, "PERSNUMBER")
        smtp_addr = get_row_value(row, "SMTP_ADDR")

        if smtp_addr:
            if addrnumber:
                email_by_addr[addrnumber] = smtp_addr
            if persnumber:
                email_by_pers[persnumber] = smtp_addr

        bname = get_row_value(row, "BNAME")
        if not bname:
            continue

        mandt = get_row_value(row, "MANDT")
        key = (mandt, bname)
        current_entry = addr_users_map.setdefault(
            key,
            {
                "MANDT": mandt,
                "BNAME": bname,
                "NAME_FIRST": "",
                "NAME_LAST": "",
                "NAME_TEXTC": "",
                "COMPANY": "",
                "DEPARTMENT": "",
                "ADDRNUMBER": "",
                "PERSNUMBER": "",
                "SMTP_ADDR": "",
            },
        )

        for field_name in [
            "NAME_FIRST",
            "NAME_LAST",
            "NAME_TEXTC",
            "COMPANY",
            "DEPARTMENT",
            "ADDRNUMBER",
            "PERSNUMBER",
            "SMTP_ADDR",
        ]:
            field_value = get_row_value(row, field_name)
            if field_value and not current_entry[field_name]:
                current_entry[field_name] = field_value

    if usr02_map:
        ordered_keys = sorted(list(usr02_map.keys()), key=lambda item: (item[0], item[1]))
    else:
        ordered_keys = sorted(list(addr_users_map.keys()), key=lambda item: (item[0], item[1]))

    preview_rows: list[dict[str, str]] = []

    for key in ordered_keys:
        usr_entry = usr02_map.get(key, {})
        addr_entry = addr_users_map.get(key, {})
        merged_mandt = usr_entry.get("MANDT") or addr_entry.get("MANDT", "")
        merged_bname = usr_entry.get("BNAME") or addr_entry.get("BNAME", "")
        review_values = get_reviewer_values(merged_mandt, merged_bname)
        usr_entry["WORK_ENVIRONMENT"] = work_environment_label
        findings_description = build_user_findings_description(usr_entry, extraction_date_text)
        email_value = (
            addr_entry.get("SMTP_ADDR", "")
            or email_by_addr.get(addr_entry.get("ADDRNUMBER", ""), "")
            or email_by_pers.get(addr_entry.get("PERSNUMBER", ""), "")
        )

        preview_rows.append(
            {
                "MANDT": merged_mandt,
                "WORK_ENVIRONMENT": work_environment_label,
                "BNAME": merged_bname,
                "NAME_FIRST": addr_entry.get("NAME_FIRST", ""),
                "NAME_LAST": addr_entry.get("NAME_LAST", ""),
                "NAME_TEXTC": addr_entry.get("NAME_TEXTC", ""),
                "COMPANY": addr_entry.get("COMPANY", ""),
                "DEPARTMENT": addr_entry.get("DEPARTMENT", ""),
                "SMTP_ADDR": email_value,
                "STATUS": usr_entry.get("STATUS", "לא זמין"),
                "UFLAG": usr_entry.get("UFLAG", ""),
                "ADDRNUMBER": addr_entry.get("ADDRNUMBER", ""),
                "PERSNUMBER": addr_entry.get("PERSNUMBER", ""),
                "TRDAT": usr_entry.get("TRDAT", ""),
                "LTIME": usr_entry.get("LTIME", ""),
                "GLTGV": usr_entry.get("GLTGV", ""),
                "GLTGB": usr_entry.get("GLTGB", ""),
                "USTYP": usr_entry.get("USTYP", ""),
                "LOCNT": usr_entry.get("LOCNT", ""),
                "PWDINITIAL": usr_entry.get("PWDINITIAL", ""),
                "PWDCHGDATE": usr_entry.get("PWDCHGDATE", ""),
                "PWDSETDATE": usr_entry.get("PWDSETDATE", ""),
                "OCOD1": usr_entry.get("OCOD1", ""),
                "PASSCODE": usr_entry.get("PASSCODE", ""),
                "PWDSALTEDHASH": usr_entry.get("PWDSALTEDHASH", ""),
                "SECURITY_POLICY": usr_entry.get("SECURITY_POLICY", ""),
                "REVIEW_STATUS": review_values.get("REVIEW_STATUS", default_review_status),
                "FINDINGS_DESCRIPTION": findings_description,
                "TECH_REVIEW_NOTES": review_values.get("TECH_REVIEW_NOTES", ""),
                "BUS_REVIEW_NOTES": review_values.get("BUS_REVIEW_NOTES", ""),
                "LAST_IMPORT_DATE": review_values.get("LAST_IMPORT_DATE", ""),
            }
        )

    return preview_rows
