from typing import Any

from src.models.validation_result import ValidationIssue

PROFILE_REQUIRED_COLUMNS: dict[str, list[str]] = {
    "USERS": ["USER_NAME"],
    "USR02": ["BNAME"],
    "ADR6_USR21": [],
    "AGR_USERS": ["AGR_NAME", "UNAME"],
    "AGR_1251": ["AGR_NAME", "OBJECT", "FIELD", "LOW"],
    "AGR_1252": ["AGR_NAME", "LOW"],
    "AGR_DEFINE": ["AGR_NAME"],
    "UST04": ["BNAME", "PROFILE"],
    "E070": ["TRKORR", "AS4USER"],
    "T000": ["MANDT"],
    "STMS": ["TRKORR"],
    "RSPARAM": ["PARAMETER", "VALUE"],
    "TPFET": ["PARAMETER", "VALUE"],
    "M_PASSWORD_POLICY": ["PROPERTY", "VALUE"],
    "GRANTED_PRIVILEGES": ["GRANTEE", "PRIVILEGE"],
    "AUDIT_POLICIES": ["AUDIT_POLICY_NAME", "IS_AUDIT_POLICY_ACTIVE"],
}

PROFILE_REQUIRED_ANY_GROUPS: dict[str, list[tuple[str, ...]]] = {
    "USERS": [("LAST_SUCCESSFUL_CONNECT", "LAST_SUCCESSFUL_CONNECT_DATE")],
    "USR02": [("TRDAT", "LTIME")],
    "ADR6_USR21": [],
    "E070": [("AS4DATE", "TRFUNCTION")],
    "RSPARAM": [("PARAMETER", "NAME")],
    "TPFET": [("PARAMETER", "NAME")],
    "M_INIFILE_CONTENTS": [
        ("SECTION", "SECTION_NAME"),
        ("KEY", "KEY_NAME", "PARAMETER_NAME", "PROPERTY"),
        ("VALUE", "CONFIGURED_VALUE", "CURRENT_VALUE"),
    ],
}

PROFILE_OPTIONAL_VALUE_COLUMNS: dict[str, set[str]] = {
    "AGR_1251": {"LOW", "HIGH"},
    "AGR_1252": {"HIGH"},
    "STMS": {"STATUS"},
}

PROFILE_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "TRKORR": ("REQUEST", "REQUEST NUMBER", "TRANSPORT REQUEST"),
    "STATUS": ("RC", "RETURN CODE", "FUNCTION", "TRFUNCTION"),
    "AS4USER": ("OWNER", "USER", "CREATED BY", "USER NAME"),
    "AS4DATE": ("DATE", "CHANGE DATE"),
    "TRFUNCTION": ("RC", "RETURN CODE", "FUNCTION", "STATUS"),
    "MANDT": ("CLT", "CLIENT"),
    "BNAME": ("USER", "USER NAME", "USERNAME"),
    "UNAME": ("USER", "USER NAME", "USERNAME"),
    "USER_NAME": ("USER", "USER NAME", "USERNAME"),
    "AGR_NAME": ("ROLE", "ROLE NAME"),
    "PROFILE": ("PROFILE NAME",),
    "PARAMETER": ("NAME", "PARAMETER NAME", "PROPERTY"),
    "VALUE": ("CURRENT VALUE", "CONFIGURED VALUE"),
    "ADDRNUMBER": ("ADDRESS NUMBER",),
    "PERSNUMBER": ("PERSON NUMBER",),
    "SMTP_ADDR": ("EMAIL", "E-MAIL", "EMAIL ADDRESS", "SMTP ADDRESS"),
    "OBJECT": ("AUTH OBJECT", "AUTHORIZATION OBJECT"),
    "FIELD": ("FIELD NAME", "AUTH FIELD"),
    "LOW": ("LOW VALUE", "FROM", "FROM VALUE"),
    "HIGH": ("HIGH VALUE", "TO", "TO VALUE"),
}

PROFILE_STRUCTURE_RULES: dict[str, dict[str, Any]] = {
    "USR02": {
        "required_all": ["BNAME", "UFLAG"],
        "required_one_of": [("TRDAT", "LTIME")],
        "friendly_name": "USR02",
    },
    "ADR6_USR21": {
        "alternatives": [
            ["ADDRNUMBER", "SMTP_ADDR"],
            ["BNAME", "PERSNUMBER"],
        ],
        "friendly_name": "ADR6 / USR21",
    },
    "AGR_USERS": {
        "required_all": ["AGR_NAME", "UNAME"],
        "friendly_name": "AGR_USERS",
    },
    "AGR_1251": {
        "required_all": ["AGR_NAME", "OBJECT", "FIELD", "LOW"],
        "friendly_name": "AGR_1251",
    },
    "AGR_1252": {
        "required_all": ["AGR_NAME", "LOW"],
        "friendly_name": "AGR_1252",
    },
    "AGR_DEFINE": {
        "required_all": ["AGR_NAME"],
        "friendly_name": "AGR_DEFINE",
    },
    "UST04": {
        "required_all": ["BNAME", "PROFILE"],
        "friendly_name": "UST04",
    },
    "E070": {
        "required_all": ["TRKORR", "AS4USER"],
        "required_one_of": [("AS4DATE", "TRFUNCTION")],
        "friendly_name": "E070",
    },
    "T000": {
        "required_all": ["MANDT"],
        "friendly_name": "T000",
    },
    "STMS": {
        "required_one_of": [("TRKORR", "REQUEST")],
        "friendly_name": "STMS",
    },
    "RSPARAM": {
        "alternatives": [
            ["PARAMETER", "VALUE"],
            ["NAME", "VALUE"],
        ],
        "friendly_name": "RSPARAM",
    },
    "TPFET": {
        "alternatives": [
            ["PARAMETER", "VALUE"],
            ["NAME", "VALUE"],
            ["NAME", "CURRENT_VALUE"],
        ],
        "friendly_name": "TPFET",
    },
}

PASSWORD_POLICY_RULES = [
    ("minimal_password_length", 8, "minimum", "אורך סיסמה מינימלי חייב להיות לפחות 8"),
    ("force_first_password_change", "TRUE", "exact", "חובת החלפת סיסמה ראשונית חייבת להיות פעילה"),
    ("maximum_invalid_connect_attempts", 6, "maximum", "מספר ניסיונות התחברות שגויים חייב להיות מוגבל"),
    ("last_used_passwords", 5, "minimum", "היסטוריית סיסמאות חייבת לכלול לפחות 5 ערכים"),
    ("password_lock_for_system_user", "TRUE", "exact", "נעילת משתמשי SYSTEM חייבת להיות פעילה"),
    ("detailed_error_on_connect", "FALSE", "exact", "אין לחשוף הודעות שגיאה מפורטות בהתחברות"),
]

INI_SECURITY_RULES = [
    ("global.ini", "auditing configuration", "global_auditing_state", "true", "exact", "Audit trail גלובלי חייב להיות פעיל"),
    ("global.ini", "persistence", "log_mode", "normal", "exact", "Log mode חייב להיות NORMAL"),
    ("indexserver.ini", "password policy", "detailed_error_on_connect", "false", "exact", "אין לחשוף הודעות שגיאה מפורטות בהתחברות"),
    ("indexserver.ini", "password policy", "password_lock_for_system_user", "true", "exact", "נעילת משתמשי SYSTEM חייבת להיות פעילה"),
    ("indexserver.ini", "password policy", "force_first_password_change", "true", "exact", "חובת החלפת סיסמה ראשונית חייבת להיות פעילה"),
    ("indexserver.ini", "password policy", "minimal_password_length", 8, "minimum", "אורך סיסמה מינימלי חייב להיות לפחות 8"),
    ("indexserver.ini", "password policy", "maximum_invalid_connect_attempts", 6, "maximum", "מספר ניסיונות התחברות שגויים חייב להיות מוגבל"),
    ("indexserver.ini", "password policy", "last_used_passwords", 5, "minimum", "היסטוריית סיסמאות חייבת לכלול לפחות 5 ערכים"),
    ("indexserver.ini", "password policy", "password_expire_warning_time", 14, "minimum", "יש להתריע מראש לפני פקיעת סיסמה"),
]

CRITICAL_PRIVILEGES = {
    "AUDIT ADMIN",
    "AUDIT OPERATOR",
    "DATA ADMIN",
    "INIFILE ADMIN",
    "LOG ADMIN",
    "ROLE ADMIN",
    "SERVICE ADMIN",
    "TRUST ADMIN",
    "USER ADMIN",
    "BACKUP ADMIN",
}


def normalize_name(value: object) -> str:
    return str(value).strip().upper()


def normalize_text(value: object) -> str:
    return str(value).strip().casefold()


def filter_required_value_columns(profile: str | None, required_columns: list[str]) -> list[str]:
    optional_columns = PROFILE_OPTIONAL_VALUE_COLUMNS.get(profile or "", set())
    return [
        column
        for column in required_columns
        if normalize_name(column) not in optional_columns
    ]


def get_column_aliases(candidate: str) -> list[str]:
    normalized_candidate = normalize_name(candidate)
    aliases = [normalized_candidate]
    aliases.extend(normalize_name(alias) for alias in PROFILE_COLUMN_ALIASES.get(normalized_candidate, ()))
    return list(dict.fromkeys(aliases))


def matches_column_alias(available_columns: set[str], candidate: str) -> bool:
    return any(alias in available_columns for alias in get_column_aliases(candidate))


def detect_validation_profile(source_name: str | None, rows: list[dict[str, Any]]) -> str | None:
    file_name = (source_name or "").strip().lower()
    columns = {normalize_name(column) for row in rows[:1] for column in row.keys()}

    slot_name_map = {
        "usr02": "USR02",
        "adr6_usr21": "ADR6_USR21",
        "adr6": "ADR6_USR21",
        "usr21": "ADR6_USR21",
        "agr_users": "AGR_USERS",
        "agr_1251": "AGR_1251",
        "agr_1252": "AGR_1252",
        "agr_define": "AGR_DEFINE",
        "ust04": "UST04",
        "e070": "E070",
        "t000": "T000",
        "stms": "STMS",
        "rsparam": "RSPARAM",
        "tpfet": "TPFET",
    }
    for token, profile in slot_name_map.items():
        if token in file_name:
            return profile

    if matches_column_alias(columns, "SMTP_ADDR") and (
        matches_column_alias(columns, "ADDRNUMBER") or matches_column_alias(columns, "PERSNUMBER")
    ):
        return "ADR6_USR21"
    if matches_column_alias(columns, "TRKORR") and (
        "SHORT TEXT" in columns or matches_column_alias(columns, "STATUS") or matches_column_alias(columns, "AS4USER")
    ):
        return "STMS"
    if matches_column_alias(columns, "BNAME") and matches_column_alias(columns, "UFLAG"):
        return "USR02"
    if matches_column_alias(columns, "AGR_NAME") and matches_column_alias(columns, "OBJECT") and matches_column_alias(columns, "FIELD"):
        return "AGR_1251"
    if matches_column_alias(columns, "AGR_NAME") and matches_column_alias(columns, "UNAME"):
        return "AGR_USERS"
    if matches_column_alias(columns, "BNAME") and matches_column_alias(columns, "PROFILE"):
        return "UST04"
    if matches_column_alias(columns, "TRKORR") and matches_column_alias(columns, "AS4USER"):
        return "E070"
    if {"PROPERTY", "VALUE"}.issubset(columns):
        return "M_PASSWORD_POLICY"
    if {"AUDIT_POLICY_NAME", "IS_AUDIT_POLICY_ACTIVE"}.issubset(columns):
        return "AUDIT_POLICIES"
    if {"GRANTEE", "PRIVILEGE"}.issubset(columns):
        return "GRANTED_PRIVILEGES"
    if ({"SECTION", "KEY", "VALUE"}.issubset(columns) or {"SECTION_NAME", "KEY_NAME", "CONFIGURED_VALUE"}.issubset(columns)):
        return "M_INIFILE_CONTENTS"
    if matches_column_alias(columns, "USER_NAME"):
        return "USERS"

    return None


def build_profile_issues(profile: str | None, rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    if not profile:
        return []

    issues: list[ValidationIssue] = []
    available_columns = {
        normalize_name(column)
        for row in rows[:1]
        for column in row.keys()
        if not str(column).startswith("__")
    }

    issues.extend(_validate_expected_structure(profile, available_columns))

    for column in PROFILE_REQUIRED_COLUMNS.get(profile, []):
        if not matches_column_alias(available_columns, column):
            issues.append(ValidationIssue(row_number=0, column_name=column, message="עמודת חובה חסרה"))

    for group in PROFILE_REQUIRED_ANY_GROUPS.get(profile, []):
        if not any(matches_column_alias(available_columns, candidate) for candidate in group):
            issues.append(
                ValidationIssue(
                    row_number=0,
                    column_name=" / ".join(group),
                    message=f"נדרשת לפחות אחת מהעמודות: {' / '.join(group)}",
                )
            )

    if profile == "M_PASSWORD_POLICY":
        issues.extend(_evaluate_password_policy(rows))
    elif profile == "AUDIT_POLICIES":
        issues.extend(_evaluate_audit_policies(rows))
    elif profile == "M_INIFILE_CONTENTS":
        issues.extend(_evaluate_ini_contents(rows))
    elif profile == "GRANTED_PRIVILEGES":
        issues.extend(_evaluate_critical_privileges(rows))

    return issues


def _validate_expected_structure(profile: str, available_columns: set[str]) -> list[ValidationIssue]:
    if profile not in PROFILE_STRUCTURE_RULES:
        return []

    rule = PROFILE_STRUCTURE_RULES[profile]
    friendly_name = str(rule.get("friendly_name", profile))

    alternatives = rule.get("alternatives", [])
    if alternatives:
        for alternative in alternatives:
            if all(matches_column_alias(available_columns, column) for column in alternative):
                return []

        expected_text = " או ".join(" + ".join(option) for option in alternatives)
        return [
            ValidationIssue(
                row_number=0,
                column_name=friendly_name,
                message=f"הקובץ אינו תואם למבנה המצופה עבור המשבצת {friendly_name}. יש לצפות לאחד מהמבנים: {expected_text}",
            )
        ]

    missing_columns = [
        column for column in rule.get("required_all", []) if not matches_column_alias(available_columns, column)
    ]
    missing_groups = [
        " / ".join(group)
        for group in rule.get("required_one_of", [])
        if not any(matches_column_alias(available_columns, column) for column in group)
    ]

    if not missing_columns and not missing_groups:
        return []

    message_parts: list[str] = []
    if missing_columns:
        message_parts.append(f"חסרות עמודות מזהות: {', '.join(missing_columns)}")
    if missing_groups:
        message_parts.append(f"נדרשת לפחות אחת מהעמודות: {' ; '.join(missing_groups)}")

    return [
        ValidationIssue(
            row_number=0,
            column_name=friendly_name,
            message=f"הקובץ אינו תואם למבנה המצופה עבור המשבצת {friendly_name}. {' '.join(message_parts)}",
        )
    ]


def _find_column_name(row: dict[str, Any], candidates: tuple[str, ...]) -> str | None:
    normalized_map = {normalize_name(column): column for column in row.keys()}
    for candidate in candidates:
        for alias in get_column_aliases(candidate):
            if alias in normalized_map:
                return normalized_map[alias]
    return None


def _parse_numeric(value: object) -> float | None:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _evaluate_password_policy(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    property_map: dict[str, tuple[int, object]] = {}

    for row_number, row in enumerate(rows, start=1):
        property_column = _find_column_name(row, ("PROPERTY",))
        value_column = _find_column_name(row, ("VALUE",))
        if property_column and value_column:
            property_map[normalize_text(row[property_column])] = (row_number, row[value_column])

    for property_name, expected, rule_type, message in PASSWORD_POLICY_RULES:
        if property_name not in property_map:
            issues.append(ValidationIssue(row_number=0, column_name=property_name, message=f"לא נמצא פרמטר נדרש: {property_name}"))
            continue

        row_number, actual_value = property_map[property_name]
        if not _compare_values(actual_value, expected, rule_type):
            issues.append(ValidationIssue(row_number=row_number, column_name=property_name, message=message))

    return issues


def _evaluate_audit_policies(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    active_found = False
    for row in rows:
        active_column = _find_column_name(row, ("IS_AUDIT_POLICY_ACTIVE",))
        if active_column and normalize_text(row.get(active_column, "")) in {"true", "yes", "1"}:
            active_found = True
            break

    if active_found:
        return []

    return [ValidationIssue(row_number=0, column_name="IS_AUDIT_POLICY_ACTIVE", message="לפחות מדיניות Audit אחת חייבת להיות פעילה")]


def _evaluate_ini_contents(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []

    for expected_file, expected_section, expected_key, expected_value, rule_type, message in INI_SECURITY_RULES:
        matched = False
        for row_number, row in enumerate(rows, start=1):
            file_column = _find_column_name(row, ("FILE_NAME", "LAYER_FILE_NAME"))
            section_column = _find_column_name(row, ("SECTION", "SECTION_NAME"))
            key_column = _find_column_name(row, ("KEY", "KEY_NAME", "PARAMETER_NAME", "PROPERTY"))
            value_column = _find_column_name(row, ("VALUE", "CONFIGURED_VALUE", "CURRENT_VALUE"))

            if not section_column or not key_column or not value_column:
                continue

            file_matches = True if not file_column else normalize_text(row[file_column]) == normalize_text(expected_file)
            section_matches = normalize_text(row[section_column]) == normalize_text(expected_section)
            key_matches = normalize_text(row[key_column]) == normalize_text(expected_key)

            if file_matches and section_matches and key_matches:
                matched = True
                if not _compare_values(row[value_column], expected_value, rule_type):
                    issues.append(ValidationIssue(row_number=row_number, column_name=expected_key, message=message))
                break

        if not matched:
            issues.append(ValidationIssue(row_number=0, column_name=expected_key, message=f"לא נמצאה הגדרה נדרשת: {expected_key}"))

    return issues


def _evaluate_critical_privileges(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    for row_number, row in enumerate(rows, start=1):
        grantee_column = _find_column_name(row, ("GRANTEE",))
        privilege_column = _find_column_name(row, ("PRIVILEGE",))
        if not grantee_column or not privilege_column:
            continue

        privilege_name = normalize_name(row[privilege_column])
        if privilege_name in CRITICAL_PRIVILEGES:
            issues.append(
                ValidationIssue(
                    row_number=row_number,
                    column_name=row[grantee_column],
                    message=f"זוהתה הרשאה קריטית הדורשת סקירה: {row[privilege_column]}",
                )
            )

    return issues


def _compare_values(actual_value: object, expected_value: object, rule_type: str) -> bool:
    if rule_type == "exact":
        return normalize_text(actual_value) == normalize_text(expected_value)

    actual_number = _parse_numeric(actual_value)
    expected_number = _parse_numeric(expected_value)
    if actual_number is None or expected_number is None:
        return False

    if rule_type == "minimum":
        return actual_number >= expected_number
    if rule_type == "maximum":
        return actual_number <= expected_number
    return False
