from functools import lru_cache
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
}

PROFILE_REQUIRED_ANY_GROUPS: dict[str, list[tuple[str, ...]]] = {
    "USERS": [("LAST_SUCCESSFUL_CONNECT", "LAST_SUCCESSFUL_CONNECT_DATE")],
    "USR02": [("TRDAT", "LTIME")],
    "ADR6_USR21": [],
    "E070": [("AS4DATE", "TRFUNCTION")],
    "RSPARAM": [("PARAMETER", "NAME")],
    "TPFET": [("PARAMETER", "NAME")],
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
    "IMPORT_USER": ("IMPORT USER", "IMPORTED BY", "IMPORTED USER"),
    "AS4DATE": ("DATE", "CHANGE DATE"),
    "TRFUNCTION": ("RC", "RETURN CODE", "FUNCTION", "STATUS"),
    "MANDT": ("CLT", "CLIENT"),
    "BNAME": ("USER", "USER NAME", "USERNAME"),
    "UNAME": ("USER", "USER NAME", "USERNAME"),
    "USER_NAME": ("USER", "USER NAME", "USERNAME"),
    "AGR_NAME": ("ROLE", "ROLE NAME"),
    "PROFILE": ("PROFILE NAME",),
    "PARAMETER": ("NAME", "PARAMETER NAME", "PROPERTY"),
    "VALUE": (
        "CURRENT VALUE",
        "CONFIGURED VALUE",
        "PARAMETER VALUE",
        "ACTUAL VALUE",
        "USER-DEFINED VALUE",
        "SYSTEM DEFAULT VALUE",
        "SYSTEM DEFAULT VALUE(UNSUBSTITUTED FORM)",
    ),
    "ADDRNUMBER": ("ADDRESS NUMBER",),
    "PERSNUMBER": ("PERSON NUMBER",),
    "SMTP_ADDR": ("EMAIL", "E-MAIL", "EMAIL ADDRESS", "SMTP ADDRESS"),
    "NAME_FIRST": ("FIRST NAME", "GIVEN NAME"),
    "NAME_LAST": ("LAST NAME", "SURNAME", "FAMILY NAME"),
    "NAME_TEXTC": ("FULL NAME", "DISPLAY NAME", "FORMAL NAME"),
    "COMPANY": ("COMPANY NAME", "ORGANIZATION"),
    "DEPARTMENT": ("DEPARTMENT NAME", "ORG UNIT", "ORGANIZATIONAL UNIT"),
    "GLTGV": ("VALID FROM",),
    "GLTGB": ("VALID TO",),
    "USTYP": ("USER TYPE",),
    "LOCNT": ("NUMBER OF FAILED LOGON ATTEMPTS", "FAILED LOGON ATTEMPTS", "FAILED LOGONS"),
    "PWDINITIAL": ("PASSWORD INITIAL", "INITIAL PASSWORD", "PWD INITIAL"),
    "PWDCHGDATE": ("PASSWORD CHANGE DATE", "PWD CHANGE DATE"),
    "PWDSETDATE": ("PASSWORD SET DATE", "PWD SET DATE"),
    "OCOD1": ("PASSWORD",),
    "PASSCODE": ("PASSWORD HASH VALUE", "PASSWORD HASH VALUE (SHA1, 160 BIT)"),
    "PWDSALTEDHASH": ("PASSWORD HASH VAL", "PASSWORD HASH VALUE SALTED", "PASSWORD HASH VALUE (SALTED)"),
    "SECURITY_POLICY": ("SECURITY POLICY",),
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
            ["MANDT", "BNAME", "NAME_TEXTC"],
            ["MANDT", "BNAME", "NAME_FIRST", "NAME_LAST"],
        ],
        "friendly_name": "ADR6 / USER_ADDR",
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

SAP_APP_RSPARAM_RULES = [
    # (parameter_name, expected_value, rule_type, message)
    ("login/min_password_lng", 8, "minimum", "אורך סיסמה מינימלי חייב להיות לפחות 8 תווים"),
    ("login/fails_to_user_lock", 6, "maximum", "נעילת משתמש לאחר ניסיונות כושלים חייבת להיות לכל היותר 6"),
    ("login/failed_user_auto_unlock", 0, "maximum", "ביטול נעילה אוטומטי לאחר כישלון חייב להיות מבוטל (0)"),
    ("login/password_expiration_time", 90, "maximum", "תקופת תפוגת סיסמה חייבת להיות לכל היותר 90 ימים"),
    ("login/password_history_size", 5, "minimum", "היסטוריית סיסמאות חייבת לכלול לפחות 5 ערכים"),
    ("login/no_automatic_user_sapstar", 1, "minimum", "פרמטר SAP* האוטומטי חייב להיות מבוטל (1)"),
]

SAP_ITGC_RELEVANT_PARAMETERS = {
    "login/min_password_lng",
    "login/min_password_digits",
    "login/min_password_letters",
    "login/min_password_lowercase",
    "login/min_password_uppercase",
    "login/min_password_specials",
    "login/fails_to_user_lock",
    "login/failed_user_auto_unlock",
    "login/password_expiration_time",
    "login/password_history_size",
    "login/password_change_for_sso",
    "login/password_downwards_compatibility",
    "login/no_automatic_user_sapstar",
}

PROFILE_SCOPED_VALUE_PARAMETERS: dict[str, set[str]] = {
    "RSPARAM": SAP_ITGC_RELEVANT_PARAMETERS,
    "TPFET": SAP_ITGC_RELEVANT_PARAMETERS,
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


@lru_cache(maxsize=None)
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
        "addr_users": "ADR6_USR21",
        "user_addr": "ADR6_USR21",
        "usrs_aadr": "ADR6_USR21",
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
    if matches_column_alias(columns, "BNAME") and (
        matches_column_alias(columns, "NAME_TEXTC")
        or matches_column_alias(columns, "NAME_FIRST")
        or matches_column_alias(columns, "NAME_LAST")
        or matches_column_alias(columns, "COMPANY")
    ):
        return "ADR6_USR21"
    if matches_column_alias(columns, "TRKORR") and matches_column_alias(columns, "IMPORT_USER"):
        return "STMS"
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

    if profile in ("RSPARAM", "TPFET"):
        issues.extend(_evaluate_rsparam_policy(rows))

    return issues


def build_control_44_issues(
    profile: str | None,
    rows: list[dict[str, Any]],
    authorized_users: set[str],
) -> list[ValidationIssue]:
    """Control 44: only authorized users may import transports to production."""
    if profile != "STMS" or not rows:
        return []

    issues: list[ValidationIssue] = []
    for row_index, row in enumerate(rows, start=1):
        normalized_row = {
            normalize_name(key): value
            for key, value in row.items()
            if not str(key).startswith("__")
        }

        trkorr = ""
        for candidate in get_column_aliases("TRKORR"):
            if candidate in normalized_row:
                trkorr = str(normalized_row.get(candidate, "")).strip()
                if trkorr:
                    break

        import_user = ""
        for candidate in get_column_aliases("IMPORT_USER"):
            if candidate in normalized_row:
                import_user = str(normalized_row.get(candidate, "")).strip().upper()
                if import_user:
                    break

        if not trkorr or not import_user:
            continue

        if import_user not in authorized_users:
            as4date = ""
            for candidate in get_column_aliases("AS4DATE"):
                if candidate in normalized_row:
                    as4date = str(normalized_row.get(candidate, "")).strip()
                    if as4date:
                        break

            issues.append(
                ValidationIssue(
                    row_number=row_index,
                    column_name="IMPORT_USER",
                    message=f"בקרה 44: המשתמש {import_user} העביר את טרנספורט {trkorr} ב-{as4date} אך אינו ברשימת המורשים",
                    source_file=str(row.get("__source_file", "")),
                )
            )

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


def should_enforce_required_value(profile: str | None, row: dict[str, Any], column: str) -> bool:
    if normalize_name(column) != "VALUE":
        return True

    relevant_parameters = PROFILE_SCOPED_VALUE_PARAMETERS.get(profile or "")
    if not relevant_parameters:
        return True

    parameter_column = _find_column_name(row, ("PARAMETER", "NAME", "PROPERTY"))
    if not parameter_column:
        return True

    parameter_name = normalize_text(row.get(parameter_column, ""))
    return parameter_name in relevant_parameters


def _parse_numeric(value: object) -> float | None:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _evaluate_rsparam_policy(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    """Evaluate SAP APP (ABAP) password policy parameters from RSPARAM / TPFET data."""
    issues: list[ValidationIssue] = []
    param_map: dict[str, tuple[int, object]] = {}

    for row_number, row in enumerate(rows, start=1):
        param_column = _find_column_name(row, ("PARAMETER", "NAME"))
        value_column = _find_column_name(row, ("VALUE",))
        if param_column and value_column:
            param_map[normalize_text(row[param_column])] = (row_number, row[value_column])

    for param_name, expected, rule_type, message in SAP_APP_RSPARAM_RULES:
        if param_name not in param_map:
            issues.append(ValidationIssue(row_number=0, column_name=param_name, message=f"לא נמצא פרמטר נדרש: {param_name}"))
            continue

        row_number, actual_value = param_map[param_name]
        if not _compare_values(actual_value, expected, rule_type):
            issues.append(ValidationIssue(row_number=row_number, column_name=param_name, message=message))

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
