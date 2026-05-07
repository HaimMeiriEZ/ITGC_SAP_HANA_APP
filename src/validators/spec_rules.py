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

AUDIT_CONTROL_DEFINITIONS: dict[str, dict[str, str]] = {
    "44": {
        "category": "MC - ניהול שינויים",
        "risk_level": "גבוה",
        "check_type": "STMS - Import מורשים בלבד",
        "description": "Import לסביבת ייצור יתבצע רק על ידי משתמשים מורשים.",
    },
    "MA-PWD-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "מדיניות סיסמאות",
        "description": "אורך סיסמה מינימלי חייב להיות לפחות 8 תווים.",
    },
    "MA-PWD-02": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "מדיניות סיסמאות",
        "description": "נעילת משתמש לאחר ניסיונות כושלים חייבת להיות לכל היותר 6.",
    },
    "MA-PWD-03": {
        "category": "MA - ניהול גישה",
        "risk_level": "בינוני",
        "check_type": "מדיניות סיסמאות",
        "description": "ביטול נעילה אוטומטי לאחר כישלון חייב להיות מבוטל (0).",
    },
    "MA-PWD-04": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "מדיניות סיסמאות",
        "description": "תקופת תפוגת סיסמה חייבת להיות לכל היותר 90 ימים.",
    },
    "MA-PWD-05": {
        "category": "MA - ניהול גישה",
        "risk_level": "בינוני",
        "check_type": "מדיניות סיסמאות",
        "description": "היסטוריית סיסמאות חייבת לכלול לפחות 5 ערכים.",
    },
    "MA-PWD-06": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "משתמשי מערכת",
        "description": "פרמטר SAP* האוטומטי חייב להיות מבוטל (1).",
    },
    "MA-PERM-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "פרופילים למשתמשים חזקים",
        "description": "הקצאת פרופילי-על למשתמש מעניקה הרשאות מערכת רחבות ודורשת בקרה הדוקה.",
    },
    "MA-USRMGMT-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאות ניהול משתמשים",
        "description": "משתמשים בעלי הרשאות לניהול משתמשים (S_TCODE/SU01/SU10, S_USER_*) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-AUTHMGMT-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאות ניהול הרשאות",
        "description": "משתמשים בעלי הרשאות לניהול הרשאות (S_TCODE/PFCG, S_DEVELOP/ACGR, S_USER_ADM) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-RSCDOK99-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאות לתוכנית RSCDOK99",
        "description": "משתמשים בעלי הרשאה להרצת תוכנית RSCDOK99 (S_PROGRAM עם P_GROUP=RSCDOK99 ו-P_ACTION=SUB) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-DATAMGMT-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאות לניהול נתונים",
        "description": "משתמשים בעלי הרשאות לניהול נתונים (S_TCODE/SE16/SM30/SM31/SE16N/SE17/SM38/SE37, S_TABU_DIS, S_TABU_NAM, S_TABU_CLI, S_DATASET) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-TRANSPORT-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאה להעברת שינויים",
        "description": "משתמשים בעלי הרשאה להעברת שינויים (S_TCODE/STMS/SCC4, S_TABU_DIS/DICBERCLS=SS, S_TRANSPORT, S_CTS_ADMI) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-DEBUG-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאות לשימוש ב-DEBUG",
        "description": "משתמשים בעלי הרשאות DEBUG (S_TCODE/SE38/SA38/SE80/ST05, S_DEVELOP/OBJTYPE=DEBUG, S_DEVELOP/ACTVT, S_PROGRAM/P_ACTION=SUB, S_PROGRAM/P_GROUP=*, S_ADMI_FCD/PADM) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-JOBMGMT-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "AGR_1251+AGR_USERS - הרשאה לעידכון ג'ובים",
        "description": "משתמשים בעלי הרשאות ניהול ג'ובים (S_TCODE/SM36 וכד', S_BTCH_ADM, S_BTCH_JOB/DELE/RELE/PROT, S_BTCH_NAM/*, S_BTCH_MONI/DELE/RELE) זוהו לפי אובייקטי הרשאה ב-AGR_1251.",
    },
    "MA-REVIEW-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "בינוני",
        "check_type": "השלמת סקירת משתמשים",
        "description": "סקירת המשתמשים טרם הושלמה במלואה בהתאם לכלל ההשלמה שהוגדר.",
    },
    "MA-SOD-01": {
        "category": "MA - ניהול גישה",
        "risk_level": "גבוה",
        "check_type": "הפרדת תפקידים - מפתחים בסביבת ייצור",
        "description": "זיהוי מפתחים המוגדרים כמשתמשים פעילים בסביבת הייצור (Segregation of Duties).",
    },
}

PROFILE_AUDIT_CONTROLS: dict[str, list[str]] = {
    "STMS": ["44"],
    "RSPARAM": ["MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06"],
    "TPFET": ["MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06"],
    "UST04": ["MA-PERM-01"],
}

# Criteria for user-management permission check (control MA-USRMGMT-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# A row in AGR_1251 qualifies when its OBJECT+FIELD match and either:
#   - its LOW or HIGH value is "*" (wildcard), OR
#   - its LOW or HIGH value is in the qualifying set below.
USER_MGMT_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",    "TCD"):   {"SU01", "SU10"},
    ("S_USER_PRO", "ACTVT"): {"01", "02"},
    ("S_USER_AGR", "ACTVT"): {"01", "02"},
    ("S_USER_AUT", "ACTVT"): {"01", "02"},
    ("S_USER_GRP", "ACTVT"): {"01", "02"},
    ("S_USER_SYS", "ACTVT"): {"01", "02"},
    ("S_USER_SAS", "ACTVT"): {"01", "02"},
}

# Criteria for authorization-management permission check (control MA-AUTHMGMT-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# A row in AGR_1251 qualifies when its OBJECT+FIELD match and either:
#   - its LOW or HIGH value is "*" (wildcard), OR
#   - its LOW or HIGH value is in the qualifying set below.
AUTH_MGMT_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",    "TCD"):    {"PFCG"},
    ("S_USER_PRO", "ACTVT"): {"01", "02"},
    ("S_USER_AGR", "ACTVT"): {"01", "02"},
    ("S_USER_AUT", "ACTVT"): {"01", "02"},
    ("S_USER_GRP", "ACTVT"): {"01", "02"},
    ("S_USER_SYS", "ACTVT"): {"01", "02"},
    ("S_USER_TCD", "ACTVT"): {"01", "02"},
    ("S_USER_VAL", "ACTVT"): {"01", "02"},
    ("S_DEVELOP",  "OBJTYPE"): {"ACGR"},
    ("S_USER_ADM", "ACTVT"): {"01"},
}

# Criteria for RSCDOK99 program permission check (control MA-RSCDOK99-01).
# An AGR_NAME qualifies only when it satisfies ALL criteria below simultaneously
# (i.e., it has BOTH S_PROGRAM/P_GROUP=RSCDOK99 AND S_PROGRAM/P_ACTION=SUB).
# Wildcard "*" in LOW or HIGH also qualifies for that criterion.
# Structure: list of (OBJECT, FIELD, qualifying_values)
RSCDOK99_PERMISSION_CRITERIA: list[tuple[str, str, set[str]]] = [
    ("S_PROGRAM", "P_GROUP",  {"RSCDOK99"}),
    ("S_PROGRAM", "P_ACTION", {"SUB"}),
]

# Criteria for data-management permission check (control MA-DATAMGMT-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# Special: {"*"} means ANY non-empty LOW or HIGH value qualifies for that criterion
# (used for S_TABU_NAM/TABLE where any table name is considered sensitive).
# OR logic: any single matching row qualifies the AGR_NAME.
DATA_MGMT_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",    "TCD"):        {"SE16", "SM30", "SM31", "SE16N", "SE17", "SM38", "SE37"},
    ("S_TABU_DIS", "ACTVT"):      {"01", "02", "06"},
    ("S_TABU_NAM", "ACTVT"):      {"01", "02", "06"},
    ("S_TABU_NAM", "TABLE"):      {"*"},   # any table name qualifies
    ("S_TABU_CLI", "CLIIDMAINT"): {"X"},
    ("S_DATASET",  "ACTVT"):      {"06", "34"},
}

# Criteria for transport/change-management permission check (control MA-TRANSPORT-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# OR logic: any single matching row qualifies the AGR_NAME.
# Note: "IMP*" is a literal SAP authorization value granting all import functions.
TRANSPORT_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",     "TCD"):        {"STMS", "STMS_IMPORT", "SCC4"},
    ("S_TABU_DIS",  "DICBERCLS"):  {"SS"},
    ("S_TABU_DIS",  "ACTVT"):      {"02"},
    ("S_TRANSPORT", "ACTVT"):      {"01", "02", "50", "60", "06", "43"},
    ("S_CTS_ADMI",  "CTS_ADMFCT"): {"IMPT", "IMPA", "IMP*"},
}

# Criteria for DEBUG permission check (control MA-DEBUG-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# Special: {"*"} means ANY non-empty LOW or HIGH value qualifies (S_PROGRAM/P_GROUP).
# OR logic: any single matching row qualifies the AGR_NAME.
DEBUG_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",    "TCD"):       {"SE38", "SA38", "SE80", "ST05"},
    ("S_DEVELOP",  "OBJTYPE"):   {"DEBUG"},
    ("S_DEVELOP",  "ACTVT"):     {"01", "02"},
    ("S_PROGRAM",  "P_ACTION"):  {"SUB"},
    ("S_PROGRAM",  "P_GROUP"):   {"*"},   # any program group qualifies
    ("S_ADMI_FCD", "S_ADMI_FCD"): {"PADM"},
}

# Criteria for job-management permission check (control MA-JOBMGMT-01).
# Maps (OBJECT, FIELD) -> set of qualifying LOW/HIGH values.
# Special: {"*"} means ANY non-empty LOW or HIGH value qualifies (S_BTCH_NAM/BTCUNAME).
# OR logic: any single matching row qualifies the AGR_NAME.
JOB_MGMT_PERMISSION_CRITERIA: dict[tuple[str, str], set[str]] = {
    ("S_TCODE",     "TCD"):      {"SE37", "SE36", "SM36", "SE35", "SE30", "SE34", "SHDB", "SM36WIZ"},
    ("S_BTCH_ADM",  "BTCADMIN"): {"Y"},
    ("S_BTCH_JOB",  "JOBACTION"): {"DELE", "RELE", "PROT"},
    ("S_BTCH_NAM",  "BTCUNAME"): {"*"},   # any BTCUNAME value qualifies
    ("S_BTCH_MONI", "BSCAKTI"):  {"DELE", "RELE"},
}

STRONG_PERMISSION_PROFILES: tuple[str, ...] = (
    "SAP_ALL",
    "SAP_NEW",
    "S_ABAP_ALL",
    "S_RZL_ADMIN",
    "S_A.SYSTEM",
    "S_A.ADMIN",
    "A_S.CUSTOMIZ",
    "S_A.DEVELOP",
    "S_A.USER",
    "S_USER_ALL",
)

SAP_APP_RSPARAM_RULES = [
    # (control_id, parameter_name, expected_value, rule_type, message)
    ("MA-PWD-01", "login/min_password_lng", 8, "minimum", "אורך סיסמה מינימלי חייב להיות לפחות 8 תווים"),
    ("MA-PWD-02", "login/fails_to_user_lock", 6, "maximum", "נעילת משתמש לאחר ניסיונות כושלים חייבת להיות לכל היותר 6"),
    ("MA-PWD-03", "login/failed_user_auto_unlock", 0, "maximum", "ביטול נעילה אוטומטי לאחר כישלון חייב להיות מבוטל (0)"),
    ("MA-PWD-04", "login/password_expiration_time", 90, "maximum", "תקופת תפוגת סיסמה חייבת להיות לכל היותר 90 ימים"),
    ("MA-PWD-05", "login/password_history_size", 5, "minimum", "היסטוריית סיסמאות חייבת לכלול לפחות 5 ערכים"),
    ("MA-PWD-06", "login/no_automatic_user_sapstar", 1, "minimum", "פרמטר SAP* האוטומטי חייב להיות מבוטל (1)"),
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
                    control_id="44",
                    category=AUDIT_CONTROL_DEFINITIONS["44"]["category"],
                    risk_level=AUDIT_CONTROL_DEFINITIONS["44"]["risk_level"],
                    check_type=AUDIT_CONTROL_DEFINITIONS["44"]["check_type"],
                    description=AUDIT_CONTROL_DEFINITIONS["44"]["description"],
                    actual_value=import_user,
                    expected_value="משתמש מורשה",
                    status="עם ממצא",
                    full_description=f"טרנספורט {trkorr} הועבר בתאריך {as4date or '-'} על ידי המשתמש {import_user}, שאינו מורשה.",
                )
            )

    return issues


def build_strong_profile_issues(
    profile: str | None,
    rows: list[dict[str, Any]],
) -> list[ValidationIssue]:
    """Detect users with strong system profiles from UST04 rows."""
    if profile != "UST04" or not rows:
        return []

    strong_profiles = {normalize_name(value) for value in STRONG_PERMISSION_PROFILES}
    control_meta = AUDIT_CONTROL_DEFINITIONS.get("MA-PERM-01", {})
    issues: list[ValidationIssue] = []

    for row_index, row in enumerate(rows, start=1):
        normalized_row = {
            normalize_name(key): value
            for key, value in row.items()
            if not str(key).startswith("__")
        }

        user_name = ""
        for candidate in get_column_aliases("BNAME"):
            if candidate in normalized_row:
                user_name = str(normalized_row.get(candidate, "")).strip().upper()
                if user_name:
                    break

        profile_name = ""
        for candidate in get_column_aliases("PROFILE"):
            if candidate in normalized_row:
                profile_name = str(normalized_row.get(candidate, "")).strip().upper()
                if profile_name:
                    break

        if not user_name or not profile_name or profile_name not in strong_profiles:
            continue

        issues.append(
            ValidationIssue(
                row_number=row_index,
                column_name="PROFILE",
                message=f"זוהה פרופיל חזק {profile_name} למשתמש {user_name}",
                source_file=str(row.get("__source_file", "")),
                control_id="MA-PERM-01",
                category=control_meta.get("category", ""),
                risk_level=control_meta.get("risk_level", ""),
                check_type=control_meta.get("check_type", ""),
                description=control_meta.get("description", ""),
                actual_value=user_name,
                expected_value=profile_name,
                status="עם ממצא",
                full_description=f"למשתמש {user_name} הוקצה הפרופיל החזק {profile_name}.",
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


def _resolve_value_by_alias_priority(row: dict[str, Any]) -> object | None:
    """Resolve VALUE-like columns by priority, preferring first non-empty value."""
    normalized_map = {normalize_name(column): column for column in row.keys()}
    fallback_value: object | None = None

    for alias in get_column_aliases("VALUE"):
        if alias not in normalized_map:
            continue
        raw_value = row.get(normalized_map[alias])
        if fallback_value is None:
            fallback_value = raw_value
        if raw_value is None:
            continue
        if isinstance(raw_value, str) and not raw_value.strip():
            continue
        return raw_value

    return fallback_value


def _evaluate_rsparam_policy(rows: list[dict[str, Any]]) -> list[ValidationIssue]:
    """Evaluate SAP APP (ABAP) password policy parameters from RSPARAM / TPFET data."""
    issues: list[ValidationIssue] = []
    param_map: dict[str, tuple[int, object]] = {}

    for row_number, row in enumerate(rows, start=1):
        param_column = _find_column_name(row, ("PARAMETER", "NAME"))
        if not param_column:
            continue

        resolved_value = _resolve_value_by_alias_priority(row)
        if resolved_value is None:
            continue

        param_map[normalize_text(row[param_column])] = (row_number, resolved_value)

    for control_id, param_name, expected, rule_type, message in SAP_APP_RSPARAM_RULES:
        control_meta = AUDIT_CONTROL_DEFINITIONS.get(control_id, {})
        if param_name not in param_map:
            issues.append(
                ValidationIssue(
                    row_number=0,
                    column_name=param_name,
                    message=f"לא נמצא פרמטר נדרש: {param_name}",
                    control_id=control_id,
                    category=control_meta.get("category", ""),
                    risk_level=control_meta.get("risk_level", ""),
                    check_type=control_meta.get("check_type", ""),
                    description=control_meta.get("description", message),
                    actual_value="-",
                    expected_value=str(expected),
                    status="עם ממצא",
                    full_description=f"הפרמטר {param_name} לא נמצא בקובץ המדיניות.",
                )
            )
            continue

        row_number, actual_value = param_map[param_name]
        if not _compare_values(actual_value, expected, rule_type):
            issues.append(
                ValidationIssue(
                    row_number=row_number,
                    column_name=param_name,
                    message=message,
                    control_id=control_id,
                    category=control_meta.get("category", ""),
                    risk_level=control_meta.get("risk_level", ""),
                    check_type=control_meta.get("check_type", ""),
                    description=control_meta.get("description", message),
                    actual_value=str(actual_value),
                    expected_value=str(expected),
                    status="עם ממצא",
                    full_description=f"הערך בפועל עבור {param_name} הוא {actual_value}, בעוד שהערך המצופה הוא {expected}.",
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


def get_audit_control_definition(control_id: str) -> dict[str, str]:
    return dict(AUDIT_CONTROL_DEFINITIONS.get(control_id, {}))


def get_profile_audit_controls(profile: str | None) -> list[str]:
    return list(PROFILE_AUDIT_CONTROLS.get(profile or "", []))
