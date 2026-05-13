from dataclasses import dataclass
from pathlib import Path


# ---------------------------------------------------------------------------
# Default control IDs pre-selected in the IPE evidence tag dialog per slot.
# A single screenshot can be linked to multiple controls (multi-select).
# ---------------------------------------------------------------------------
SLOT_DEFAULT_CONTROLS: dict[str, list[str]] = {
    "USR02": [
        "MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06",
        "MA-PERM-01", "MA-REVIEW-01", "MA-SOD-01",
    ],
    "ADR6_USR21": ["MA-REVIEW-01"],
    "AGR_USERS": [
        "MA-USRMGMT-01", "MA-AUTHMGMT-01", "MA-RSCDOK99-01",
        "MA-DATAMGMT-01", "MA-TRANSPORT-01", "MA-DEBUG-01", "MA-JOBMGMT-01",
        "MA-SOD-01",
    ],
    "AGR_1251": [
        "MA-USRMGMT-01", "MA-AUTHMGMT-01", "MA-RSCDOK99-01",
        "MA-DATAMGMT-01", "MA-TRANSPORT-01", "MA-DEBUG-01", "MA-JOBMGMT-01",
    ],
    "AGR_1252": ["MA-USRMGMT-01", "MA-AUTHMGMT-01", "MA-TRANSPORT-01"],
    "AGR_DEFINE": ["MA-USRMGMT-01", "MA-AUTHMGMT-01"],
    "UST04": ["MA-PERM-01"],
    "RSPARAM": [
        "MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06",
    ],
    "TPFET": [
        "MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06",
    ],
    "E070": ["MA-SOD-01"],
    "T000": ["MA-TRANSPORT-01"],
    "STMS": ["MA-TRANSPORT-01", "MC-44"],
}

# Human-readable labels for every control ID used in the tag dialog.
CONTROL_LABELS: dict[str, str] = {
    "MA-PWD-01": "MA-PWD-01 — מדיניות סיסמאות (1)",
    "MA-PWD-02": "MA-PWD-02 — מדיניות סיסמאות (2)",
    "MA-PWD-03": "MA-PWD-03 — מדיניות סיסמאות (3)",
    "MA-PWD-04": "MA-PWD-04 — מדיניות סיסמאות (4)",
    "MA-PWD-05": "MA-PWD-05 — מדיניות סיסמאות (5)",
    "MA-PWD-06": "MA-PWD-06 — מדיניות סיסמאות (6)",
    "MA-PERM-01": "MA-PERM-01 — פרופילים חזקים (SAP_ALL / SAP_NEW)",
    "MA-USRMGMT-01": "MA-USRMGMT-01 — הרשאות ניהול משתמשים",
    "MA-AUTHMGMT-01": "MA-AUTHMGMT-01 — הרשאות ניהול הרשאות",
    "MA-RSCDOK99-01": "MA-RSCDOK99-01 — הרשאות לתוכנית RSCDOK99",
    "MA-DATAMGMT-01": "MA-DATAMGMT-01 — הרשאות לניהול נתונים",
    "MA-TRANSPORT-01": "MA-TRANSPORT-01 — הרשאות להעברת שינויים",
    "MA-DEBUG-01": "MA-DEBUG-01 — הרשאות לשימוש ב-DEBUG",
    "MA-JOBMGMT-01": "MA-JOBMGMT-01 — הרשאות לניהול ג'ובים",
    "MA-REVIEW-01": "MA-REVIEW-01 — השלמת סקירת משתמשים",
    "MA-SOD-01": "MA-SOD-01 — הפרדת תפקידים (SoD) / מפתחים בייצור",
    "MC-44": "MC-44 — משתמשים מורשים ל-Import בסביבת ייצור",
}

# Grouped order for the tag dialog checkboxes.
CONTROL_GROUPS: list[tuple[str, list[str]]] = [
    (
        "MA - ניהול גישה",
        [
            "MA-PWD-01", "MA-PWD-02", "MA-PWD-03", "MA-PWD-04", "MA-PWD-05", "MA-PWD-06",
            "MA-PERM-01",
            "MA-USRMGMT-01", "MA-AUTHMGMT-01", "MA-RSCDOK99-01",
            "MA-DATAMGMT-01", "MA-TRANSPORT-01", "MA-DEBUG-01", "MA-JOBMGMT-01",
            "MA-REVIEW-01", "MA-SOD-01",
        ],
    ),
    (
        "MC - ניהול שינויים",
        ["MA-TRANSPORT-01", "MC-44"],
    ),
]


@dataclass
class AppConfig:
    input_dir: Path
    output_dir: Path
    supported_extensions: tuple[str, ...] = (".txt", ".csv", ".xlsx", ".xlsm")

    @classmethod
    def default(cls, base_dir: Path | None = None) -> "AppConfig":
        root_dir = base_dir or Path.cwd()
        data_dir = root_dir / "data"
        return cls(
            input_dir=data_dir / "input",
            output_dir=data_dir / "output",
        )
