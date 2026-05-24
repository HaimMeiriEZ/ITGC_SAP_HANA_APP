from dataclasses import dataclass
from pathlib import Path


# ---------------------------------------------------------------------------
# Default control IDs pre-selected in the IPE evidence tag dialog per slot.
# A single screenshot can be linked to multiple controls (multi-select).
# ---------------------------------------------------------------------------
SLOT_DEFAULT_CONTROLS: dict[str, list[str]] = {
    "USR02": [
        "MA2-2_AYALON_6", "MA1-1_AYALON_5", "MA3-3_AYALON_14",
        "MA1-1&MA7-17_AYALON_2", "MC5-23_AYALON_48", "MA5.1-13_AYALON_24",
    ],
    "ADR6_USR21": ["MA1-1&MA7-17_AYALON_2"],
    "AGR_USERS": [
        "MA1-1_AYALON_10", "MA1-1_AYALON_11", "MA1-1_AYALON_12",
        "MA1-1_AYALON_16", "MA1-1_AYALON_43", "MA1-1_AYALON_45", "MA1-1_AYALON_67",
        "MC5-23_AYALON_48", "MA7-17_AYALON_30",
    ],
    "AGR_1251": [
        "MA1-1_AYALON_10", "MA1-1_AYALON_11", "MA1-1_AYALON_12",
        "MA1-1_AYALON_16", "MA1-1_AYALON_43", "MA1-1_AYALON_45", "MA1-1_AYALON_67",
        "MA7-17_AYALON_30",
    ],
    "AGR_1252": ["MA1-1_AYALON_10", "MA1-1_AYALON_11", "MA1-1_AYALON_43"],
    "AGR_DEFINE": ["MA1-1_AYALON_10", "MA1-1_AYALON_11"],
    "UST04": ["MA3-3_AYALON_14"],
    "USH04": ["MA3-3_AYALON_14", "MA5.3-13_AYALON_25"],
    "RSPARAM": ["MA2-2_AYALON_6", "MA1-1_AYALON_5"],
    "TPFET": ["MA2-2_AYALON_6", "MA1-1_AYALON_5"],
    "E070": ["MC5-23_AYALON_48"],
    "T000": ["MA1-1_AYALON_43"],
    "STMS": ["MA1-1_AYALON_43", "MC7-25_AYALON_44"],
}

# Human-readable labels for every control ID used in the tag dialog.
CONTROL_LABELS: dict[str, str] = {
    "MA2-2_AYALON_6": "MA2-2_AYALON_6 — מדיניות סיסמאות",
    "MA1-1_AYALON_5": "MA1-1_AYALON_5 — משתמשי מערכת (SAP*, DDIC)",
    "MA3-3_AYALON_14": "MA3-3_AYALON_14 — פרופילים חזקים (SAP_ALL / SAP_NEW)",
    "MA1-1&MA7-17_AYALON_2": "MA1-1&MA7-17_AYALON_2 — השלמת סקירת משתמשים",
    "MC5-23_AYALON_48": "MC5-23_AYALON_48 — הפרדת תפקידים (SoD) / מפתחים בייצור",
    "MA5.1-13_AYALON_24": "MA5.1-13_AYALON_24 — משתמשים חדשים",
    "MA5.3-13_AYALON_25": "MA5.3-13_AYALON_25 — משתמשים מנויידים",
    "MA7-17_AYALON_30": "MA7-17_AYALON_30 — סקירת הרשאות משתמשים",
    "MA1-1_AYALON_10": "MA1-1_AYALON_10 — הרשאות ניהול משתמשים",
    "MA1-1_AYALON_11": "MA1-1_AYALON_11 — הרשאות ניהול הרשאות",
    "MA1-1_AYALON_12": "MA1-1_AYALON_12 — הרשאות לתוכנית RSCDOK99",
    "MA1-1_AYALON_16": "MA1-1_AYALON_16 — הרשאות לניהול נתונים",
    "MA1-1_AYALON_43": "MA1-1_AYALON_43 — הרשאה להעברת שינויים",
    "MA1-1_AYALON_45": "MA1-1_AYALON_45 — הרשאות לשימוש ב-DEBUG",
    "MA1-1_AYALON_67": "MA1-1_AYALON_67 — הרשאות לניהול ג'ובים",
    "MC7-25_AYALON_44": "MC7-25_AYALON_44 — משתמשים מורשים ל-Import בסביבת ייצור",
}

# Grouped order for the tag dialog checkboxes.
CONTROL_GROUPS: list[tuple[str, list[str]]] = [
    (
        "MA - ניהול גישה",
        [
            "MA2-2_AYALON_6",
            "MA1-1_AYALON_5",
            "MA3-3_AYALON_14",
            "MA1-1&MA7-17_AYALON_2",
            "MA5.1-13_AYALON_24",
            "MA5.3-13_AYALON_25",
            "MA7-17_AYALON_30",
            "MA1-1_AYALON_10", "MA1-1_AYALON_11", "MA1-1_AYALON_12",
            "MA1-1_AYALON_16", "MA1-1_AYALON_43", "MA1-1_AYALON_45", "MA1-1_AYALON_67",
            "MC5-23_AYALON_48",
        ],
    ),
    (
        "MC - ניהול שינויים",
        ["MA1-1_AYALON_43", "MC7-25_AYALON_44"],
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
