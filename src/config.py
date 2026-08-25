from __future__ import annotations

import shutil
import sys
from dataclasses import dataclass
from pathlib import Path


def is_frozen() -> bool:
    """True when running from a PyInstaller (or similar) bundled executable."""
    return bool(getattr(sys, "frozen", False))


def get_install_root() -> Path:
    """Writable application root: folder of the EXE when frozen, else project root."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_bundle_root() -> Path:
    """Read-only bundled resources (_MEIPASS when frozen, else project root)."""
    if is_frozen():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_path(*parts: str | Path) -> Path:
    """Path to a bundled resource file (logo, seed JSON, etc.)."""
    return get_bundle_root().joinpath(*parts)


def knowledge_base_dir(install_root: Path | None = None) -> Path:
    """Writable knowledge_base under the install root (after seed)."""
    root = install_root or get_install_root()
    return root / "data" / "knowledge_base"


def ensure_runtime_data(install_root: Path | None = None) -> Path:
    """Create writable data dirs and seed knowledge_base from the bundle if missing.

    Does not overwrite existing files (e.g. client ``system_settings.json`` or
    an already-edited catalog).
    """
    root = install_root or get_install_root()
    data_dir = root / "data"
    for sub in (
        "input",
        "output",
        "evidence",
        "compensating_controls",
        "knowledge_base",
    ):
        (data_dir / sub).mkdir(parents=True, exist_ok=True)

    bundle_kb = get_bundle_root() / "data" / "knowledge_base"
    dest_kb = data_dir / "knowledge_base"
    if bundle_kb.is_dir():
        for name in ("controls_catalog.json", "field_labels.json"):
            src = bundle_kb / name
            dest = dest_kb / name
            if src.is_file() and not dest.exists():
                shutil.copy2(src, dest)
    return root


# ---------------------------------------------------------------------------
# Default control IDs pre-selected in the IPE evidence tag dialog per slot.
# A single screenshot can be linked to multiple controls (multi-select).
# ---------------------------------------------------------------------------
SLOT_DEFAULT_CONTROLS: dict[str, list[str]] = {
    "USR02": [
        "MA2-2_AYALON_6", "MA1-1_AYALON_5", "MA3-3_AYALON_14",
        "MA1-1&MA7-17_AYALON_2", "MC5-23_AYALON_48", "MA5.1-13_AYALON_24",
        "MA7-17_AYALON_30",
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
        root_dir = base_dir or get_install_root()
        data_dir = root_dir / "data"
        return cls(
            input_dir=data_dir / "input",
            output_dir=data_dir / "output",
        )
