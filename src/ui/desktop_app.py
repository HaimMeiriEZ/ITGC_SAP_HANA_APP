import json
import os
import re
import subprocess
import sys
import copy
from typing import Any, Sequence
from datetime import datetime, date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from PySide6.QtCore import QCoreApplication, QDate, QObject, QThread, Qt, Signal, Slot
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QStyledItemDelegate,
    QSizePolicy,
    QTabWidget,
    QApplication,
    QFileDialog,
    QComboBox,
    QDateEdit,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QPlainTextEdit,
    QTextEdit,
    QProgressBar,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from src.config import AppConfig
from src.models.validation_result import ValidationIssue
from src.pipeline import process_file
from src.readers.excel_reader import ExcelFileReader
from src.readers.text_reader import TextFileReader
from src.reporting.excel_report import ExcelReportWriter
from src.validators.spec_rules import (
    AUTH_MGMT_PERMISSION_CRITERIA,
    DATA_MGMT_PERMISSION_CRITERIA,
    DEBUG_PERMISSION_CRITERIA,
    JOB_MGMT_PERMISSION_CRITERIA,
    RSCDOK99_PERMISSION_CRITERIA,
    SAP_APP_RSPARAM_RULES,
    TRANSPORT_PERMISSION_CRITERIA,
    USER_MGMT_PERMISSION_CRITERIA,
    get_audit_control_definition,
    get_column_aliases,
    get_profile_audit_controls,
)


def get_qt_app() -> QApplication:
    if "unittest" in sys.modules and "QT_QPA_PLATFORM" not in os.environ:
        os.environ["QT_QPA_PLATFORM"] = "offscreen"

    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication(sys.argv)
        app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        app.setFont(QFont("Segoe UI", 9))
    return app


class _RightAlignDelegate(QStyledItemDelegate):
    """Forces physical-right alignment on all cells, regardless of RTL layout direction."""

    def initStyleOption(self, option: Any, index: Any) -> None:  # type: ignore[override]
        super().initStyleOption(option, index)
        option.displayAlignment = (
            Qt.AlignmentFlag.AlignAbsolute
            | Qt.AlignmentFlag.AlignRight
            | Qt.AlignmentFlag.AlignVCenter
        )


class SortableTableWidgetItem(QTableWidgetItem):
    SORT_ROLE = Qt.ItemDataRole.UserRole + 2

    def __lt__(self, other: object) -> bool:
        if isinstance(other, QTableWidgetItem):
            self_sort_value = self.data(self.SORT_ROLE)
            other_sort_value = other.data(self.SORT_ROLE)
            if self_sort_value is not None or other_sort_value is not None:
                left = "" if self_sort_value is None else str(self_sort_value)
                right = "" if other_sort_value is None else str(other_sort_value)
                return left < right
            return super().__lt__(other)
        return NotImplemented


class SlotValidationWorker(QObject):
    succeeded = Signal(str, list, object)
    failed = Signal(str, list, str)
    finished = Signal()

    def __init__(
        self,
        slot_key: str,
        file_paths: list[str],
        input_files_dict: dict[str, list[str | Path]],
        required_columns: list[str],
        output_dir: Path,
        authorized_users: list[str],
    ) -> None:
        super().__init__()
        self.slot_key = slot_key
        self.file_paths = file_paths
        self.input_files_dict = input_files_dict
        self.required_columns = required_columns
        self.output_dir = output_dir
        self.authorized_users = authorized_users

    @Slot()
    def run(self) -> None:
        try:
            result = process_file(
                input_files=self.input_files_dict,
                required_columns=self.required_columns,
                output_dir=self.output_dir,
                source_name_override=self.slot_key,
                authorized_users=self.authorized_users,
            )
            self.succeeded.emit(self.slot_key, list(self.file_paths), result)
        except Exception as error:
            self.failed.emit(self.slot_key, list(self.file_paths), str(error))
        finally:
            self.finished.emit()


class ValidationDesktopApp(QMainWindow):
    USER_PREVIEW_SLOTS = {"USR02", "ADR6_USR21"}

    MULTI_FILE_SLOTS = {
        "USR02",
        "ADR6_USR21",
        "AGR_USERS",
        "AGR_1251",
        "AGR_1252",
        "AGR_DEFINE",
        "UST04",
        "E070",
        "STMS",
    }

    USER_PREVIEW_COLUMN_DEFINITIONS = [
        {"field": "MANDT", "formal": "CLIENT", "technical": "MANDT", "source": "USR02", "default": True, "width": 90},
        {"field": "WORK_ENVIRONMENT", "formal": "סביבת עבודה", "technical": "WORK_ENVIRONMENT", "source": "הגדרות מערכת", "default": True, "width": 210},
        {"field": "BNAME", "formal": "משתמש", "technical": "BNAME", "source": "USR02", "default": True, "width": 130},
        {"field": "NAME_FIRST", "formal": "שם פרטי", "technical": "NAME_FIRST", "source": "USER_ADDR", "default": True, "width": 120},
        {"field": "NAME_LAST", "formal": "שם משפחה", "technical": "NAME_LAST", "source": "USER_ADDR", "default": True, "width": 120},
        {"field": "NAME_TEXTC", "formal": "שם מלא", "technical": "NAME_TEXTC", "source": "USER_ADDR", "default": True, "width": 180},
        {"field": "COMPANY", "formal": "חברה", "technical": "COMPANY", "source": "USER_ADDR", "default": True, "width": 150},
        {"field": "SMTP_ADDR", "formal": 'דוא"ל', "technical": "SMTP_ADDR", "source": "ADR6", "default": True, "width": 200},
        {"field": "STATUS", "formal": "סטטוס משתמש", "technical": "STATUS", "source": "USR02", "default": True, "width": 120},
        {"field": "ADDRNUMBER", "formal": "מספר כתובת", "technical": "ADDRNUMBER", "source": "USER_ADDR / ADR6", "default": True, "width": 120},
        {"field": "PERSNUMBER", "formal": "מספר פרסונה", "technical": "PERSNUMBER", "source": "USER_ADDR / ADR6", "default": True, "width": 120},
        {"field": "TRDAT", "formal": "תאריך כניסה אחרון", "technical": "TRDAT", "source": "USR02", "default": True, "width": 140},
        {"field": "LTIME", "formal": "שעת כניסה אחרונה", "technical": "LTIME", "source": "USR02", "default": True, "width": 140},
        {"field": "PWDINITIAL", "formal": "סיסמה ראשונית", "technical": "PWDINITIAL", "source": "USR02", "default": True, "width": 120},
        {"field": "PWDCHGDATE", "formal": "תאריך שינוי סיסמה", "technical": "PWDCHGDATE", "source": "USR02", "default": True, "width": 145},
        {"field": "PWDSETDATE", "formal": "תאריך הגדרת סיסמה", "technical": "PWDSETDATE", "source": "USR02", "default": True, "width": 145},
        {"field": "DEPARTMENT", "formal": "מחלקה", "technical": "DEPARTMENT", "source": "USER_ADDR", "default": True, "width": 150},
        {"field": "GLTGV", "formal": "תקף מתאריך", "technical": "GLTGV", "source": "USR02", "default": True, "width": 120},
        {"field": "GLTGB", "formal": "תקף עד תאריך", "technical": "GLTGB", "source": "USR02", "default": True, "width": 120},
        {"field": "USTYP", "formal": "סוג משתמש", "technical": "USTYP", "source": "USR02", "default": True, "width": 110},
        {"field": "LOCNT", "formal": "מספר ניסיונות כניסה כושלים", "technical": "LOCNT", "source": "USR02", "default": True, "width": 125},
        {"field": "OCOD1", "formal": "סיסמה", "technical": "OCOD1", "source": "USR02", "default": True, "width": 130},
        {"field": "PASSCODE", "formal": "ערך Hash של סיסמה", "technical": "PASSCODE", "source": "USR02", "default": True, "width": 220},
        {"field": "PWDSALTEDHASH", "formal": "ערך Hash מוצפן של סיסמה", "technical": "PWDSALTEDHASH", "source": "USR02", "default": True, "width": 220},
        {"field": "SECURITY_POLICY", "formal": "מדיניות אבטחה", "technical": "SECURITY_POLICY", "source": "USR02", "default": True, "width": 160},
        {"field": "REVIEW_STATUS", "formal": "בוצעה סקירה", "technical": "REVIEW_STATUS", "source": "סוקר", "default": True, "width": 165},
        {"field": "FINDINGS_DESCRIPTION", "formal": "תיאור ממצאים", "technical": "FINDINGS_DESCRIPTION", "source": "מערכת", "default": True, "width": 280},
        {"field": "TECH_REVIEW_NOTES", "formal": "הערות סוקר גורם טכני", "technical": "TECH_REVIEW_NOTES", "source": "סוקר טכני", "default": True, "width": 240},
        {"field": "BUS_REVIEW_NOTES", "formal": "הערות סוקר גורם עסקי", "technical": "BUS_REVIEW_NOTES", "source": "סוקר עסקי", "default": True, "width": 240},
        {"field": "UFLAG", "formal": "קוד נעילה", "technical": "UFLAG", "source": "USR02", "default": False, "width": 100},
    ]
    DEFAULT_USER_PREVIEW_COLUMNS = [
        column["field"]
        for column in USER_PREVIEW_COLUMN_DEFINITIONS
        if bool(column.get("default"))
    ]
    CURRENT_USER_PREVIEW_SETTINGS_VERSION = 7
    USER_PREVIEW_SETTINGS_MIGRATIONS = {
        2: ["PWDINITIAL", "PWDCHGDATE", "PWDSETDATE"],
        3: ["DEPARTMENT", "GLTGV", "GLTGB", "USTYP", "LOCNT", "OCOD1", "PASSCODE", "PWDSALTEDHASH", "SECURITY_POLICY"],
        4: ["REVIEW_STATUS", "REVIEW_NOTES"],
        5: ["FINDINGS_DESCRIPTION"],
        6: ["TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES"],
        7: ["WORK_ENVIRONMENT"],
    }
    USER_PREVIEW_FILTER_OPTIONS = [
        ("all", "כלל האוכלוסייה"),
        ("active", "פעילים בתקופה הנבדקת"),
        ("inactive", "לא פעילים בתקופה הנבדקת"),
    ]
    REVIEW_STATUS_OPTIONS = ["טרם נבדק", "נבדק - תקין", "נבדק - לא תקין"]
    DEFAULT_REVIEW_STATUS = "טרם נבדק"
    REVIEWED_STATUSES = {"נבדק - תקין", "נבדק - לא תקין"}
    REVIEW_COMPLETION_CONTROL_ID = "MA-REVIEW-01"
    USER_TYPE_RULES = {
        "Dialog": ["A"],
        "System": ["B"],
        "Communication": ["C"],
        "Service": ["S"],
        "Reference": ["L"],
    }
    USER_PREVIEW_DATE_FIELDS = {"TRDAT", "PWDCHGDATE", "PWDSETDATE", "GLTGV", "GLTGB"}
    EXPORT_REVIEW_FIELDS = [
        "MANDT", "WORK_ENVIRONMENT", "BNAME", "NAME_TEXTC", "SMTP_ADDR", "STATUS", "USTYP",
        "GLTGV", "GLTGB", "TRDAT", "PWDSETDATE", "PWDCHGDATE",
        "FINDINGS_DESCRIPTION", "REVIEW_STATUS", "TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES",
    ]
    WORK_ENVIRONMENT_OPTIONS = [
        ("FPD", "FPD - DEV - סביבת מפתחים"),
        ("PFT", "PFT - PRE PROD - סביבת קדם ייצור"),
        ("FPP", "FPP - PROD - סביבת ייצור"),
        ("FPQ", "FPQ - QA - סביבת בדיקות"),
    ]

    SLOT_DEFINITIONS = {
        "USR02": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.1 - Joiners / Movers / Leavers",
            "description": "משתמשים - מקור חובה לבדיקות גישה, סטטוס ותאריכי התחברות.",
            "expected_file": "usr02_100.txt",
            "required": True,
        },
        "ADR6_USR21": {
            "label": "ADR6 / USER_ADDR",
            "domain": "MA - ניהול גישה",
            "sub_category": "1.1 - Joiners / Movers / Leavers",
            "description": "ניתן להזין קובצי ADR6 או USER_ADDR או את שניהם יחד לצורך העשרת נתוני המשתמשים מתוך USR02.",
            "expected_file": "adr6.txt או user_addr.txt",
            "required": False,
        },
        "AGR_USERS": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.2 - סקר הרשאות תקופתי",
            "description": "רולים-משתמשים - מיפוי המשתמשים לרולים במערכת.",
            "expected_file": "agr_users_100.txt",
            "required": True,
        },
        "AGR_1251": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.2 - סקר הרשאות תקופתי",
            "description": "רולים-אובייקטי הרשאה - זיהוי אובייקטי הרשאות רגישים.",
            "expected_file": "agr_1251_100.txt",
            "required": True,
        },
        "UST04": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.2 - סקר הרשאות תקופתי",
            "description": "פרופילים-משתמשים - שיוך פרופילים ישיר למשתמשים.",
            "expected_file": "ust04.txt",
            "required": True,
        },
        "AGR_1252": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.2 - סקר הרשאות תקופתי",
            "description": "רולים-טרנזקציות - זיהוי גישות עסקיות וטרנזקציות.",
            "expected_file": "agr_1252_100.txt",
            "required": False,
        },
        "AGR_DEFINE": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.2 - סקר הרשאות תקופתי",
            "description": "רולים מורחב - מידע כללי על הגדרת הרול.",
            "expected_file": "agr_define.txt",
            "required": False,
        },
        "RSPARAM": {
            "domain": "MA - ניהול גישה",
            "sub_category": "1.3 - משתמשי-על ומדיניות סיסמאות",
            "description": "פרמטרים סיסטמאיים - פרמטרי אבטחה והקשחת מערכת.",
            "expected_file": "rsparam.xlsx",
            "required": True,
        },
        "TPFET": {
            "label": "TPFET / RZ10",
            "domain": "MA - ניהול גישה",
            "sub_category": "1.3 - משתמשי-על ומדיניות סיסמאות",
            "description": "פרמטרים סיסטמאיים נוספים, כולל פרופילי login כגון RZ10.",
            "expected_file": "rz10.txt",
            "required": False,
        },
        "E070": {
            "domain": "MC - ניהול שינויים",
            "sub_category": "2.1 - תיעוד ובקשות שינוי",
            "description": "רשימת שינויים - נתוני transport requests ושינויים בסביבה.",
            "expected_file": "e070_100.txt",
            "required": True,
        },
        "T000": {
            "domain": "MC - ניהול שינויים",
            "sub_category": "2.3 - הפרדת סביבות DEV/QA/PRD",
            "description": "לוג פעילות שינוי SCC4 - בקרות שינוי ברמת client.",
            "expected_file": "t000.txt",
            "required": False,
        },
        "STMS": {
            "domain": "MC - ניהול שינויים",
            "sub_category": "2.4 - Import לסביבת ייצור",
            "description": "רשימת שינויים שהועברה דרך SCC4 או STMS.",
            "expected_file": "stms.txt",
            "required": False,
        },
    }

    DOMAIN_DEFINITIONS: dict[str, Any] = {
        "MA - ניהול גישה": {
            "in_development": False,
            "sub_categories": [
                {
                    "key": "1.1 - Joiners / Movers / Leavers",
                    "description": "תהליכי הצטרפות, שינוי תפקיד ועזיבה — ווידוא שקיים אישור מנהל ושהגישה הוסרה בזמן.",
                },
                {
                    "key": "1.2 - סקר הרשאות תקופתי",
                    "description": "ביצוע סקר הרשאות תקופתי על ידי גורמים עסקיים (Business Owners) לווידוא נחיצות ההרשאות.",
                },
                {
                    "key": "1.3 - משתמשי-על ומדיניות סיסמאות",
                    "description": "ניהול משתמשי-על (BASIS, SAP*, DDIC) ובדיקת פרמטרי מדיניות סיסמאות במערכת.",
                },
            ],
        },
        "MC - ניהול שינויים": {
            "in_development": False,
            "sub_categories": [
                {
                    "key": "2.1 - תיעוד ובקשות שינוי",
                    "description": "תיעוד בקשת השינוי ואישור עסקי לפני הפיתוח.",
                },
                {
                    "key": "2.3 - הפרדת סביבות DEV/QA/PRD",
                    "description": "הפרדה בין סביבת הפיתוח (DEV), הבחינה (QA) והייצור (PRD).",
                },
                {
                    "key": "2.4 - Import לסביבת ייצור",
                    "description": "בחינה של תהליך ה-Import לסביבת הייצור ומי מורשה לבצע אותו (בדרך כלל צוות ה-Basis).",
                },
            ],
        },
        "MO - ניהול תפעולי": {
            "in_development": True,
            "sub_categories": [
                {
                    "key": "3.1 - ניטור תהליכי רקע (Batch Jobs)",
                    "description": "בודקים מה קורה אם ג'וב קריטי נכשל והאם יש בקרה על שינוי לוחות זמנים של ג'ובים.",
                },
                {
                    "key": "3.2 - גיבויים ושחזור (Backups)",
                    "description": "ווידוא שהגיבויים מבוצעים כסדרם ושיש שחזור תקופתי מוצלח (Restoration Test).",
                },
                {
                    "key": "3.3 - ניהול תקלות (Incidents)",
                    "description": "תהליך הטיפול בתקלות מערכת ותיעודן.",
                },
            ],
        },
    }

    SETTINGS_SECTION_DEPENDENCIES = {
        "user_review_period": set(),
        "authorized_stms_users": set(),
        "super_users": set(),
        "generic_users": set(),
        "critical_roles": set(),
        "critical_privileges": set(),
        "password_policy_defaults": set(),
        "file_mappings": set(),
        "inactive_days_threshold": set(),
    }

    def __init__(self, base_dir: Path | None = None) -> None:
        super().__init__()
        self.config = AppConfig.default(base_dir or Path.cwd())
        self.config.input_dir.mkdir(parents=True, exist_ok=True)
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.report_path: Path | None = None
        self.log_export_path: Path | None = None
        self.slot_widgets: dict[str, dict[str, Any]] = {}
        self.category_run_buttons: dict[str, QPushButton] = {}
        self.category_sections: dict[str, QGroupBox] = {}
        self.selected_slot_key: str | None = None
        self.load_history: list[str] = []
        self.summary_labels: dict[str, QLabel] = {}
        self.run_log_records: list[dict[str, Any]] = []
        self.audit_summary_records: dict[str, dict[str, Any]] = {}
        self.audit_details_by_control: dict[str, list[dict[str, Any]]] = {}
        self.permissions_summary_records: dict[str, dict[str, Any]] = {}
        self.permissions_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.agr_1251_cached_rows: list[dict[str, Any]] = []
        self.agr_users_cached_rows: list[dict[str, Any]] = []
        self.user_mgmt_summary_records: dict[str, dict[str, Any]] = {}
        self.user_mgmt_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.auth_mgmt_summary_records: dict[str, dict[str, Any]] = {}
        self.auth_mgmt_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.rscdok99_summary_records: dict[str, dict[str, Any]] = {}
        self.rscdok99_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.data_mgmt_summary_records: dict[str, dict[str, Any]] = {}
        self.data_mgmt_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.transport_summary_records: dict[str, dict[str, Any]] = {}
        self.transport_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.debug_summary_records: dict[str, dict[str, Any]] = {}
        self.debug_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.job_mgmt_summary_records: dict[str, dict[str, Any]] = {}
        self.job_mgmt_users_by_control: dict[str, list[dict[str, Any]]] = {}
        self.audit_findings_export_path: Path | None = None
        self.validation_thread: QThread | None = None
        self.validation_worker: SlotValidationWorker | None = None
        self._allow_user_preview_persistence = base_dir is not None or "unittest" not in sys.modules
        self.last_file_dialog_directory = self._load_last_file_dialog_directory()
        self._refreshing_user_preview = False
        self.user_preview_export_path: Path | None = None
        self.user_reviewer_state = self._load_user_reviewer_state()
        self.user_preview_visible_columns = self._load_user_preview_column_selection()
        self.system_settings_widgets: dict[str, Any] = {}
        self.system_settings_sections: dict[str, QGroupBox] = {}
        self.system_settings_unavailable_labels: dict[str, QLabel] = {}
        self.system_settings_file_mapping_order: list[str] = []

        self._configure_window()
        self._build_ui()
        self._load_system_settings_into_form(
            self._current_system_settings(),
            load_review_period=self._system_settings_path().exists(),
        )
        self._apply_system_settings_availability()

    @staticmethod
    def format_rtl_text(text: object) -> str:
        raw_text = "" if text is None else str(text)
        return re.sub(r"[\u2066\u2067\u2068\u2069\u200e\u200f]", "", raw_text)

    @staticmethod
    def format_ui_rtl_text(text: object) -> str:
        normalized_text = ValidationDesktopApp.format_rtl_text(text).strip()
        if normalized_text and re.search(r"[\u0590-\u05FF]", normalized_text):
            return f"\u202B{normalized_text}\u202C"
        return normalized_text

    def _configure_window(self) -> None:
        self.setWindowTitle(self.format_rtl_text("כלי להערכת בקרות ITGC בסביבת SAP HANA APP"))
        self.setMinimumSize(1180, 760)
        self.resize(1280, 860)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

    def _build_ui(self) -> None:
        central_widget = QWidget()
        central_widget.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        _title_container = QWidget()
        _title_container.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        _title_row = QHBoxLayout(_title_container)
        _title_row.setContentsMargins(0, 0, 0, 0)
        _title_row.setSpacing(0)
        self.app_title_label = QLabel("כלי להערכת בקרות ITGC בסביבת SAP HANA APP")
        self.app_title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #16325c;")
        _title_row.addStretch(1)
        _title_row.addWidget(self.app_title_label)
        main_layout.addWidget(_title_container)

        _header_container = QWidget()
        _header_container.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        _header_row = QHBoxLayout(_header_container)
        _header_row.setContentsMargins(0, 0, 0, 0)
        _header_row.setSpacing(0)
        self.header_label = QLabel("מסך בדיקת קלטי SAP HANA APP")
        self.header_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #16325c;")
        _header_row.addStretch(1)
        _header_row.addWidget(self.header_label)

        self.hint_label = QTextEdit()
        self.hint_label.setReadOnly(True)
        self.hint_label.setHtml(
            '<p dir="rtl" style="color: #4f5d73; margin: 0; padding: 0;">'
            "בחר קבצים לפי המשבצת המתאימה. כוכבית מציינת משבצת חובה. חובה לציין את תאריך ההפקה של הקבצים. ניתן להריץ בדיקה נפרדת לכל קבוצת קבצים בלי להמתין לטעינת כל הדוחות."
            "</p>"
        )
        self.hint_label.setFixedHeight(38)
        self.hint_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.hint_label.setStyleSheet(
            "background: transparent; border: none; padding: 0;"
        )

        self.work_environment_row = QWidget()
        self.work_environment_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        work_environment_layout = QHBoxLayout(self.work_environment_row)
        work_environment_layout.setContentsMargins(0, 0, 0, 0)
        work_environment_layout.setSpacing(8)
        work_environment_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.work_environment_label = QLabel(self.format_ui_rtl_text("סביבת עבודה:"))
        self.work_environment_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.work_environment_combo = QComboBox()
        self.work_environment_combo.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.work_environment_combo.setMinimumWidth(260)
        for env_code, env_label in self.WORK_ENVIRONMENT_OPTIONS:
            self.work_environment_combo.addItem(self.format_rtl_text(env_label), env_code)
        self.work_environment_combo.currentIndexChanged.connect(self._persist_work_environment_selection)

        work_environment_layout.addWidget(self.work_environment_label)
        work_environment_layout.addWidget(self.work_environment_combo)
        work_environment_layout.addStretch(1)

        self.actions_row = QWidget()
        self.actions_row.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.actions_row.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        buttons_layout = QHBoxLayout(self.actions_row)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        buttons_layout.addStretch(1)

        self.clear_last_load_button = QPushButton(self.format_ui_rtl_text("נקה טעינה אחרונה"))
        self.clear_last_load_button.clicked.connect(self.clear_last_loaded_slot)
        # ...existing code for the rest of _build_ui...
        buttons_layout.addWidget(self.clear_last_load_button)

        self.clear_button = QPushButton(self.format_ui_rtl_text("נקה מסך"))
        self.clear_button.clicked.connect(self.clear_results)
        buttons_layout.addWidget(self.clear_button)

        self.export_log_button = QPushButton(self.format_ui_rtl_text("ייצוא רישום לאקסל"))
        self.export_log_button.clicked.connect(lambda: self.export_run_log_to_excel(open_after_export=True))
        buttons_layout.addWidget(self.export_log_button)

        self.output_button = QPushButton(self.format_ui_rtl_text("פתח תיקיית פלט"))
        self.output_button.clicked.connect(self.open_output_folder)
        buttons_layout.addWidget(self.output_button)

        self.report_button = QPushButton(self.format_ui_rtl_text("פתח דוח אקסל"))
        self.report_button.setEnabled(False)
        self.report_button.clicked.connect(self.open_report)
        buttons_layout.addWidget(self.report_button)

        self.tabs = QTabWidget()
        self.tabs.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.tabs.setDocumentMode(True)
        self.tabs.setTabPosition(QTabWidget.TabPosition.North)
        self.tabs.setMovable(False)
        self.tabs.setStyleSheet(
            """
            QTabBar::tab {
                background-color: #e9eef7;
                color: #16325c;
                border: 1px solid #b7c4d8;
                border-bottom: none;
                padding: 6px 12px;
                margin-left: 2px;
                min-width: 150px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background-color: #6d002f;
                color: white;
            }
            QTabWidget::pane {
                border: 1px solid #c7cfda;
                top: -1px;
                background: #f5f7fb;
            }
            """
        )

        self.intake_tab = QWidget()
        self.intake_tab.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.intake_layout = QVBoxLayout(self.intake_tab)
        self.intake_layout.setContentsMargins(8, 8, 8, 8)
        self.intake_layout.setSpacing(10)
        self.intake_layout.addWidget(_header_container)
        self.intake_layout.addWidget(self.hint_label)
        self.intake_layout.addWidget(self.work_environment_row)
        self.intake_layout.addWidget(self.actions_row)

        self.analysis_tab = QWidget()
        self.analysis_tab.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.analysis_layout = QVBoxLayout(self.analysis_tab)
        self.analysis_layout.setContentsMargins(12, 12, 12, 12)
        self.analysis_layout.setSpacing(10)
        self.analysis_hint_label = QLabel(
            self.format_ui_rtl_text("לאחר טעינת הקבצים ניתן לבצע ניתוח לביקורת ולסקור כאן את הממצאים המרכזיים.")
        )
        self.analysis_hint_label.setWordWrap(True)
        self.analysis_hint_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.analysis_hint_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        self.analysis_layout.addWidget(self.analysis_hint_label)
        self.audit_run_button = QPushButton(self.format_ui_rtl_text("בצע ניתוח לביקורת עבור המשבצת שנבחרה"))
        self.audit_run_button.clicked.connect(self.run_validation)
        self.analysis_layout.addWidget(self.audit_run_button, 0, Qt.AlignmentFlag.AlignRight)

        self.analysis_progress_container = QWidget()
        self.analysis_progress_container.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        analysis_progress_layout = QHBoxLayout(self.analysis_progress_container)
        analysis_progress_layout.setContentsMargins(0, 0, 0, 0)
        analysis_progress_layout.setSpacing(8)
        analysis_progress_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.analysis_progress_label = QLabel(self.format_ui_rtl_text("מעבד קובץ..."))
        self.analysis_progress_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.analysis_progress_label.setStyleSheet("color: #16325c; font-weight: bold;")

        self.analysis_progress_bar = QProgressBar()
        self.analysis_progress_bar.setMinimumWidth(220)
        self.analysis_progress_bar.setMaximumWidth(320)
        self.analysis_progress_bar.setTextVisible(False)
        self.analysis_progress_bar.setRange(0, 0)

        analysis_progress_layout.addWidget(self.analysis_progress_label)
        analysis_progress_layout.addWidget(self.analysis_progress_bar)
        analysis_progress_layout.addStretch(1)
        self.analysis_progress_container.hide()
        self.analysis_layout.addWidget(self.analysis_progress_container, 0, Qt.AlignmentFlag.AlignRight)

        self.audit_export_button = QPushButton(self.format_ui_rtl_text("ייצוא ממצאים לאקסל"))
        self.audit_export_button.clicked.connect(lambda: self.export_audit_findings_to_excel(open_after_export=True))
        self.analysis_layout.addWidget(self.audit_export_button, 0, Qt.AlignmentFlag.AlignRight)

        self.audit_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי ביקורת כללי - ריכוז"))
        self.audit_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        audit_summary_layout = QVBoxLayout(self.audit_summary_group)
        audit_summary_layout.setContentsMargins(8, 14, 8, 8)
        self.audit_summary_table = QTableWidget(0, 10)
        self.audit_summary_table.setItemDelegate(_RightAlignDelegate(self.audit_summary_table))
        self.audit_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("מזהה בקרה"),
            self.format_rtl_text("סוג בדיקה"),
            self.format_rtl_text("קובץ מקור"),
            self.format_rtl_text("תאריך הפקה"),
            self.format_rtl_text("סביבת עבודה"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("תיאור הבדיקה"),
            self.format_rtl_text("רשומות תקינות"),
            self.format_rtl_text("רשומות עם ממצא"),
            self.format_rtl_text("סהכ רשומות"),
        ])
        _audit_summary_hdr = self.audit_summary_table.horizontalHeader()
        _audit_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _audit_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _audit_summary_hdr.setStretchLastSection(False)
        self.audit_summary_table.setColumnWidth(6, 220)  # תיאור הבדיקה
        self.audit_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.audit_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.audit_summary_table.setAlternatingRowColors(True)
        self.audit_summary_table.setMinimumHeight(220)
        self.audit_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.audit_summary_table.itemSelectionChanged.connect(self._refresh_selected_audit_detail)
        audit_summary_layout.addWidget(self.audit_summary_table)
        self.analysis_layout.addWidget(self.audit_summary_group)

        self.audit_detail_group = QGroupBox(self.format_ui_rtl_text("פירוט ממצאי ביקורת"))
        self.audit_detail_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        audit_detail_layout = QVBoxLayout(self.audit_detail_group)
        audit_detail_layout.setContentsMargins(8, 14, 8, 8)
        self.audit_detail_table = QTableWidget(0, 11)
        self.audit_detail_table.setItemDelegate(_RightAlignDelegate(self.audit_detail_table))
        self.audit_detail_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קובץ מקור"),
            self.format_rtl_text("תאריך הפקה"),
            self.format_rtl_text("סביבת עבודה"),
            self.format_rtl_text("קטגוריה"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("תיאור"),
            self.format_rtl_text("סוג בדיקה"),
            self.format_rtl_text("ערך בפועל"),
            self.format_rtl_text("ערך מצופה"),
            self.format_rtl_text("סטטוס"),
            self.format_rtl_text("תיאור מלא"),
        ])
        _audit_detail_hdr = self.audit_detail_table.horizontalHeader()
        _audit_detail_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _audit_detail_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _audit_detail_hdr.setStretchLastSection(True)  # תיאור מלא
        self.audit_detail_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.audit_detail_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.audit_detail_table.setAlternatingRowColors(True)
        self.audit_detail_table.setMinimumHeight(220)
        self.audit_detail_table.setToolTip(self.format_ui_rtl_text("לחיצה כפולה על שורה תציג פירוט מלא של הממצא"))
        self.audit_detail_table.cellDoubleClicked.connect(self.show_audit_detail_dialog)
        audit_detail_layout.addWidget(self.audit_detail_table)
        self.analysis_layout.addWidget(self.audit_detail_group)

        self.review_tab = QWidget()
        self.review_tab.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.review_layout = QVBoxLayout(self.review_tab)
        self.review_layout.setContentsMargins(6, 6, 6, 6)
        self.review_layout.setSpacing(6)

        self.permissions_review_tab = QWidget()
        self.permissions_review_tab.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.permissions_review_layout = QVBoxLayout(self.permissions_review_tab)
        self.permissions_review_layout.setContentsMargins(12, 12, 12, 12)
        self.permissions_review_layout.setSpacing(10)

        self.permissions_hint_label = QLabel(
            self.format_ui_rtl_text("סקירת הרשאות מרכזת ממצאים רוחביים על הרשאות משתמשים במערכת.")
        )
        self.permissions_hint_label.setWordWrap(True)
        self.permissions_hint_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        self.permissions_review_layout.addWidget(self.permissions_hint_label)

        self.permissions_inner_tabs = QTabWidget()
        self.permissions_inner_tabs.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        strong_profiles_page = QWidget()
        strong_profiles_page_layout = QVBoxLayout(strong_profiles_page)
        strong_profiles_page_layout.setContentsMargins(0, 0, 0, 0)
        strong_profiles_page_layout.setSpacing(8)
        self._build_strong_profiles_permissions_section(strong_profiles_page_layout)
        self.permissions_inner_tabs.addTab(strong_profiles_page, self.format_rtl_text("פרופילים למשתמשים חזקים"))

        placeholder_titles: list[str] = []

        user_mgmt_page = QWidget()
        user_mgmt_page_layout = QVBoxLayout(user_mgmt_page)
        user_mgmt_page_layout.setContentsMargins(0, 0, 0, 0)
        user_mgmt_page_layout.setSpacing(8)
        self._build_user_mgmt_permissions_section(user_mgmt_page_layout)
        self.permissions_inner_tabs.addTab(user_mgmt_page, self.format_rtl_text("הרשאות ניהול משתמשים"))

        auth_mgmt_page = QWidget()
        auth_mgmt_page_layout = QVBoxLayout(auth_mgmt_page)
        auth_mgmt_page_layout.setContentsMargins(0, 0, 0, 0)
        auth_mgmt_page_layout.setSpacing(8)
        self._build_auth_mgmt_permissions_section(auth_mgmt_page_layout)
        self.permissions_inner_tabs.addTab(auth_mgmt_page, self.format_rtl_text("הרשאות ניהול הרשאות"))

        rscdok99_page = QWidget()
        rscdok99_page_layout = QVBoxLayout(rscdok99_page)
        rscdok99_page_layout.setContentsMargins(0, 0, 0, 0)
        rscdok99_page_layout.setSpacing(8)
        self._build_rscdok99_permissions_section(rscdok99_page_layout)
        self.permissions_inner_tabs.addTab(rscdok99_page, self.format_rtl_text("הרשאות לתוכנית RSCDOK99"))

        data_mgmt_page = QWidget()
        data_mgmt_page_layout = QVBoxLayout(data_mgmt_page)
        data_mgmt_page_layout.setContentsMargins(0, 0, 0, 0)
        data_mgmt_page_layout.setSpacing(8)
        self._build_data_mgmt_permissions_section(data_mgmt_page_layout)
        self.permissions_inner_tabs.addTab(data_mgmt_page, self.format_rtl_text("הרשאות לניהול נתונים"))

        transport_page = QWidget()
        transport_page_layout = QVBoxLayout(transport_page)
        transport_page_layout.setContentsMargins(0, 0, 0, 0)
        transport_page_layout.setSpacing(8)
        self._build_transport_permissions_section(transport_page_layout)
        self.permissions_inner_tabs.addTab(transport_page, self.format_rtl_text("הרשאה להעברת שינויים"))

        debug_page = QWidget()
        debug_page_layout = QVBoxLayout(debug_page)
        debug_page_layout.setContentsMargins(0, 0, 0, 0)
        debug_page_layout.setSpacing(8)
        self._build_debug_permissions_section(debug_page_layout)
        self.permissions_inner_tabs.addTab(debug_page, self.format_rtl_text("הרשאות לשימוש ב DEBUG"))

        job_mgmt_page = QWidget()
        job_mgmt_page_layout = QVBoxLayout(job_mgmt_page)
        job_mgmt_page_layout.setContentsMargins(0, 0, 0, 0)
        job_mgmt_page_layout.setSpacing(8)
        self._build_job_mgmt_permissions_section(job_mgmt_page_layout)
        self.permissions_inner_tabs.addTab(job_mgmt_page, self.format_rtl_text("הרשאה לעידכון ג'ובים"))

        for placeholder_title in placeholder_titles:
            placeholder_page = QWidget()
            placeholder_layout = QVBoxLayout(placeholder_page)
            placeholder_layout.setContentsMargins(0, 0, 0, 0)
            placeholder_layout.setSpacing(8)
            placeholder_label = QLabel(
                self.format_ui_rtl_text("חוצץ זה ישופעל בשלב הבא של סקירת ההרשאות.")
            )
            placeholder_label.setWordWrap(True)
            placeholder_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
            placeholder_layout.addWidget(placeholder_label)
            placeholder_layout.addStretch(1)
            self.permissions_inner_tabs.addTab(placeholder_page, self.format_rtl_text(placeholder_title))

        self.permissions_review_layout.addWidget(self.permissions_inner_tabs)

        self.settings_tab = QWidget()
        self.settings_tab.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.settings_tab_layout = QVBoxLayout(self.settings_tab)
        self.settings_tab_layout.setContentsMargins(12, 12, 12, 12)
        self.settings_tab_layout.setSpacing(0)

        self.settings_scroll = QScrollArea()
        self.settings_scroll.setWidgetResizable(True)
        self.settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.settings_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.settings_scroll.setMinimumHeight(520)
        self.settings_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.settings_content = QWidget()
        self.settings_content.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.settings_content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.settings_layout = QVBoxLayout(self.settings_content)
        self.settings_layout.setContentsMargins(0, 0, 0, 0)
        self.settings_layout.setSpacing(10)
        self.settings_layout.setAlignment(Qt.AlignmentFlag.AlignRight)

        self.settings_intro_label = QLabel(
            self.format_ui_rtl_text("בטאב זה ניתן לנהל את הגדרות הביקורת והעמודות הנדרשות לכל משבצת.")
        )
        # השתמש ב-QLabel עם rich text HTML RTL ו-align right
        self.settings_intro_label.setText(
            '<div dir="rtl" align="right">בטאב זה ניתן לנהל את הגדרות הביקורת והעמודות הנדרשות לכל משבצת.</div>'
        )
        self.settings_intro_label.setWordWrap(True)
        self.settings_intro_label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.settings_intro_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        self.settings_intro_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.settings_intro_label.setMinimumHeight(40)
        self.settings_intro_label.setMaximumWidth(16777215)
        self.settings_layout.addWidget(self.settings_intro_label)
        self._build_system_settings_sections()

        self.settings_scroll.setWidget(self.settings_content)
        self.settings_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.settings_tab_layout.addWidget(self.settings_scroll)

        self.tabs.addTab(self.intake_tab, self.format_rtl_text("קליטת קבצים"))
        self.tabs.addTab(self.settings_tab, self.format_rtl_text("הגדרות מערכת לביקורת"))
        self.tabs.addTab(self.review_tab, self.format_rtl_text("סקירת דוח משתמשים"))
        self.tabs.addTab(self.permissions_review_tab, self.format_rtl_text("סקירת הרשאות"))
        self.tabs.addTab(self.analysis_tab, self.format_rtl_text("ביצוע ניתוח לביקורת"))
        main_layout.addWidget(self.tabs)

        self.slots_group = QGroupBox(self.format_ui_rtl_text("מקורות קלט לבדיקת SAP HANA APP"))
        self.slots_group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.slots_group.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        slots_group_layout = QVBoxLayout(self.slots_group)
        slots_group_layout.setContentsMargins(8, 18, 8, 8)

        self.slots_scroll = QScrollArea()
        self.slots_scroll.setWidgetResizable(True)
        self.slots_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.slots_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.slots_scroll.setMinimumHeight(280)
        self.slots_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        slots_container = QWidget()
        slots_container.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        slots_layout = QGridLayout(slots_container)
        slots_layout.setContentsMargins(12, 12, 12, 12)
        slots_layout.setHorizontalSpacing(12)
        slots_layout.setVerticalSpacing(10)
        slots_layout.setColumnStretch(0, 0)
        slots_layout.setColumnStretch(1, 1)
        slots_layout.setColumnStretch(2, 2)
        slots_layout.setColumnStretch(3, 0)
        slots_layout.setColumnMinimumWidth(0, 140)
        slots_layout.setColumnMinimumWidth(3, 120)
        slots_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)

        current_row = 0
        for domain in self._ordered_categories():
            palette = self._category_palette(domain)
            in_development = bool(self.DOMAIN_DEFINITIONS.get(domain, {}).get("in_development", False))

            domain_section = QGroupBox(self.format_ui_rtl_text(domain))
            domain_section.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            domain_section.setStyleSheet(
                f"""
                QGroupBox {{
                    font-weight: bold;
                    border: 2px solid {palette['border']};
                    border-radius: 10px;
                    margin-top: 12px;
                    padding-top: 14px;
                    background-color: #ffffff;
                }}
                QGroupBox::title {{
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 4px 12px;
                    background-color: {palette['header']};
                    color: white;
                    border-radius: 6px;
                    font-weight: bold;
                }}
                """
            )
            self.category_sections[domain] = domain_section
            domain_section.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
            domain_layout = QVBoxLayout(domain_section)
            domain_layout.setContentsMargins(12, 18, 12, 12)
            domain_layout.setSpacing(10)

            domain_button = QPushButton("הרץ בדיקות תחום")
            domain_button.setMinimumHeight(34)
            domain_button.setToolTip(self.format_rtl_text(f"הרצת בדיקות עבור תחום {domain}"))
            domain_button.setStyleSheet(
                f"background-color: {palette['button']}; border: 2px solid {palette['border']}; color: white; font-weight: bold;"
            )
            domain_button.clicked.connect(
                lambda _checked=False, d=domain: self.run_domain_validation(d)
            )
            self.category_run_buttons[domain] = domain_button

            sub_cat_style = f"""
                QGroupBox {{
                    font-weight: bold;
                    border: 1px solid {palette['border']};
                    border-radius: 7px;
                    margin-top: 8px;
                    padding-top: 10px;
                    background-color: #f9fafb;
                }}
                QGroupBox::title {{
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 2px 8px;
                    background-color: {palette['header']};
                    color: white;
                    border-radius: 4px;
                    font-weight: bold;
                    font-size: 11px;
                    opacity: 0.85;
                }}
            """

            for sub_cat_info in self.DOMAIN_DEFINITIONS.get(domain, {}).get("sub_categories", []):
                sub_cat_key = str(sub_cat_info["key"])
                sub_cat_desc = str(sub_cat_info["description"])

                sub_group = QGroupBox(self.format_ui_rtl_text(sub_cat_key))
                sub_group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                sub_group.setStyleSheet(sub_cat_style)
                sub_group.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                sub_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

                if in_development:
                    sub_group_layout = QVBoxLayout(sub_group)
                    sub_group_layout.setContentsMargins(10, 14, 10, 10)
                    sub_group_layout.setSpacing(6)

                    desc_label = QLabel(self.format_ui_rtl_text(sub_cat_desc))
                    desc_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                    desc_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
                    desc_label.setWordWrap(True)
                    desc_label.setStyleSheet("color: #4f5d73;")

                    dev_label = QLabel(self.format_ui_rtl_text("⚙ בפיתוח — אין בדיקות אוטומטיות זמינות עדיין"))
                    dev_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                    dev_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    dev_label.setStyleSheet(
                        "color: #6b7280; font-style: italic; background: #f3f4f6;"
                        " border: 1px dashed #d1d5db; border-radius: 4px; padding: 4px 8px;"
                    )

                    sub_group_layout.addWidget(desc_label)
                    sub_group_layout.addWidget(dev_label)
                else:
                    category_layout = QGridLayout(sub_group)
                    category_layout.setContentsMargins(12, 16, 12, 10)
                    category_layout.setHorizontalSpacing(12)
                    category_layout.setVerticalSpacing(10)
                    category_layout.setColumnStretch(0, 0)
                    category_layout.setColumnStretch(1, 1)
                    category_layout.setColumnStretch(2, 2)
                    category_layout.setColumnStretch(3, 0)
                    category_layout.setColumnMinimumWidth(0, 140)
                    category_layout.setColumnMinimumWidth(3, 120)
                    category_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)

                    # Sub-category description row
                    sub_desc_label = QLabel(self.format_ui_rtl_text(sub_cat_desc))
                    sub_desc_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                    sub_desc_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    sub_desc_label.setWordWrap(True)
                    sub_desc_label.setStyleSheet("color: #4f5d73; padding: 2px 0;")
                    category_layout.addWidget(sub_desc_label, 0, 0, 1, 4)

                    section_row = 1
                    for slot_key, metadata in self.SLOT_DEFINITIONS.items():
                        if metadata.get("sub_category") != sub_cat_key:
                            continue

                        display_name = metadata.get("label", slot_key)
                        slot_title = QLabel(self.format_ui_rtl_text(f"{display_name}{' *' if metadata['required'] else ''}"))
                        slot_title.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                        slot_title.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
                        slot_title.setStyleSheet("font-weight: bold;")
                        slot_title.setMinimumWidth(110)

                        description = QLabel(self.format_ui_rtl_text(metadata["description"]))
                        description.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                        description.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
                        description.setWordWrap(True)
                        description.setMinimumHeight(34)
                        description.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

                        sample = QLabel(self.format_ui_rtl_text(f"קובץ צפוי: {metadata['expected_file']}"))
                        sample.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                        sample.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
                        sample.setWordWrap(True)
                        sample.setStyleSheet("color: #5b6573;")
                        sample.setMinimumWidth(120)
                        sample.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

                        status_label = QLabel(self.format_ui_rtl_text("טרם נבחר קובץ"))
                        status_label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                        status_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        status_label.setWordWrap(True)
                        status_label.setMinimumHeight(32)
                        status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                        status_label.setStyleSheet("padding: 6px; background: #ffffff; border: 1px solid #cfd6e4;")

                        extraction_date_label = QLabel(self.format_ui_rtl_text("תאריך הפקה:"))
                        extraction_date_label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                        extraction_date_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        extraction_date_label.setStyleSheet("color: #5b6573;")
                        extraction_date_edit = QLineEdit(self._default_extraction_date())
                        extraction_date_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
                        extraction_date_edit.setPlaceholderText("YYYY-MM-DD")
                        extraction_date_edit.setMinimumHeight(32)
                        extraction_date_edit.setMaximumWidth(170)

                        extraction_date_row = QWidget()
                        extraction_date_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                        extraction_date_layout = QHBoxLayout(extraction_date_row)
                        extraction_date_layout.setContentsMargins(0, 0, 0, 0)
                        extraction_date_layout.setSpacing(6)
                        extraction_date_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        extraction_date_layout.addWidget(extraction_date_label, 0, Qt.AlignmentFlag.AlignRight)
                        extraction_date_layout.addWidget(extraction_date_edit, 0, Qt.AlignmentFlag.AlignRight)
                        extraction_date_layout.addStretch(1)

                        select_button = QPushButton("בחירת קבצים" if slot_key in self.MULTI_FILE_SLOTS else "בחירת קובץ")
                        select_button.setMinimumHeight(34)
                        select_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
                        select_button.clicked.connect(lambda _checked=False, sk=slot_key: self.choose_file(sk))

                        clear_slot_button = QPushButton("נקה")
                        clear_slot_button.setMinimumHeight(34)
                        clear_slot_button.setMinimumWidth(74)
                        clear_slot_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
                        clear_slot_button.clicked.connect(lambda _checked=False, sk=slot_key: self.clear_slot_selection(sk))

                        slot_buttons = QWidget()
                        slot_buttons.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                        slot_buttons_layout = QHBoxLayout(slot_buttons)
                        slot_buttons_layout.setContentsMargins(0, 0, 0, 0)
                        slot_buttons_layout.setSpacing(6)
                        slot_buttons_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        slot_buttons_layout.addWidget(select_button)
                        slot_buttons_layout.addWidget(clear_slot_button)

                        category_layout.setRowMinimumHeight(section_row, 42)
                        category_layout.addWidget(slot_title, section_row, 3)
                        category_layout.addWidget(description, section_row, 2)
                        category_layout.addWidget(sample, section_row, 1)
                        category_layout.addWidget(slot_buttons, section_row, 0)
                        section_row += 1
                        category_layout.setRowMinimumHeight(section_row, 36)
                        category_layout.addWidget(status_label, section_row, 0, 1, 4)
                        section_row += 1
                        category_layout.setRowMinimumHeight(section_row, 34)
                        category_layout.addWidget(extraction_date_row, section_row, 0, 1, 4)
                        section_row += 1

                        self.slot_widgets[slot_key] = {
                            "path_label": status_label,
                            "button": select_button,
                            "clear_button": clear_slot_button,
                            "metadata": metadata,
                            "selected_paths": [],
                            "extraction_date_edit": extraction_date_edit,
                            "extraction_date_label": extraction_date_label,
                        }

                domain_layout.addWidget(sub_group)

            domain_layout.addWidget(domain_button, 0, Qt.AlignmentFlag.AlignRight)
            domain_section.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            slots_layout.addWidget(domain_section, current_row, 0, 1, 4)
            current_row += 1

        bottom_spacer = QLabel("")
        bottom_spacer.setMinimumHeight(120)
        slots_layout.addWidget(bottom_spacer, current_row, 0, 1, 4)
        current_row += 1
        slots_layout.setRowStretch(current_row, 1)
        slots_layout.setRowMinimumHeight(current_row, 20)
        self.slots_scroll.setWidget(slots_container)
        slots_group_layout.addWidget(self.slots_scroll)

        self.slots_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.intake_layout.addWidget(self.slots_group, 1)


        self.user_preview_group = QGroupBox(self.format_ui_rtl_text("רשימת משתמשים שנטענו"))
        self.user_preview_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.user_preview_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.user_preview_group.setMinimumHeight(460)
        user_preview_layout = QVBoxLayout(self.user_preview_group)
        user_preview_layout.setContentsMargins(8, 12, 8, 8)
        user_preview_layout.setSpacing(4)
        user_preview_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # --- Create progress group first ---
        self.user_review_progress_group = QGroupBox(self.format_ui_rtl_text("סיכום התקדמות סקירה"))
        self.user_review_progress_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.user_review_progress_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        user_review_progress_layout = QVBoxLayout(self.user_review_progress_group)
        user_review_progress_layout.setContentsMargins(8, 8, 8, 8)
        user_review_progress_layout.setSpacing(6)

        user_review_counts_row = QWidget()
        user_review_counts_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        user_review_counts_layout = QHBoxLayout(user_review_counts_row)
        user_review_counts_layout.setContentsMargins(0, 0, 0, 0)
        user_review_counts_layout.setSpacing(14)
        user_review_counts_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.user_review_total_label = QLabel(self.format_ui_rtl_text("סה\"כ משתמשים בדוח: 0"))
        self.user_review_total_label.setStyleSheet("font-weight: bold;")
        self.user_review_reviewed_label = QLabel(self.format_ui_rtl_text("משתמשים שנבדקו: 0"))
        self.user_review_reviewed_label.setStyleSheet("font-weight: bold; color: #2e7d32;")
        self.user_review_unreviewed_label = QLabel(self.format_ui_rtl_text("משתמשים שטרם נבדקו: 0"))
        self.user_review_unreviewed_label.setStyleSheet("font-weight: bold; color: #1565c0;")

        user_review_counts_layout.addWidget(self.user_review_total_label, 0, Qt.AlignmentFlag.AlignRight)
        user_review_counts_layout.addWidget(self.user_review_reviewed_label, 0, Qt.AlignmentFlag.AlignRight)
        user_review_counts_layout.addWidget(self.user_review_unreviewed_label, 0, Qt.AlignmentFlag.AlignRight)
        user_review_counts_layout.addStretch(1)
        user_review_progress_layout.addWidget(user_review_counts_row)

        self.user_review_progress_bar = QProgressBar()
        self.user_review_progress_bar.setMinimum(0)
        self.user_review_progress_bar.setMaximum(100)
        self.user_review_progress_bar.setValue(0)
        self.user_review_progress_bar.setTextVisible(True)
        self.user_review_progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.user_review_progress_bar.setFormat("0%")
        self.user_review_progress_bar.setStyleSheet(
            "QProgressBar {"
            "border: 1px solid #b0bec5; border-radius: 4px;"
            "background-color: #f5f7fa; text-align: center;"
            "font-weight: bold; color: #0d47a1;"
            "}"
            "QProgressBar::chunk {"
            "background-color: #42a5f5; border-radius: 3px;"
            "}"
        )
        user_review_progress_layout.addWidget(self.user_review_progress_bar)

        self.user_review_progress_percent_label = QLabel(self.format_ui_rtl_text("התקדמות השלמת סקירה: 0%"))
        self.user_review_progress_percent_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.user_review_progress_percent_label.setStyleSheet("font-weight: bold; color: #0d47a1;")
        user_review_progress_layout.addWidget(self.user_review_progress_percent_label)

        # --- Now create actions row ---
        self.user_preview_actions_row = QWidget()
        self.user_preview_actions_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.user_preview_actions_row.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.user_preview_actions_row.setMaximumHeight(40)
        user_preview_actions_layout = QHBoxLayout(self.user_preview_actions_row)
        user_preview_actions_layout.setContentsMargins(0, 0, 0, 0)
        user_preview_actions_layout.setSpacing(8)
        user_preview_actions_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.user_preview_export_button = QPushButton("ייצוא סקירה לאקסל")
        self.user_preview_export_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.user_preview_export_button.clicked.connect(lambda: self.export_user_preview_to_excel(open_after_export=True))
        user_preview_actions_layout.addWidget(self.user_preview_export_button, 0, Qt.AlignmentFlag.AlignRight)

        self.user_preview_import_button = QPushButton("ייבוא סקירה מאקסל")
        self.user_preview_import_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.user_preview_import_button.clicked.connect(self.import_user_review_from_excel)
        user_preview_actions_layout.addWidget(self.user_preview_import_button, 0, Qt.AlignmentFlag.AlignRight)

        self.user_preview_send_business_button = QPushButton("שליחת הדוח לגורם עסקי")
        self.user_preview_send_business_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.user_preview_send_business_button.clicked.connect(self.draft_user_review_email_to_business)
        user_preview_actions_layout.addWidget(self.user_preview_send_business_button, 0, Qt.AlignmentFlag.AlignRight)

        self.user_preview_send_technical_button = QPushButton("שליחת הדוח לגורם טכני")
        self.user_preview_send_technical_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.user_preview_send_technical_button.clicked.connect(self.draft_user_review_email_to_technical)
        user_preview_actions_layout.addWidget(self.user_preview_send_technical_button, 0, Qt.AlignmentFlag.AlignRight)

        self.user_preview_columns_button = QPushButton("הוסף / מחק עמודות")
        self.user_preview_columns_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.user_preview_columns_button.clicked.connect(self.show_user_preview_column_dialog)
        user_preview_actions_layout.addWidget(self.user_preview_columns_button, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_actions_layout.addStretch(1)

        # Add the progress group first, then the actions row
        user_preview_layout.addWidget(self.user_review_progress_group, 0, Qt.AlignmentFlag.AlignTop)
        user_preview_layout.addWidget(self.user_preview_actions_row, 0, Qt.AlignmentFlag.AlignTop)

        self.user_preview_filter_row = QWidget()
        self.user_preview_filter_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.user_preview_filter_row.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        user_preview_filter_layout = QHBoxLayout(self.user_preview_filter_row)
        user_preview_filter_layout.setContentsMargins(0, 0, 0, 0)
        user_preview_filter_layout.setSpacing(8)
        user_preview_filter_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.user_preview_filter_label = QLabel(self.format_ui_rtl_text("סינון משתמשים:"))
        self.user_preview_status_filter = QComboBox()
        self.user_preview_status_filter.setMinimumWidth(220)
        self.user_preview_status_filter.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        for filter_value, filter_label in self.USER_PREVIEW_FILTER_OPTIONS:
            self.user_preview_status_filter.addItem(self.format_rtl_text(filter_label), filter_value)

        self.audit_period_from_label = QLabel(self.format_ui_rtl_text("מתאריך:"))
        self.audit_period_from_edit = QLineEdit("")
        self.audit_period_from_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.audit_period_from_edit.setPlaceholderText("YYYY-MM-DD")
        self.audit_period_from_edit.setMaximumWidth(130)

        self.audit_period_to_label = QLabel(self.format_ui_rtl_text("עד תאריך:"))
        self.audit_period_to_edit = QLineEdit("")
        self.audit_period_to_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.audit_period_to_edit.setPlaceholderText("YYYY-MM-DD")
        self.audit_period_to_edit.setMaximumWidth(130)

        user_preview_filter_layout.addWidget(self.user_preview_filter_label, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addWidget(self.user_preview_status_filter, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_from_label, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_from_edit, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_to_label, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_to_edit, 0, Qt.AlignmentFlag.AlignRight)
        user_preview_filter_layout.addStretch(1)
        user_preview_layout.addWidget(self.user_preview_filter_row, 0, Qt.AlignmentFlag.AlignTop)

        self.user_preview_hint = QLabel(
            '<p align="right" style="margin:0">הטבלה מציגה את משתמשי USR02 עם העשרת נתונים מקובצי USER_ADDR ו-ADR6 בלבד.</p>'
        )
        self.user_preview_hint.setWordWrap(True)
        self.user_preview_hint.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.user_preview_hint.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        self.user_preview_hint.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.user_preview_hint.setMaximumHeight(44)
        user_preview_layout.addWidget(self.user_preview_hint, 0, Qt.AlignmentFlag.AlignTop)

        self.user_preview_table = QTableWidget(0, 0)
        self.user_preview_table.setItemDelegate(_RightAlignDelegate(self.user_preview_table))
        self.user_preview_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self.user_preview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.user_preview_table.setAlternatingRowColors(True)
        self.user_preview_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.user_preview_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.user_preview_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.user_preview_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.user_preview_table.setMinimumHeight(420)
        self.user_preview_table.setMaximumHeight(16777215)
        self._configure_user_preview_table()
        self.user_preview_table.itemChanged.connect(self._handle_user_preview_item_changed)
        self.user_preview_status_filter.currentIndexChanged.connect(self.refresh_user_preview)
        self.audit_period_from_edit.editingFinished.connect(self.refresh_user_preview)
        self.audit_period_to_edit.editingFinished.connect(self.refresh_user_preview)
        user_preview_layout.addWidget(self.user_preview_table, 1)
        self.review_layout.addWidget(self.user_preview_group, 1)

        self.run_log_group = QGroupBox(self.format_ui_rtl_text("לוג קבצים שנקלטו"))
        self.run_log_group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        run_log_layout = QVBoxLayout(self.run_log_group)
        run_log_layout.setContentsMargins(12, 18, 12, 12)
        self.run_log_table = QTableWidget(0, 10)
        self.run_log_table.setItemDelegate(_RightAlignDelegate(self.run_log_table))
        self.run_log_table.setHorizontalHeaderLabels(["משבצת", "קבוצת דוחות", "קובץ", "תאריך הפקה", "רשומות שנקלטו", "סטטוס", "מספר שגיאות", "תיאור שגיאה", "תאריך בדיקה", "שעת בדיקה"])
        self.run_log_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        self.run_log_table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        _run_log_hdr = self.run_log_table.horizontalHeader()
        _run_log_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _run_log_hdr.setStretchLastSection(False)
        self.run_log_table.setColumnWidth(1, 150)
        self.run_log_table.setColumnWidth(2, 180)
        self.run_log_table.setColumnWidth(7, 220)  # תיאור שגיאה
        self.run_log_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.run_log_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.run_log_table.setAlternatingRowColors(True)
        self.run_log_table.setWordWrap(True)
        self.run_log_table.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.run_log_table.setMinimumHeight(160)
        self.run_log_table.setToolTip("לחיצה כפולה על שורה תפתח פירוט מלא עבור הקובץ")
        self.run_log_table.cellDoubleClicked.connect(self.show_log_details)
        run_log_layout.addWidget(self.run_log_table)
        self.run_log_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.run_log_group.setMaximumHeight(320)
        self.intake_layout.addWidget(self.run_log_group, 0)

        self.required_columns_group = QGroupBox(self.format_ui_rtl_text("עמודות חובה לבדיקה"))
        self.required_columns_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        required_layout = QHBoxLayout(self.required_columns_group)
        self.required_columns_edit = QLineEdit("")
        self.required_columns_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.required_columns_edit.setPlaceholderText("יוזן אוטומטית לפי המשבצת שנבחרה")
        required_layout.addWidget(self.required_columns_edit)
        self.required_columns_group.hide()

        self.summary_group = QGroupBox(self.format_ui_rtl_text("סיכום בדיקה"))
        self.summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        summary_layout = QGridLayout(self.summary_group)
        summary_layout.setContentsMargins(12, 18, 12, 12)
        summary_layout.setHorizontalSpacing(10)
        summary_layout.setVerticalSpacing(8)
        summary_items = [
            ("שורות שנבדקו", "total", "0"),
            ("שורות תקינות", "valid", "0"),
            ("שורות שגויות", "invalid", "0"),
            ("סטטוס", "status", "ממתין להרצה"),
        ]
        for column, (title, key, default_value) in enumerate(summary_items):
            title_label = QLabel(title)
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            title_label.setStyleSheet("font-weight: bold; qproperty-alignment: 'AlignCenter|AlignVCenter';")
            value_label = QLabel(default_value)
            value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            value_label.setStyleSheet("font-size: 18px; padding: 6px; qproperty-alignment: 'AlignCenter|AlignVCenter';")
            summary_layout.addWidget(title_label, 0, column)
            summary_layout.addWidget(value_label, 1, column)
            self.summary_labels[key] = value_label
        self.summary_group.hide()
        self.intake_layout.addWidget(self.summary_group)

        self.results_group = QGroupBox(self.format_ui_rtl_text("שגיאות קליטה"))
        self.results_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        results_layout = QVBoxLayout(self.results_group)
        results_layout.setContentsMargins(12, 18, 12, 12)
        self.issues_table = QTableWidget(0, 3)
        self.issues_table.setItemDelegate(_RightAlignDelegate(self.issues_table))
        self.issues_table.setHorizontalHeaderLabels(["מספר שורה", "שם עמודה", "הודעת שגיאה"])
        self.issues_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        self.issues_table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        _issues_hdr = self.issues_table.horizontalHeader()
        _issues_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _issues_hdr.setStretchLastSection(True)  # הודעת שגיאה
        self.issues_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.issues_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.issues_table.setAlternatingRowColors(True)
        results_layout.addWidget(self.issues_table)
        self.issues_table.setMinimumHeight(180)
        self.results_group.hide()
        self.intake_layout.addWidget(self.results_group)

        central_widget.setStyleSheet(
            """
            QWidget {
                background-color: #f5f7fb;
                font-family: 'Segoe UI';
                font-size: 9.5pt;
            }
            QLabel {
                qproperty-alignment: 'AlignRight|AlignVCenter';
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #c7cfda;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: #f9fbfe;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                background-color: #f5f7fb;
            }
            QPushButton {
                background-color: #e9eef7;
                border: 1px solid #b7c4d8;
                border-radius: 6px;
                padding: 4px 10px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #dbe7f8;
            }
            QLineEdit {
                background-color: white;
                border: 1px solid #cfd6e4;
                padding: 4px;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #cfd6e4;
                gridline-color: #d7deea;
            }
            QTableWidget::item {
                padding-right: 4px;
            }
            """
        )

    def _ordered_categories(self) -> list[str]:
        return list(self.DOMAIN_DEFINITIONS.keys())

    def _ordered_sub_categories(self, domain: str) -> list[str]:
        return [
            sub["key"]
            for sub in self.DOMAIN_DEFINITIONS.get(domain, {}).get("sub_categories", [])
        ]

    def _category_palette(self, category: str) -> dict[str, str]:
        if category.startswith("MA"):
            return {"header": "#16325c", "button": "#16325c", "border": "#16325c"}
        if category.startswith("MC"):
            return {"header": "#1b5e20", "button": "#1b5e20", "border": "#1b5e20"}
        # MO — in development, shown in gray
        return {"header": "#6b7280", "button": "#6b7280", "border": "#9ca3af"}

    @staticmethod
    def _default_extraction_date() -> str:
        return datetime.now().strftime("%Y-%m-%d")


    def _build_system_settings_sections(self) -> None:
        # --- Outer wrapper for flush right alignment ---
        buttons_row = QWidget()
        buttons_row.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        buttons_row.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        buttons_layout = QHBoxLayout(buttons_row)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        buttons_layout.setAlignment(Qt.AlignmentFlag.AlignRight)

        self.settings_save_btn = QPushButton(self.format_ui_rtl_text("שמור הגדרות"))
        self.settings_save_btn.clicked.connect(self._save_system_settings)
        self.settings_save_btn.setStyleSheet("text-align: left; padding-right: 18px;")

        self.settings_reset_btn = QPushButton(self.format_ui_rtl_text("טען ברירות מחדל"))
        self.settings_reset_btn.clicked.connect(self._reset_system_settings_form)
        self.settings_reset_btn.setStyleSheet("text-align: left; padding-right: 18px;")

        buttons_layout.addWidget(self.settings_save_btn)
        buttons_layout.addWidget(self.settings_reset_btn)
        # Ensure the row itself is aligned right in the parent layout
        self.settings_layout.addWidget(buttons_row, alignment=Qt.AlignmentFlag.AlignRight)

        review_group, review_layout, review_unavailable_label = self._build_settings_group(
            "טווח תקופת הביקורת",
        )

        start_widget = QDateEdit()
        start_widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        start_widget.setDisplayFormat("yyyy-MM-dd")
        start_widget.setCalendarPopup(True)
        end_widget = QDateEdit()
        end_widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        end_widget.setDisplayFormat("yyyy-MM-dd")
        end_widget.setCalendarPopup(True)

        self.system_settings_widgets["user_review_period.start_date"] = start_widget
        self.system_settings_widgets["user_review_period.end_date"] = end_widget

        date_row = QWidget()
        date_row.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        date_row_layout = QHBoxLayout(date_row)
        date_row_layout.setContentsMargins(0, 0, 0, 0)
        date_row_layout.setSpacing(8)
        # LTR layout: stretch pushes content to the right.
        # Items added last are rightmost. For RTL label convention (label to the right of field):
        # Visual (RTL read): מתאריך: [start] | עד תאריך: [end]
        date_row_layout.addStretch(1)
        date_row_layout.addWidget(end_widget)
        date_row_layout.addWidget(QLabel(self.format_ui_rtl_text("עד תאריך:")))
        date_row_layout.addWidget(start_widget)
        date_row_layout.addWidget(QLabel(self.format_ui_rtl_text("מתאריך:")))
        review_layout.addWidget(date_row)

        self.settings_layout.addWidget(review_group)
        self.system_settings_sections["user_review_period"] = review_group
        self.system_settings_unavailable_labels["user_review_period"] = review_unavailable_label

        authorized_stms_group, authorized_stms_table, authorized_stms_unavailable_label = self._build_super_users_section()
        self.settings_layout.addWidget(authorized_stms_group)
        self.system_settings_widgets["authorized_stms_users"] = authorized_stms_table
        self.system_settings_widgets["super_users"] = authorized_stms_table
        self.system_settings_sections["authorized_stms_users"] = authorized_stms_group
        self.system_settings_sections["super_users"] = authorized_stms_group
        self.system_settings_unavailable_labels["authorized_stms_users"] = authorized_stms_unavailable_label
        self.system_settings_unavailable_labels["super_users"] = authorized_stms_unavailable_label

        generic_users_group = self._add_settings_text_list_section(
            "generic_users",
            "משתמשים גנריים",
            "רשימה מופרדת שורות",
            read_only=True,
        )
        self.system_settings_sections["generic_users"] = generic_users_group

        self._add_settings_text_list_section(
            "critical_roles",
            "פרופילים משתמשיי על",
            "רשימה מופרדת שורות",
            read_only=True,
        )

        password_group, password_layout, password_unavailable_label = self._build_settings_group(
            "ברירות מחדל למדיניות סיסמה",
            "ערכים לוגיים לבקרות סיסמה ב-APP (כאשר נתוני המקור זמינים).",
        )
        password_grid = QGridLayout()
        password_fields = [
            ("minimal_password_length", "אורך סיסמה מינימלי"),
            ("maximum_invalid_connect_attempts", "login/fails_to_session_end - מקסימום ניסיונות כושלים"),
            ("max_password_age_days", "תוקף סיסמה (ימים)"),
            ("password_max_idle_initial", "מספר הימים לתוקף סיסמה ראשונית"),
            ("password_change_for_SSO", "חובת שינוי סיסמה ראשונית SSO"),
            ("login/fails_to_user_lock", "מקסימום ניסיונות כושלים נעילת משתמש"),
            ("password_history_size", "היסטוריית סיסמאות"),
            ("min_password_digits", "ספרות מינימליות"),
            ("min_password_letters", "אותיות מינימליות"),
            ("min_password_lowercase", "אותיות קטנות מינימליות"),
            ("min_password_uppercase", "אותיות גדולות מינימליות"),
            ("min_password_specials", "תווים מיוחדים מינימליים"),
            ("failed_user_auto_unlock", "שיחרור אוטומטי של משתמש נעול"),
            ("rdisp/gui_auto_logout", "התנתקות אוטומטית GUI")
        ]
        for index, (field_name, label_text) in enumerate(password_fields):
            row = index // 2
            col = (index % 2) * 2
            label = QLabel(self.format_ui_rtl_text(label_text))
            widget: object
            if field_name in {"password_change_for_SSO", "failed_user_auto_unlock"}:
                widget = QComboBox()
                widget.addItems(["TRUE", "FALSE"])
            else:
                widget = QLineEdit()
                if isinstance(widget, QLineEdit):
                    widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
                    widget.setMaxLength(6)
            if hasattr(widget, "setMaximumWidth"):
                widget.setMaximumWidth(120)
            self.system_settings_widgets[f"password_policy_defaults.{field_name}"] = widget
            password_grid.addWidget(label, row, col)
            password_grid.addWidget(widget, row, col + 1)
        password_layout.addLayout(password_grid)
        self.settings_layout.addWidget(password_group)
        self.system_settings_sections["password_policy_defaults"] = password_group
        self.system_settings_unavailable_labels["password_policy_defaults"] = password_unavailable_label

        threshold_group, threshold_layout, threshold_unavailable_label = self._build_settings_group(
            "הגדרות נוספות",
            "סף חוסר פעילות משמש לבניית ממצאים אוטומטיים בסקירת משתמשים.",
        )
        threshold_form = QFormLayout()
        threshold_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        threshold_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        threshold_widget = QLineEdit()
        self.system_settings_widgets["inactive_days_threshold"] = threshold_widget
        threshold_form.addRow("סף חוסר פעילות (ימים)", threshold_widget)
        threshold_layout.addLayout(threshold_form)
        self.settings_layout.addWidget(threshold_group)
        self.system_settings_sections["inactive_days_threshold"] = threshold_group
        self.system_settings_unavailable_labels["inactive_days_threshold"] = threshold_unavailable_label

        email_group, email_layout, email_unavailable_label = self._build_settings_group(
            "הגדרות תפוצת מייל",
            "הגדר כתובות מייל של גורם עסקי וגורם טכני עבור יצירת טיוטות שליחת דוח סקירת משתמשים.",
        )
        email_form = QFormLayout()
        email_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        email_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        business_email_widget = QLineEdit()
        business_email_widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        business_email_widget.setPlaceholderText("business.reviewer@company.com")
        technical_email_widget = QLineEdit()
        technical_email_widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        technical_email_widget.setPlaceholderText("technical.reviewer@company.com")

        self.system_settings_widgets["business_reviewer_email"] = business_email_widget
        self.system_settings_widgets["technical_reviewer_email"] = technical_email_widget

        email_form.addRow("כתובת מייל גורם עסקי", business_email_widget)
        email_form.addRow("כתובת מייל גורם טכני", technical_email_widget)
        email_layout.addLayout(email_form)
        self.settings_layout.addWidget(email_group)
        self.system_settings_sections["business_reviewer_email"] = email_group
        self.system_settings_sections["technical_reviewer_email"] = email_group
        self.system_settings_unavailable_labels["business_reviewer_email"] = email_unavailable_label
        self.system_settings_unavailable_labels["technical_reviewer_email"] = email_unavailable_label

    def _build_settings_group(self, title: str, description: str | None = None) -> tuple[QGroupBox, QVBoxLayout, QLabel]:
        group = QGroupBox(self.format_ui_rtl_text(title))
        group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 14, 10, 10)
        layout.setSpacing(8)
        if description:
            # Use HTML for robust RTL right alignment
            html = f'<div dir="rtl" align="right">{self.format_ui_rtl_text(description)}</div>'
            description_label = QLabel(html)
            description_label.setWordWrap(True)
            description_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
            layout.addWidget(description_label)

        unavailable_label = QLabel(self.format_ui_rtl_text("הגדרה זו לא זמינה ללא קובץ מקור רלוונטי"))
        unavailable_label.setWordWrap(True)
        unavailable_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        unavailable_label.setStyleSheet("color: gray; font-style: italic;")
        unavailable_label.setVisible(False)
        layout.addWidget(unavailable_label)
        return group, layout, unavailable_label

    def _build_super_users_section(self) -> tuple[QGroupBox, QTableWidget, QLabel]:
        group, layout, unavailable_label = self._build_settings_group(
            "מורשי STMS / ניהול שינויים",
            "רשימה לבנה של משתמשים מורשים להעברת טרנספורטים לייצור. יש להזין CLIENT ו-BNAME.",
        )
        table = QTableWidget(0, 2)
        table.setItemDelegate(_RightAlignDelegate(table))
        table.setHorizontalHeaderLabels(["CLIENT", "BNAME מורשה STMS"])
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setMinimumHeight(140)
        layout.addWidget(table)

        control_row = QWidget()
        control_layout = QHBoxLayout(control_row)
        control_layout.setContentsMargins(0, 0, 0, 0)
        control_layout.setSpacing(8)
        control_layout.addStretch(1)

        add_button = QPushButton(self.format_ui_rtl_text("הוסף שורה"))
        remove_button = QPushButton(self.format_ui_rtl_text("הסר שורה"))
        add_button.clicked.connect(lambda: self._append_super_user_row(table))
        remove_button.clicked.connect(lambda: self._remove_selected_super_user_row(table))
        control_layout.addWidget(remove_button)
        control_layout.addWidget(add_button)
        layout.addWidget(control_row)

        return group, table, unavailable_label

    def _append_super_user_row(self, table: QTableWidget) -> None:
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(""))
        table.setItem(row, 1, QTableWidgetItem(""))

    def _remove_selected_super_user_row(self, table: QTableWidget) -> None:
        selected_rows = sorted((index.row() for index in table.selectionModel().selectedRows()), reverse=True)
        for row_index in selected_rows:
            table.removeRow(row_index)

    def _add_settings_text_list_section(
        self,
        key: str,
        title: str,
        description: str,
        read_only: bool = False,
    ) -> QGroupBox:
        group, group_layout, unavailable_label = self._build_settings_group(title, description)
        editor = QPlainTextEdit()
        editor.setMinimumHeight(90)
        if read_only:
            editor.setReadOnly(True)
            editor.setToolTip(self.format_ui_rtl_text("שדה זה מוצג לקריאה בלבד ואינו ניתן לשינוי"))
        self.system_settings_widgets[key] = editor
        group_layout.addWidget(editor)
        self.settings_layout.addWidget(group)
        self.system_settings_sections[key] = group
        self.system_settings_unavailable_labels[key] = unavailable_label
        return group

    @staticmethod
    def _safe_int(value: object, fallback: int) -> int:
        try:
            return int(str(value).strip())
        except Exception:
            return fallback

    def _system_settings_path(self) -> Path:
        return self.config.output_dir / "system_settings.json"

    def _default_system_settings(self) -> dict[str, Any]:
        return {
            "work_environment": "FPP",
            "business_reviewer_email": "",
            "technical_reviewer_email": "",
            "generic_users": ["SAP", "DDIC", "TMSADM", "SAPCPIC"],
            "authorized_stms_users": [],
            "super_users": [],
            "critical_roles": ["SAP_ALL", "SAP_NEW"],
            "critical_privileges": ["S_TABU_DIS", "S_USER_GRP", "S_USER_AGR"],
            "password_policy_defaults": {
                "minimal_password_length": 8,
                "maximum_invalid_connect_attempts": 6,
                "password_expire_warning_time": 14,
                "max_password_age_days": 90,
                "initial_password_change_max_days": 2,
                "password_change_for_SSO": "TRUE",
                "password_max_idle_initial": 0,
                "login/fails_to_user_lock": 0,
                "password_history_size": 5,
                "min_password_digits": 0,
                "min_password_letters": 0,
                "min_password_lowercase": 0,
                "min_password_uppercase": 0,
                "min_password_specials": 0,
                "failed_user_auto_unlock": "FALSE",
                "fails_to_session_end": 0,
                "rdisp/gui_auto_logout": 0,
            },
            "inactive_days_threshold": 90,
            "user_review_period": {
                "start_date": datetime.now().replace(month=1, day=1).strftime("%Y-%m-%d"),
                "end_date": datetime.now().replace(month=12, day=31).strftime("%Y-%m-%d"),
            },
            "file_mappings": {
                slot_key: str(metadata.get("expected_file", ""))
                for slot_key, metadata in self.SLOT_DEFINITIONS.items()
            },
        }

    def _current_system_settings(self) -> dict[str, Any]:
        defaults = self._default_system_settings()
        settings_path = self._system_settings_path()
        if not settings_path.exists():
            return defaults
        try:
            loaded = json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception:
            return defaults
        if not isinstance(loaded, dict):
            return defaults
        if "generic_users" not in loaded and "critical_users" in loaded:
            loaded["generic_users"] = loaded.get("critical_users", [])
        if "authorized_stms_users" not in loaded and "super_users" in loaded:
            loaded["authorized_stms_users"] = loaded.get("super_users", [])
        if "super_users" not in loaded and "authorized_stms_users" in loaded:
            loaded["super_users"] = loaded.get("authorized_stms_users", [])
        merged = copy.deepcopy(defaults)
        for key, value in loaded.items():
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key].update(value)
            else:
                merged[key] = value
        return merged

    def _sync_review_filters_from_settings(self, settings: dict[str, Any]) -> None:
        period_cfg = settings.get("user_review_period", {}) if isinstance(settings, dict) else {}
        start_text = str(period_cfg.get("start_date", "")).strip()
        end_text = str(period_cfg.get("end_date", "")).strip()
        if hasattr(self, "audit_period_from_edit") and start_text:
            self.audit_period_from_edit.setText(start_text)
        if hasattr(self, "audit_period_to_edit") and end_text:
            self.audit_period_to_edit.setText(end_text)

    def _load_system_settings_into_form(self, settings: dict[str, Any], load_review_period: bool = True) -> None:
        settings = settings or self._default_system_settings()

        selected_environment = self._normalize_work_environment_code(settings.get("work_environment", "FPP"))
        if hasattr(self, "work_environment_combo") and isinstance(self.work_environment_combo, QComboBox):
            for index in range(self.work_environment_combo.count()):
                if str(self.work_environment_combo.itemData(index) or "").strip().upper() == selected_environment:
                    self.work_environment_combo.setCurrentIndex(index)
                    break

        def _fill_lines(key: str) -> None:
            editor = self.system_settings_widgets.get(key)
            values = settings.get(key, [])
            if isinstance(editor, QPlainTextEdit):
                editor.setPlainText("\n".join(str(item).strip() for item in values if str(item).strip()))

        _fill_lines("generic_users")
        _fill_lines("critical_roles")
        _fill_lines("critical_privileges")

        authorized_table = self.system_settings_widgets.get("authorized_stms_users")
        authorized_users = settings.get("authorized_stms_users", []) if isinstance(settings, dict) else []
        if not authorized_users and isinstance(settings, dict):
            authorized_users = settings.get("super_users", [])
        if isinstance(authorized_table, QTableWidget):
            authorized_table.setRowCount(0)
            if isinstance(authorized_users, list):
                for authorized_user in authorized_users:
                    if isinstance(authorized_user, dict):
                        mandt = str(authorized_user.get("MANDT", "")).strip()
                        bname = str(authorized_user.get("BNAME", "")).strip()
                        if mandt or bname:
                            row = authorized_table.rowCount()
                            authorized_table.insertRow(row)
                            authorized_table.setItem(row, 0, QTableWidgetItem(mandt))
                            authorized_table.setItem(row, 1, QTableWidgetItem(bname))

        if load_review_period:
            period_cfg = settings.get("user_review_period", {}) if isinstance(settings, dict) else {}
            start_text = str(period_cfg.get("start_date", self._default_extraction_date())).strip()
            end_text = str(period_cfg.get("end_date", self._default_extraction_date())).strip()

            start_widget = self.system_settings_widgets.get("user_review_period.start_date")
            end_widget = self.system_settings_widgets.get("user_review_period.end_date")
            if isinstance(start_widget, QDateEdit):
                date_value = QDate.fromString(start_text, "yyyy-MM-dd")
                start_widget.setDate(date_value if date_value.isValid() else QDate.currentDate())
            if isinstance(end_widget, QDateEdit):
                date_value = QDate.fromString(end_text, "yyyy-MM-dd")
                end_widget.setDate(date_value if date_value.isValid() else QDate.currentDate())

        for mapping_key in self.system_settings_file_mapping_order:
            widget = self.system_settings_widgets.get(f"file_mappings.{mapping_key}")
            file_mappings = settings.get("file_mappings", {})
            if isinstance(widget, QLineEdit) and isinstance(file_mappings, dict):
                widget.setText(str(file_mappings.get(mapping_key, "")))

        threshold_widget = self.system_settings_widgets.get("inactive_days_threshold")
        if isinstance(threshold_widget, QLineEdit):
            threshold_widget.setText(str(settings.get("inactive_days_threshold", 90)))

        business_email_widget = self.system_settings_widgets.get("business_reviewer_email")
        if isinstance(business_email_widget, QLineEdit):
            business_email_widget.setText(str(settings.get("business_reviewer_email", "")).strip())
        technical_email_widget = self.system_settings_widgets.get("technical_reviewer_email")
        if isinstance(technical_email_widget, QLineEdit):
            technical_email_widget.setText(str(settings.get("technical_reviewer_email", "")).strip())

        password_defaults = settings.get("password_policy_defaults", {}) if isinstance(settings, dict) else {}
        if isinstance(password_defaults, dict):
            for key, value in password_defaults.items():
                widget = self.system_settings_widgets.get(f"password_policy_defaults.{key}")
                if isinstance(widget, QComboBox):
                    widget.setCurrentText(str(value))
                elif isinstance(widget, QLineEdit):
                    widget.setText(str(value))

    def _collect_system_settings_from_form(self) -> dict[str, Any]:
        def _lines_from_editor(editor: object) -> list[str]:
            if not isinstance(editor, QPlainTextEdit):
                return []
            return [line.strip() for line in editor.toPlainText().splitlines() if line.strip()]

        settings = copy.deepcopy(self._current_system_settings())
        default_settings = self._default_system_settings()
        settings["work_environment"] = self._current_work_environment_code()

        generic_users_editor = self.system_settings_widgets.get("generic_users")
        if isinstance(generic_users_editor, QPlainTextEdit):
            settings["generic_users"] = _lines_from_editor(generic_users_editor)

        critical_roles_editor = self.system_settings_widgets.get("critical_roles")
        if isinstance(critical_roles_editor, QPlainTextEdit):
            settings["critical_roles"] = _lines_from_editor(critical_roles_editor)

        critical_privileges_editor = self.system_settings_widgets.get("critical_privileges")
        if isinstance(critical_privileges_editor, QPlainTextEdit):
            settings["critical_privileges"] = _lines_from_editor(critical_privileges_editor)

        authorized_table = self.system_settings_widgets.get("authorized_stms_users")
        authorized_users: list[dict[str, str]] = []
        if isinstance(authorized_table, QTableWidget):
            for row_index in range(authorized_table.rowCount()):
                mandt_item = authorized_table.item(row_index, 0)
                bname_item = authorized_table.item(row_index, 1)
                mandt_text = str(mandt_item.text()).strip() if isinstance(mandt_item, QTableWidgetItem) else ""
                bname_text = str(bname_item.text()).strip() if isinstance(bname_item, QTableWidgetItem) else ""
                if mandt_text or bname_text:
                    authorized_users.append({"MANDT": mandt_text, "BNAME": bname_text})
        settings["authorized_stms_users"] = authorized_users
        settings["super_users"] = authorized_users

        period_start_widget = self.system_settings_widgets.get("user_review_period.start_date")
        period_end_widget = self.system_settings_widgets.get("user_review_period.end_date")
        if isinstance(period_start_widget, QDateEdit) and isinstance(period_end_widget, QDateEdit):
            settings["user_review_period"] = {
                "start_date": period_start_widget.date().toString("yyyy-MM-dd"),
                "end_date": period_end_widget.date().toString("yyyy-MM-dd"),
            }

        file_mappings = {}
        has_mapping_widgets = False
        for mapping_key in self.system_settings_file_mapping_order:
            widget = self.system_settings_widgets.get(f"file_mappings.{mapping_key}")
            if isinstance(widget, QLineEdit):
                has_mapping_widgets = True
                file_mappings[mapping_key] = widget.text().strip()
        if has_mapping_widgets:
            settings["file_mappings"] = file_mappings

        threshold_widget = self.system_settings_widgets.get("inactive_days_threshold")
        if isinstance(threshold_widget, QLineEdit):
            settings["inactive_days_threshold"] = self._safe_int(threshold_widget.text(), 90)

        business_email_widget = self.system_settings_widgets.get("business_reviewer_email")
        if isinstance(business_email_widget, QLineEdit):
            settings["business_reviewer_email"] = business_email_widget.text().strip()
        technical_email_widget = self.system_settings_widgets.get("technical_reviewer_email")
        if isinstance(technical_email_widget, QLineEdit):
            settings["technical_reviewer_email"] = technical_email_widget.text().strip()

        password_defaults = {}
        for field_name in [
            "minimal_password_length",
            "maximum_invalid_connect_attempts",
            "max_password_age_days",
            "password_max_idle_initial",
            "password_change_for_SSO",
            "login/fails_to_user_lock",
            "password_history_size",
            "min_password_digits",
            "min_password_letters",
            "min_password_lowercase",
            "min_password_uppercase",
            "min_password_specials",
            "failed_user_auto_unlock",
            "rdisp/gui_auto_logout",
        ]:
            widget = self.system_settings_widgets.get(f"password_policy_defaults.{field_name}")
            if isinstance(widget, QComboBox):
                password_defaults[field_name] = widget.currentText().strip()
            elif isinstance(widget, QLineEdit):
                default_value = default_settings["password_policy_defaults"].get(field_name, 0)
                password_defaults[field_name] = self._safe_int(widget.text(), int(default_value))
        settings["password_policy_defaults"] = password_defaults

        return settings

    def _normalize_work_environment_code(self, value: object) -> str:
        normalized = str(value or "").strip().upper()
        valid_codes = {code for code, _label in self.WORK_ENVIRONMENT_OPTIONS}
        return normalized if normalized in valid_codes else "FPP"

    def _current_work_environment_code(self) -> str:
        if hasattr(self, "work_environment_combo") and isinstance(self.work_environment_combo, QComboBox):
            selected_data = self.work_environment_combo.currentData()
            return self._normalize_work_environment_code(selected_data)
        settings = self._current_system_settings()
        return self._normalize_work_environment_code(settings.get("work_environment", "FPP"))

    def _current_work_environment_label(self) -> str:
        selected_code = self._current_work_environment_code()
        for env_code, env_label in self.WORK_ENVIRONMENT_OPTIONS:
            if env_code == selected_code:
                return env_label
        return selected_code

    def _persist_work_environment_selection(self, _index: int) -> None:
        try:
            settings = self._current_system_settings()
            settings["work_environment"] = self._current_work_environment_code()
            settings_path = self._system_settings_path()
            settings_path.parent.mkdir(parents=True, exist_ok=True)
            settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            # Avoid interrupting the user flow if persistence fails.
            pass

    def _save_system_settings(self) -> None:
        try:
            settings = self._collect_system_settings_from_form()
            settings_path = self._system_settings_path()
            settings_path.parent.mkdir(parents=True, exist_ok=True)
            settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
            self._sync_review_filters_from_settings(settings)
            self.refresh_user_preview()
            QMessageBox.information(self, "הצלחה", "הגדרות המערכת נשמרו בהצלחה.")
        except Exception as error:
            QMessageBox.critical(self, "שגיאת הגדרות", f"לא ניתן לשמור את הגדרות המערכת.\n\n{error}")

    def _reset_system_settings_form(self) -> None:
        defaults = self._default_system_settings()
        self._load_system_settings_into_form(defaults)
        self._apply_system_settings_availability()

    def _available_selected_slots(self) -> set[str]:
        return {
            slot_key
            for slot_key, widget_data in self.slot_widgets.items()
            if list(widget_data.get("selected_paths", []))
        }

    def _apply_system_settings_availability(self) -> None:
        for section_key, section_widget in self.system_settings_sections.items():
            section_widget.setVisible(True)
            section_widget.setEnabled(True)
            if section_key in self.system_settings_unavailable_labels:
                self.system_settings_unavailable_labels[section_key].setVisible(False)
            section_widget.setToolTip("")

    def _build_issue_preview(self, issues: list) -> str:
        if not issues:
            return "ללא שגיאות"

        unique_messages: list[str] = []
        for issue in issues:
            if issue.message not in unique_messages:
                unique_messages.append(issue.message)
            if len(unique_messages) == 2:
                break

        preview = " | ".join(unique_messages)
        return preview if len(preview) <= 90 else f"{preview[:87]}..."

    @staticmethod
    def _is_intake_issue(issue: "ValidationIssue") -> bool:
        """Returns True only for structural/technical intake issues.

        Intake issues (shown in the intake log):
          - Missing required column  (row_number==0, message contains "עמודת חובה חסרה")
          - File structure mismatch  (row_number==0, message contains "אינו תואם למבנה")
          - Missing required column group (row_number==0, message contains "נדרשת לפחות")
          - Missing required value in a data row (row_number>0, message contains "ערך חובה חסר")

        NOT intake issues (audit/analysis findings shown in the analysis tab):
          - RSPARAM / TPFET policy violations
          - "לא נמצא פרמטר נדרש"
          - Any other business-logic row-level check
        """
        msg = issue.message
        if issue.row_number == 0:
            return (
                "עמודת חובה חסרה" in msg
                or "אינו תואם למבנה" in msg
                or "נדרשת לפחות" in msg
            )
        return "ערך חובה חסר" in msg

    @staticmethod
    def _compute_intake_summary(total_rows: int, intake_issues: list["ValidationIssue"]) -> tuple[int, int]:
        """Return (valid_rows, invalid_rows) for intake-only issues."""
        row_level_issues = {issue.row_number for issue in intake_issues if issue.row_number > 0}
        invalid_rows = len(row_level_issues)

        if any(issue.row_number == 0 for issue in intake_issues):
            invalid_rows = total_rows if total_rows else invalid_rows

        valid_rows = max(total_rows - invalid_rows, 0)
        return valid_rows, invalid_rows

    def _get_slot_category(self, slot_key: str) -> str:
        return str(self.SLOT_DEFINITIONS.get(slot_key, {}).get("sub_category", "לא סווג"))

    def _get_slot_display_name(self, slot_key: str) -> str:
        metadata = self.SLOT_DEFINITIONS.get(slot_key, {})
        return str(metadata.get("label", slot_key))

    def _get_slot_extraction_date(self, slot_key: str) -> str:
        widget_data = self.slot_widgets.get(slot_key, {})
        date_edit = widget_data.get("extraction_date_edit")
        if isinstance(date_edit, QLineEdit):
            date_text = date_edit.text().strip()
            return date_text or "לא צוין"
        return "לא צוין"

    @staticmethod
    def _is_a_dialog_user(user_type: object) -> bool:
        normalized_value = "" if user_type is None else str(user_type).strip().upper()
        return normalized_value in ValidationDesktopApp.USER_TYPE_RULES.get("Dialog", [])

    @staticmethod
    def _has_initial_password(password_flag: object) -> bool:
        normalized_value = "" if password_flag is None else str(password_flag).strip().upper()
        return normalized_value in {"X", "1", "TRUE", "YES", "Y"}

    def _is_user_locked(self, uflag: object) -> bool:
        normalized_value = "" if uflag is None else str(uflag).strip()
        if normalized_value in {"", "0", "00"}:
            return False
        try:
            return int(normalized_value) != 0
        except ValueError:
            return True

    def _validity_period_overlaps(
        self,
        gltgv_value: object,
        gltgb_value: object,
        period_start: date | None,
        period_end: date | None,
    ) -> bool:
        if period_start is None or period_end is None:
            return False

        gltgv = self._parse_user_preview_date(gltgv_value)
        gltgb = self._parse_user_preview_date(gltgb_value)
        if gltgv is None and gltgb is None:
            return False

        valid_from = gltgv.date() if gltgv is not None else datetime.min.date()
        valid_to = gltgb.date() if gltgb is not None else datetime.max.date()
        return valid_from <= period_end and valid_to >= period_start

    def _is_user_active_in_period(
        self,
        usr_entry: dict[str, str],
        period_start: date | None,
        period_end: date | None,
    ) -> bool:
        if period_start is None or period_end is None:
            return False

        last_login_date = self._parse_user_preview_date(usr_entry.get("TRDAT", ""))
        if last_login_date is not None and period_start <= last_login_date.date() <= period_end:
            return True

        return self._validity_period_overlaps(
            usr_entry.get("GLTGV", ""),
            usr_entry.get("GLTGB", ""),
            period_start,
            period_end,
        )

    def _is_generic_user(self, bname: object, settings: dict[str, Any]) -> bool:
        if not bname or not isinstance(settings, dict):
            return False
        normalized_name = str(bname).strip().casefold()
        generic_users = settings.get("generic_users", [])
        if not isinstance(generic_users, list):
            return False
        return any(str(item).strip().casefold() == normalized_name for item in generic_users)

    def _is_super_user(self, mandt: object, bname: object, settings: dict[str, Any]) -> bool:
        if not bname or not isinstance(settings, dict):
            return False
        normalized_mandt = str(mandt).strip()
        normalized_bname = str(bname).strip().casefold()
        super_users = settings.get("authorized_stms_users", [])
        if not super_users:
            super_users = settings.get("super_users", [])
        if not isinstance(super_users, list):
            return False
        for row in super_users:
            if not isinstance(row, dict):
                continue
            if str(row.get("MANDT", "")).strip() == normalized_mandt and str(row.get("BNAME", "")).strip().casefold() == normalized_bname:
                return True
        return False

    def _build_user_findings_description(self, usr_entry: dict[str, str], extraction_date_text: str) -> str:
        findings: list[str] = []
        settings = self._current_system_settings()
        period_cfg = settings.get("user_review_period", {}) if isinstance(settings, dict) else {}
        start_date = self._parse_user_preview_date(period_cfg.get("start_date", ""))
        end_date = self._parse_user_preview_date(period_cfg.get("end_date", ""))
        period_start = start_date.date() if start_date is not None else None
        period_end = end_date.date() if end_date is not None else None

        is_super_user = self._is_super_user(usr_entry.get("MANDT", ""), usr_entry.get("BNAME", ""), settings)
        is_generic_user = self._is_generic_user(usr_entry.get("BNAME", ""), settings)
        is_locked = self._is_user_locked(usr_entry.get("UFLAG", ""))
        is_active = self._is_user_active_in_period(usr_entry, period_start, period_end)

        if (is_super_user or is_generic_user) and not is_locked and is_active:
            if is_super_user and is_generic_user:
                findings.append("משתמש על / גנרי פעיל ולא נעול")
            elif is_super_user:
                findings.append("משתמש על פעיל ולא נעול")
            else:
                findings.append("משתמש גנרי פעיל ולא נעול")

        if self._is_a_dialog_user(usr_entry.get("USTYP", "")):
            inactivity_threshold = self._safe_int(settings.get("inactive_days_threshold", 90), 90)
            password_policy = settings.get("password_policy_defaults", {}) if isinstance(settings, dict) else {}
            max_password_age_days = self._safe_int(
                password_policy.get("max_password_age_days", 90) if isinstance(password_policy, dict) else 90,
                90,
            )
            initial_password_change_max_days = self._safe_int(
                password_policy.get("initial_password_change_max_days", 2) if isinstance(password_policy, dict) else 2,
                2,
            )

            extraction_date = self._parse_user_preview_date(extraction_date_text)
            last_login_date = self._parse_user_preview_date(usr_entry.get("TRDAT", ""))
            password_change_date = self._parse_user_preview_date(usr_entry.get("PWDCHGDATE", ""))
            password_set_date = self._parse_user_preview_date(usr_entry.get("PWDSETDATE", ""))

            if extraction_date is not None and last_login_date is not None:
                inactivity_days = (extraction_date.date() - last_login_date.date()).days
                if inactivity_days > inactivity_threshold:
                    findings.append(f"משתמש לא פעיל מעל {inactivity_threshold} יום")

            if last_login_date is not None and password_change_date is not None:
                password_gap_days = (last_login_date.date() - password_change_date.date()).days
                if password_gap_days > max_password_age_days:
                    findings.append(f"סיסמה לא הוחלפה מעל {max_password_age_days} יום")

            if extraction_date is not None and password_set_date is not None and self._has_initial_password(usr_entry.get("PWDINITIAL", "")):
                initial_password_age_days = (extraction_date.date() - password_set_date.date()).days
                if initial_password_age_days > initial_password_change_max_days:
                    if initial_password_change_max_days == 2:
                        findings.append("סיסמה ראשונית לא הוחלפה תוך 48 שעות")
                    else:
                        findings.append(f"סיסמה ראשונית לא הוחלפה תוך {initial_password_change_max_days} ימים")

        return " | ".join(findings)

    @staticmethod
    def _export_sort_key(preview_row: dict[str, str]) -> int:
        review_status = preview_row.get("REVIEW_STATUS", "").strip()
        has_findings = bool(preview_row.get("FINDINGS_DESCRIPTION", "").strip())
        if review_status == "טרם נבדק":
            return 1
        if review_status == "נבדק - לא תקין":
            return 2
        if review_status == "נבדק - תקין" and has_findings:
            return 3
        return 4

    @staticmethod
    def _has_review_note(technical_note: object, business_note: object) -> bool:
        return bool(str(technical_note or "").strip() or str(business_note or "").strip())

    def _is_user_review_complete(
        self,
        review_status: object,
        findings_description: object,
        technical_note: object,
        business_note: object,
    ) -> bool:
        normalized_status = self._normalize_reviewer_status(review_status)
        if normalized_status not in self.REVIEWED_STATUSES:
            return False

        if normalized_status == "נבדק - לא תקין":
            return self._has_review_note(technical_note, business_note)

        if str(findings_description or "").strip():
            return self._has_review_note(technical_note, business_note)

        return True

    def _load_all_user_preview_rows(self) -> list[dict[str, str]]:
        usr02_rows = self._load_preview_rows("USR02")
        combined_rows = self._load_preview_rows("ADR6_USR21")
        return self._build_user_preview_rows(usr02_rows, combined_rows)

    def _get_user_review_completion_snapshot(self) -> tuple[list[dict[str, str]], int, list[dict[str, str]]]:
        preview_rows = self._load_all_user_preview_rows()
        incomplete_rows: list[dict[str, str]] = []
        reviewed_rows = 0
        for preview_row in preview_rows:
            if self._is_user_review_complete(
                preview_row.get("REVIEW_STATUS", ""),
                preview_row.get("FINDINGS_DESCRIPTION", ""),
                preview_row.get("TECH_REVIEW_NOTES", ""),
                preview_row.get("BUS_REVIEW_NOTES", ""),
            ):
                reviewed_rows += 1
            else:
                incomplete_rows.append(preview_row)
        return preview_rows, reviewed_rows, incomplete_rows

    def _build_user_review_incomplete_reason(self, preview_row: dict[str, str]) -> str:
        review_status = self._normalize_reviewer_status(preview_row.get("REVIEW_STATUS", ""))
        findings_description = str(preview_row.get("FINDINGS_DESCRIPTION", "")).strip()
        if review_status not in self.REVIEWED_STATUSES:
            return "סטטוס הסקירה עדיין אינו מסומן כמשתמש שנבדק."
        if review_status == "נבדק - לא תקין":
            return "המשתמש סומן כלא תקין אך לא הוזנה הערה טכנית או עסקית."
        if findings_description:
            return "המשתמש סומן כתקין למרות שקיים ממצא, אך לא הוזנה הערה טכנית או עסקית."
        return "הסקירה טרם הושלמה בהתאם לכלל ההשלמה שהוגדר."

    def _update_review_row_highlight(self, row_index: int, preview_row: dict[str, str] | None = None) -> None:
        review_status_col: int | None = None
        technical_notes_col: int | None = None
        business_notes_col: int | None = None
        for col_idx, field_name in enumerate(self.user_preview_visible_columns):
            if field_name == "REVIEW_STATUS":
                review_status_col = col_idx
            elif field_name in {"TECH_REVIEW_NOTES", "REVIEW_NOTES"}:
                technical_notes_col = col_idx
            elif field_name == "BUS_REVIEW_NOTES":
                business_notes_col = col_idx

        if preview_row is not None:
            review_status_text = str(preview_row.get("REVIEW_STATUS", "")).strip()
            findings_text = str(preview_row.get("FINDINGS_DESCRIPTION", "")).strip()
            technical_notes_text = str(preview_row.get("TECH_REVIEW_NOTES", "")).strip()
            business_notes_text = str(preview_row.get("BUS_REVIEW_NOTES", "")).strip()
        else:
            review_status_text = ""
            if review_status_col is not None:
                combo = self.user_preview_table.cellWidget(row_index, review_status_col)
                if isinstance(combo, QComboBox):
                    review_status_text = self.format_rtl_text(combo.currentText())
                else:
                    status_item = self.user_preview_table.item(row_index, review_status_col)
                    if status_item is not None:
                        review_status_text = status_item.text().strip()

            findings_text = ""
            try:
                findings_col = self.user_preview_visible_columns.index("FINDINGS_DESCRIPTION")
            except ValueError:
                findings_col = -1
            if findings_col >= 0:
                findings_item = self.user_preview_table.item(row_index, findings_col)
                if findings_item is not None:
                    findings_text = findings_item.text().strip()

            technical_notes_text = ""
            if technical_notes_col is not None:
                notes_item = self.user_preview_table.item(row_index, technical_notes_col)
                if notes_item is not None:
                    technical_notes_text = notes_item.text().strip()

            business_notes_text = ""
            if business_notes_col is not None:
                business_notes_item = self.user_preview_table.item(row_index, business_notes_col)
                if business_notes_item is not None:
                    business_notes_text = business_notes_item.text().strip()

        is_not_reviewed = review_status_text == "טרם נבדק"
        is_review_complete = self._is_user_review_complete(
            review_status_text,
            findings_text,
            technical_notes_text,
            business_notes_text,
        )
        needs_warning = (not is_not_reviewed) and (not is_review_complete)

        unreviewed_color = QColor("#d6e8ff")
        warning_color = QColor("#fff0c2")
        clear_color = QColor(0, 0, 0, 0)

        for col_idx, field_name in enumerate(self.user_preview_visible_columns):
            combo = self.user_preview_table.cellWidget(row_index, col_idx)
            item = self.user_preview_table.item(row_index, col_idx)

            if is_not_reviewed:
                if isinstance(combo, QComboBox):
                    combo.setStyleSheet("background-color: #d6e8ff;")
                elif item is not None:
                    item.setBackground(unreviewed_color)
                continue

            if needs_warning and field_name in {"REVIEW_STATUS", "TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES", "REVIEW_NOTES"}:
                if isinstance(combo, QComboBox):
                    combo.setStyleSheet("background-color: #fff0c2;")
                elif item is not None:
                    item.setBackground(warning_color)
                continue

            if isinstance(combo, QComboBox):
                combo.setStyleSheet("")
            elif item is not None:
                if field_name == "STATUS":
                    status_value = item.text()
                    if status_value == "פעיל":
                        item.setBackground(QColor("#eaf7ea"))
                    elif "אי-התאמה" in status_value:
                        item.setBackground(QColor("#fff4cc"))
                    elif status_value == "נעול":
                        item.setBackground(QColor("#fdecec"))
                    else:
                        item.setBackground(clear_color)
                else:
                    item.setBackground(clear_color)

    def _get_user_preview_row_review_status(self, row_index: int) -> str:
        try:
            review_status_col = self.user_preview_visible_columns.index("REVIEW_STATUS")
        except ValueError:
            return self.DEFAULT_REVIEW_STATUS

        combo = self.user_preview_table.cellWidget(row_index, review_status_col)
        if isinstance(combo, QComboBox):
            return self._normalize_reviewer_status(combo.currentText())

        item = self.user_preview_table.item(row_index, review_status_col)
        if item is not None:
            return self._normalize_reviewer_status(item.text())
        return self.DEFAULT_REVIEW_STATUS

    def _update_user_review_progress_summary(self, total: int, reviewed: int, unreviewed: int) -> None:
        total = max(0, int(total))
        reviewed = max(0, int(reviewed))
        unreviewed = max(0, int(unreviewed))
        if reviewed + unreviewed != total:
            unreviewed = max(0, total - reviewed)

        percent_complete = int(round((reviewed / total) * 100)) if total > 0 else 0

        self.user_review_total_label.setText(self.format_ui_rtl_text(f"סה\"כ משתמשים בדוח: {total}"))
        self.user_review_reviewed_label.setText(self.format_ui_rtl_text(f"משתמשים שנבדקו: {reviewed}"))
        self.user_review_unreviewed_label.setText(self.format_ui_rtl_text(f"משתמשים שטרם נבדקו: {unreviewed}"))
        self.user_review_progress_percent_label.setText(
            self.format_ui_rtl_text(f"התקדמות השלמת סקירה: {percent_complete}%")
        )

        self.user_review_progress_bar.setMaximum(max(total, 1))
        self.user_review_progress_bar.setValue(min(reviewed, max(total, 1)))
        self.user_review_progress_bar.setFormat(f"{percent_complete}%")

    def _refresh_user_review_progress_summary_from_table(self) -> None:
        preview_rows, reviewed_rows, incomplete_rows = self._get_user_review_completion_snapshot()
        total_rows = len(preview_rows)
        self._update_user_review_progress_summary(total_rows, reviewed_rows, len(incomplete_rows))

    def _update_slot_path_label(self, slot_key: str, file_paths: list[str] | None = None) -> None:
        widget_data = self.slot_widgets.get(slot_key, {})
        label = widget_data.get("path_label")
        if not isinstance(label, QLabel):
            return

        paths = file_paths if file_paths is not None else list(widget_data.get("selected_paths", []))
        if not paths:
            label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            label.setText(self.format_ui_rtl_text("טרם נבחר קובץ"))
            return

        if len(paths) == 1:
            label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
            label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            label.setText(self.format_rtl_text(paths[0]))
            return

        label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        label.setText(self.format_ui_rtl_text(self._format_selected_files(paths)))

    def _remember_slot_load(self, slot_key: str) -> None:
        self.load_history = [item for item in self.load_history if item != slot_key]
        if list(self.slot_widgets.get(slot_key, {}).get("selected_paths", [])):
            self.load_history.append(slot_key)

    def clear_slot_selection(self, slot_key: str) -> None:
        if slot_key not in self.slot_widgets:
            return

        self.slot_widgets[slot_key]["selected_paths"] = []
        self._update_slot_path_label(slot_key, [])
        self.load_history = [item for item in self.load_history if item != slot_key]

        if self.selected_slot_key == slot_key:
            self.selected_slot_key = self.load_history[-1] if self.load_history else None
            if self.selected_slot_key:
                self.required_columns_edit.setText(self._suggest_required_columns(self.selected_slot_key))
            else:
                self.required_columns_edit.setText("")

        if slot_key in self.USER_PREVIEW_SLOTS:
            self.refresh_user_preview()
        self._apply_system_settings_availability()

    def clear_last_loaded_slot(self) -> None:
        while self.load_history:
            last_slot_key = self.load_history[-1]
            if list(self.slot_widgets.get(last_slot_key, {}).get("selected_paths", [])):
                self.clear_slot_selection(last_slot_key)
                return
            self.load_history.pop()

    def choose_file(self, slot_key: str) -> None:
        initial_directory = self._get_last_file_dialog_directory()
        if slot_key in self.MULTI_FILE_SLOTS:
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                f"בחירת קבצים עבור {slot_key}",
                initial_directory,
                "Supported files (*.txt *.csv *.xlsx *.xlsm);;All files (*.*)",
            )
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                f"בחירת קובץ עבור {slot_key}",
                initial_directory,
                "Supported files (*.txt *.csv *.xlsx *.xlsm);;All files (*.*)",
            )
            file_paths = [file_path] if file_path else []

        if file_paths:
            self._save_last_file_dialog_directory(Path(file_paths[0]).parent)
            self.selected_slot_key = slot_key
            self.slot_widgets[slot_key]["selected_paths"] = file_paths
            self._remember_slot_load(slot_key)
            self._update_slot_path_label(slot_key, file_paths)
            self.required_columns_edit.setText(self._suggest_required_columns(slot_key))
            if slot_key in self.USER_PREVIEW_SLOTS:
                self.refresh_user_preview()
            self._apply_system_settings_availability()

    def _parse_required_columns(self, raw_value: str | None = None) -> list[str]:
        value = self.required_columns_edit.text() if raw_value is None else raw_value
        normalized_value = value.replace(";", ",").replace("\n", ",")
        return [item.strip() for item in normalized_value.split(",") if item.strip()]

    def _required_columns_for_slot(self, slot_key: str) -> list[str]:
        if self.selected_slot_key == slot_key and self.required_columns_edit.text().strip():
            return self._parse_required_columns()
        return self._parse_required_columns(self._suggest_required_columns(slot_key))

    def _get_category_slots(self, sub_category: str) -> list[str]:
        return [
            slot_key
            for slot_key, metadata in self.SLOT_DEFINITIONS.items()
            if metadata.get("sub_category") == sub_category
        ]

    def _get_domain_slots(self, domain: str) -> list[str]:
        return [
            slot_key
            for slot_key, metadata in self.SLOT_DEFINITIONS.items()
            if metadata.get("domain") == domain
        ]

    def _current_file_paths(self) -> list[str]:
        if not self.selected_slot_key:
            return []
        return list(self.slot_widgets[self.selected_slot_key].get("selected_paths", []))

    def _build_input_files_dict(
        self,
        preferred_slot_key: str | None = None,
        fallback_file_paths: Sequence[str | Path] | None = None,
    ) -> dict[str, list[str | Path]]:
        """Collect selected input files from all slot widgets for pipeline execution."""
        input_files: dict[str, list[str | Path]] = {}

        if preferred_slot_key and fallback_file_paths:
            normalized_fallback: list[str | Path] = [
                str(path).strip()
                for path in fallback_file_paths
                if str(path).strip()
            ]
            if normalized_fallback:
                input_files[preferred_slot_key] = normalized_fallback

        for slot_key, widget_data in self.slot_widgets.items():
            selected_paths: list[str | Path] = []
            raw_paths = widget_data.get("selected_paths", [])

            if isinstance(raw_paths, str):
                if raw_paths.strip():
                    selected_paths = [raw_paths.strip()]
            elif isinstance(raw_paths, (list, tuple, set)):
                selected_paths = [str(path).strip() for path in raw_paths if str(path).strip()]

            if selected_paths:
                input_files[slot_key] = selected_paths

        return input_files

    def _format_selected_files(self, file_paths: list[str]) -> str:
        if len(file_paths) == 1:
            return self.format_rtl_text(file_paths[0])

        preview_names = [Path(path).name for path in file_paths[:3]]
        suffix = "" if len(file_paths) <= 3 else " ..."
        return self.format_rtl_text(
            f"נבחרו {len(file_paths)} קבצים: {', '.join(preview_names)}{suffix}"
        )

    def _suggest_required_columns(self, slot_key: str) -> str:
        suggestions = {
            "USR02": "BNAME,UFLAG,TRDAT,LTIME",
            "ADR6_USR21": "",
            "AGR_USERS": "AGR_NAME,UNAME",
            "AGR_1251": "AGR_NAME,OBJECT,FIELD",
            "AGR_1252": "AGR_NAME,LOW",
            "AGR_DEFINE": "AGR_NAME,PARENT_AGR",
            "UST04": "BNAME,PROFILE",
            "E070": "TRKORR,AS4USER,TRFUNCTION",
            "T000": "MANDT,CCCATEGORY",
            "STMS": "TRKORR",
            "RSPARAM": "PARAMETER,VALUE",
            "TPFET": "PARAMETER,VALUE",
        }
        return suggestions.get(slot_key, "")

    def _file_dialog_state_path(self) -> Path:
        return self.config.output_dir / "file_dialog_state.json"

    def _load_last_file_dialog_directory(self) -> Path:
        default_directory = self.config.input_dir
        if not self._allow_user_preview_persistence:
            return default_directory

        state_path = self._file_dialog_state_path()
        if not state_path.exists():
            return default_directory

        try:
            raw_data = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            return default_directory

        saved_directory = ""
        if isinstance(raw_data, dict):
            saved_directory = str(raw_data.get("last_directory", "")).strip()

        candidate_directory = Path(saved_directory).expanduser() if saved_directory else default_directory
        if candidate_directory.exists() and candidate_directory.is_dir():
            return candidate_directory
        return default_directory

    def _save_last_file_dialog_directory(self, directory_path: object) -> None:
        if directory_path is None:
            return

        candidate_directory = Path(str(directory_path)).expanduser()
        if candidate_directory.is_file():
            candidate_directory = candidate_directory.parent
        if not candidate_directory.exists() or not candidate_directory.is_dir():
            return

        self.last_file_dialog_directory = candidate_directory
        if not self._allow_user_preview_persistence:
            return

        state_path = self._file_dialog_state_path()
        payload = {"last_directory": str(candidate_directory)}
        state_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _get_last_file_dialog_directory(self) -> str:
        candidate_directory = getattr(self, "last_file_dialog_directory", self.config.input_dir)
        if not isinstance(candidate_directory, Path) or not candidate_directory.exists() or not candidate_directory.is_dir():
            candidate_directory = self.config.input_dir
        return str(candidate_directory)

    def _user_preview_settings_path(self) -> Path:
        return self.config.output_dir / "user_preview_columns.json"

    def _user_reviewer_state_path(self) -> Path:
        return self.config.output_dir / "user_preview_reviewer_state.json"

    @staticmethod
    def _user_reviewer_state_key(mandt: object, bname: object) -> str:
        mandt_value = "" if mandt is None else str(mandt).strip()
        bname_value = "" if bname is None else str(bname).strip()
        return f"{mandt_value}|{bname_value}"

    @classmethod
    def _normalize_reviewer_status(cls, value: object) -> str:
        normalized_value = "" if value is None else str(value).strip()
        if normalized_value in cls.REVIEW_STATUS_OPTIONS:
            return normalized_value
        return cls.DEFAULT_REVIEW_STATUS

    @classmethod
    def _default_reviewer_values(cls) -> dict[str, str]:
        return {
            "REVIEW_STATUS": cls.DEFAULT_REVIEW_STATUS,
            "TECH_REVIEW_NOTES": "",
            "BUS_REVIEW_NOTES": "",
        }

    def _load_user_reviewer_state(self) -> dict[str, dict[str, str]]:
        if not self._allow_user_preview_persistence:
            return {}

        state_path = self._user_reviewer_state_path()
        if not state_path.exists():
            return {}

        try:
            raw_data = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

        if not isinstance(raw_data, dict):
            return {}

        normalized_state: dict[str, dict[str, str]] = {}
        for review_key, review_values in raw_data.items():
            if not isinstance(review_values, dict):
                continue
            legacy_notes = str(review_values.get("REVIEW_NOTES", "")).strip()
            normalized_state[str(review_key)] = {
                "REVIEW_STATUS": self._normalize_reviewer_status(review_values.get("REVIEW_STATUS")),
                "TECH_REVIEW_NOTES": str(review_values.get("TECH_REVIEW_NOTES", "")).strip() or legacy_notes,
                "BUS_REVIEW_NOTES": str(review_values.get("BUS_REVIEW_NOTES", "")).strip(),
            }
        return normalized_state

    def _save_user_reviewer_state(self) -> None:
        if not self._allow_user_preview_persistence:
            return

        state_path = self._user_reviewer_state_path()
        state_path.write_text(
            json.dumps(self.user_reviewer_state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _get_reviewer_values(self, mandt: object, bname: object) -> dict[str, str]:
        review_key = self._user_reviewer_state_key(mandt, bname)
        stored_values = self.user_reviewer_state.get(review_key)
        if not isinstance(stored_values, dict):
            return self._default_reviewer_values().copy()
        legacy_notes = str(stored_values.get("REVIEW_NOTES", "")).strip()
        return {
            "REVIEW_STATUS": self._normalize_reviewer_status(stored_values.get("REVIEW_STATUS")),
            "TECH_REVIEW_NOTES": str(stored_values.get("TECH_REVIEW_NOTES", "")).strip() or legacy_notes,
            "BUS_REVIEW_NOTES": str(stored_values.get("BUS_REVIEW_NOTES", "")).strip(),
        }

    def _update_reviewer_value(self, review_key: str, field_name: str, value: object) -> None:
        normalized_field = "TECH_REVIEW_NOTES" if field_name == "REVIEW_NOTES" else field_name
        if not review_key or normalized_field not in {"REVIEW_STATUS", "TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES"}:
            return

        current_values = self.user_reviewer_state.setdefault(review_key, self._default_reviewer_values().copy())
        if normalized_field == "REVIEW_STATUS":
            current_values[normalized_field] = self._normalize_reviewer_status(value)
        else:
            current_values[normalized_field] = "" if value is None else str(value).strip()
        self._save_user_reviewer_state()

    def _normalize_user_preview_columns(self, selected_columns: list[str] | None) -> list[str]:
        allowed_fields = [column["field"] for column in self.USER_PREVIEW_COLUMN_DEFINITIONS]
        if not selected_columns:
            return list(self.DEFAULT_USER_PREVIEW_COLUMNS)

        normalized_input = ["TECH_REVIEW_NOTES" if field == "REVIEW_NOTES" else field for field in selected_columns]
        normalized = [field for field in allowed_fields if field in normalized_input]
        return normalized or list(self.DEFAULT_USER_PREVIEW_COLUMNS)

    def _load_user_preview_column_selection(self) -> list[str]:
        if not self._allow_user_preview_persistence:
            return list(self.DEFAULT_USER_PREVIEW_COLUMNS)

        settings_path = self._user_preview_settings_path()
        if not settings_path.exists():
            return list(self.DEFAULT_USER_PREVIEW_COLUMNS)

        try:
            raw_data = json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception:
            return list(self.DEFAULT_USER_PREVIEW_COLUMNS)

        loaded_columns = list(raw_data.get("visible_columns", [])) if isinstance(raw_data, dict) else []
        settings_version = int(raw_data.get("version", 0)) if isinstance(raw_data, dict) else 0

        for version in range(settings_version + 1, self.CURRENT_USER_PREVIEW_SETTINGS_VERSION + 1):
            for field_name in self.USER_PREVIEW_SETTINGS_MIGRATIONS.get(version, []):
                if field_name not in loaded_columns:
                    loaded_columns.append(field_name)

        return self._normalize_user_preview_columns(loaded_columns)

    def _save_user_preview_column_selection(self) -> None:
        if not self._allow_user_preview_persistence:
            return

        settings_path = self._user_preview_settings_path()
        payload = {
            "version": self.CURRENT_USER_PREVIEW_SETTINGS_VERSION,
            "visible_columns": self.user_preview_visible_columns,
        }
        settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _get_user_preview_column_definition(self, field_name: str) -> dict[str, Any]:
        for column in self.USER_PREVIEW_COLUMN_DEFINITIONS:
            if column["field"] == field_name:
                return column
        return {"field": field_name, "formal": field_name, "technical": field_name, "source": "לא ידוע", "width": 120}

    def _handle_user_preview_item_changed(self, item: QTableWidgetItem) -> None:
        if self._refreshing_user_preview or item is None:
            return

        field_name = item.data(Qt.ItemDataRole.UserRole + 1)
        review_key = item.data(Qt.ItemDataRole.UserRole)
        if field_name not in {"TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES", "REVIEW_NOTES"} or not review_key:
            return

        normalized_text = self.format_rtl_text(item.text())
        item.setToolTip(normalized_text)
        self._update_reviewer_value(str(review_key), str(field_name), normalized_text)
        self._update_review_row_highlight(item.row())

    def _get_user_preview_cell_text(self, row_index: int, column_index: int) -> str:
        cell_widget = self.user_preview_table.cellWidget(row_index, column_index)
        if isinstance(cell_widget, QComboBox):
            return self.format_rtl_text(cell_widget.currentText())

        item = self.user_preview_table.item(row_index, column_index)
        if item is None:
            return ""
        return self.format_rtl_text(item.text())

    def _configure_user_preview_table(self) -> None:
        self.user_preview_visible_columns = self._normalize_user_preview_columns(self.user_preview_visible_columns)
        self.user_preview_table.setColumnCount(len(self.user_preview_visible_columns))
        self.user_preview_table.setHorizontalHeaderLabels([
            str(self._get_user_preview_column_definition(field_name).get("formal", field_name))
            for field_name in self.user_preview_visible_columns
        ])
        header = self.user_preview_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        header.setSectionsMovable(False)
        header.setSectionsClickable(True)
        header.setMinimumSectionSize(70)
        self.user_preview_table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        for column_index, field_name in enumerate(self.user_preview_visible_columns):
            header.setSectionResizeMode(column_index, QHeaderView.ResizeMode.Interactive)
            default_width = int(self._get_user_preview_column_definition(field_name).get("width", 120))
            self.user_preview_table.setColumnWidth(column_index, default_width)
        self.user_preview_table.setSortingEnabled(True)

    def _create_user_preview_columns_dialog(self) -> tuple[QDialog, QTableWidget]:
        dialog = QDialog(self)
        dialog.setWindowTitle(self.format_rtl_text("בחירת עמודות לסקירת משתמשים"))
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(720, 460)

        layout = QVBoxLayout(dialog)
        hint_label = QLabel(
            self.format_ui_rtl_text("סמן את העמודות שברצונך להציג. לחיצה על OK תרענן את הטבלה, ו-Cancel תשאיר את המצב הקיים.")
        )
        hint_label.setWordWrap(True)
        hint_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(hint_label)

        selection_table = QTableWidget(len(self.USER_PREVIEW_COLUMN_DEFINITIONS), 3)
        selection_table.setHorizontalHeaderLabels(["שם פורמלי", "שם טכני", "הצג"])
        selection_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        selection_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        selection_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        selection_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        selection_table.verticalHeader().setVisible(False)
        selection_table.setAlternatingRowColors(True)
        selection_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

        for row_index, column in enumerate(self.USER_PREVIEW_COLUMN_DEFINITIONS):
            formal_item = QTableWidgetItem(self.format_rtl_text(str(column["formal"])))
            formal_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            formal_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            technical_item = QTableWidgetItem(str(column["technical"]))
            technical_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            technical_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            checkbox_item = QTableWidgetItem("")
            checkbox_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsUserCheckable)
            checkbox_item.setCheckState(Qt.CheckState.Checked if column["field"] in self.user_preview_visible_columns else Qt.CheckState.Unchecked)
            checkbox_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            checkbox_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            selection_table.setItem(row_index, 0, formal_item)
            selection_table.setItem(row_index, 1, technical_item)
            selection_table.setItem(row_index, 2, checkbox_item)

        layout.addWidget(selection_table)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        return dialog, selection_table

    def _get_selected_user_preview_columns(self, selection_table: QTableWidget) -> list[str]:
        selected_columns: list[str] = []
        for row_index, column in enumerate(self.USER_PREVIEW_COLUMN_DEFINITIONS):
            checkbox_item = selection_table.item(row_index, 2)
            if checkbox_item is not None and checkbox_item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(str(column["field"]))
        return selected_columns

    def _apply_user_preview_columns(self, selected_columns: list[str]) -> None:
        normalized_columns = self._normalize_user_preview_columns(selected_columns)
        if not normalized_columns:
            QMessageBox.warning(self, "בחירת עמודות", "יש לבחור לפחות עמודה אחת להצגה בטבלת הסקירה.")
            return

        self.user_preview_visible_columns = normalized_columns
        self._save_user_preview_column_selection()
        self._configure_user_preview_table()
        self.refresh_user_preview()

    def show_user_preview_column_dialog(self) -> None:
        dialog, selection_table = self._create_user_preview_columns_dialog()
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        selected_columns = self._get_selected_user_preview_columns(selection_table)
        if not selected_columns:
            QMessageBox.warning(self, "בחירת עמודות", "יש לבחור לפחות עמודה אחת להצגה בטבלת הסקירה.")
            return

        self._apply_user_preview_columns(selected_columns)

    @staticmethod
    def _get_row_value(row: dict[str, Any], *candidates: str) -> str:
        normalized_row = {
            str(key).strip().upper(): value
            for key, value in row.items()
            if not str(key).startswith("__")
        }
        for candidate in candidates:
            for alias in get_column_aliases(candidate):
                if alias in normalized_row:
                    value = normalized_row[alias]
                    if value is None:
                        continue
                    return str(value).strip()
        return ""

    def _get_authorized_stms_users(self) -> list[str]:
        table = self.system_settings_widgets.get("authorized_stms_users")
        if not isinstance(table, QTableWidget):
            table = self.system_settings_widgets.get("super_users")
        if not isinstance(table, QTableWidget):
            return []

        users: list[str] = []
        for row_index in range(table.rowCount()):
            bname_item = table.item(row_index, 1)
            if isinstance(bname_item, QTableWidgetItem):
                bname = bname_item.text().strip().upper()
                if bname and bname not in users:
                    users.append(bname)
        return users

    def _load_preview_rows(self, slot_key: str) -> list[dict[str, Any]]:
        file_paths = list(self.slot_widgets.get(slot_key, {}).get("selected_paths", []))
        if not file_paths:
            return []

        preview_rows: list[dict[str, Any]] = []
        text_reader = TextFileReader()
        excel_reader = ExcelFileReader()

        for raw_path in file_paths:
            try:
                path = Path(raw_path)
                suffix = path.suffix.lower()
                if suffix in {".txt", ".csv"}:
                    rows = text_reader.read(path)
                elif suffix in {".xlsx", ".xlsm"}:
                    rows = excel_reader.read(path)
                else:
                    continue

                preview_rows.extend(
                    {
                        **row,
                        "__source_file": path.name,
                    }
                    for row in rows
                )
            except Exception:
                continue

        return preview_rows

    @staticmethod
    def _format_user_status(flag_value: object) -> str:
        normalized_value = "" if flag_value is None else str(flag_value).strip()
        if normalized_value in {"", "0", "00"}:
            return "פעיל"
        if normalized_value in {"64", "128", "129"}:
            return "נעול"
        return normalized_value

    @staticmethod
    def _parse_user_preview_date(raw_value: object) -> datetime | None:
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

    @classmethod
    def _format_user_preview_value_for_display(cls, field_name: str, value: object) -> str:
        _ = field_name
        return "" if value is None else str(value).strip()

    @classmethod
    def _get_user_preview_sort_value(cls, field_name: str, value: object) -> str:
        normalized_value = "" if value is None else str(value).strip()
        if field_name in cls.USER_PREVIEW_DATE_FIELDS:
            parsed_date = cls._parse_user_preview_date(normalized_value)
            return parsed_date.strftime("%Y%m%d") if parsed_date is not None else ""
        return normalized_value.casefold()

    def _get_user_preview_filter_mode(self) -> str:
        filter_widget = getattr(self, "user_preview_status_filter", None)
        if isinstance(filter_widget, QComboBox):
            selected_value = filter_widget.currentData()
            if selected_value:
                return str(selected_value)
        return "all"

    def _filter_user_preview_rows(self, preview_rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], str]:
        filter_mode = self._get_user_preview_filter_mode()
        if filter_mode == "all":
            return preview_rows, ""

        start_text = self.audit_period_from_edit.text().strip() if hasattr(self, "audit_period_from_edit") else ""
        end_text = self.audit_period_to_edit.text().strip() if hasattr(self, "audit_period_to_edit") else ""
        if not start_text or not end_text:
            return preview_rows, "כדי לסנן לפי פעילות בתקופה יש להזין תאריך התחלה ותאריך סיום."
        if not start_text:
            start_text = self._default_extraction_date()
        if not end_text:
            end_text = self._default_extraction_date()

        start_date = self._parse_user_preview_date(start_text)
        end_date = self._parse_user_preview_date(end_text)
        if start_date is None or end_date is None:
            return preview_rows, "יש להזין את טווח התאריכים בפורמט YYYY-MM-DD."
        if start_date > end_date:
            return preview_rows, "תאריך ההתחלה חייב להיות מוקדם או שווה לתאריך הסיום."

        filtered_rows: list[dict[str, str]] = []
        for preview_row in preview_rows:
            last_login_date = self._parse_user_preview_date(preview_row.get("TRDAT", ""))
            was_active_in_period = last_login_date is not None and start_date <= last_login_date <= end_date
            if filter_mode == "active" and was_active_in_period:
                filtered_rows.append(preview_row)
            elif filter_mode == "inactive" and not was_active_in_period:
                filtered_rows.append(preview_row)

        return filtered_rows, ""

    def _build_user_preview_rows(
        self,
        usr02_rows: list[dict[str, Any]],
        combined_rows: list[dict[str, Any]],
    ) -> list[dict[str, str]]:
        usr02_map: dict[tuple[str, str], dict[str, str]] = {}
        addr_users_map: dict[tuple[str, str], dict[str, str]] = {}
        email_by_addr: dict[str, str] = {}
        email_by_pers: dict[str, str] = {}

        for row in usr02_rows:
            mandt = self._get_row_value(row, "MANDT")
            bname = self._get_row_value(row, "BNAME")
            if not bname:
                continue
            raw_uflag = self._get_row_value(row, "UFLAG")
            usr02_map[(mandt, bname)] = {
                "MANDT": mandt,
                "BNAME": bname,
                "UFLAG": raw_uflag,
                "STATUS": self._format_user_status(raw_uflag),
                "TRDAT": self._get_row_value(row, "TRDAT"),
                "LTIME": self._get_row_value(row, "LTIME"),
                "GLTGV": self._get_row_value(row, "GLTGV"),
                "GLTGB": self._get_row_value(row, "GLTGB"),
                "USTYP": self._get_row_value(row, "USTYP"),
                "LOCNT": self._get_row_value(row, "LOCNT"),
                "PWDINITIAL": self._get_row_value(row, "PWDINITIAL"),
                "PWDCHGDATE": self._get_row_value(row, "PWDCHGDATE"),
                "PWDSETDATE": self._get_row_value(row, "PWDSETDATE"),
                "OCOD1": self._get_row_value(row, "OCOD1"),
                "PASSCODE": self._get_row_value(row, "PASSCODE"),
                "PWDSALTEDHASH": self._get_row_value(row, "PWDSALTEDHASH"),
                "SECURITY_POLICY": self._get_row_value(row, "SECURITY_POLICY"),
            }

        for row in combined_rows:
            addrnumber = self._get_row_value(row, "ADDRNUMBER")
            persnumber = self._get_row_value(row, "PERSNUMBER")
            smtp_addr = self._get_row_value(row, "SMTP_ADDR")

            if smtp_addr:
                if addrnumber:
                    email_by_addr[addrnumber] = smtp_addr
                if persnumber:
                    email_by_pers[persnumber] = smtp_addr

            bname = self._get_row_value(row, "BNAME")
            if not bname:
                continue

            mandt = self._get_row_value(row, "MANDT")
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

            for field_name in ["NAME_FIRST", "NAME_LAST", "NAME_TEXTC", "COMPANY", "DEPARTMENT", "ADDRNUMBER", "PERSNUMBER", "SMTP_ADDR"]:
                field_value = self._get_row_value(row, field_name)
                if field_value and not current_entry[field_name]:
                    current_entry[field_name] = field_value

        if usr02_map:
            ordered_keys = sorted(list(usr02_map.keys()), key=lambda item: (item[0], item[1]))
        else:
            ordered_keys = sorted(list(addr_users_map.keys()), key=lambda item: (item[0], item[1]))

        preview_rows: list[dict[str, str]] = []
        extraction_date_text = self._get_slot_extraction_date("USR02")
        work_environment_label = self._current_work_environment_label()

        for key in ordered_keys:
            usr_entry = usr02_map.get(key, {})
            addr_entry = addr_users_map.get(key, {})
            merged_mandt = usr_entry.get("MANDT") or addr_entry.get("MANDT", "")
            merged_bname = usr_entry.get("BNAME") or addr_entry.get("BNAME", "")
            review_values = self._get_reviewer_values(merged_mandt, merged_bname)
            findings_description = self._build_user_findings_description(usr_entry, extraction_date_text)
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
                    "REVIEW_STATUS": review_values.get("REVIEW_STATUS", self.DEFAULT_REVIEW_STATUS),
                    "FINDINGS_DESCRIPTION": findings_description,
                    "TECH_REVIEW_NOTES": review_values.get("TECH_REVIEW_NOTES", ""),
                    "BUS_REVIEW_NOTES": review_values.get("BUS_REVIEW_NOTES", ""),
                }
            )

        return preview_rows

    def refresh_user_preview(self) -> None:
        # Prevent nested refresh calls from transient Qt events while the table is rebuilding.
        if self._refreshing_user_preview:
            return
        self._configure_user_preview_table()
        self._refreshing_user_preview = True
        self.user_preview_table.blockSignals(True)
        self.user_preview_table.setSortingEnabled(False)

        try:
            self.user_preview_table.setRowCount(0)

            usr02_rows = self._load_preview_rows("USR02")
            combined_rows = self._load_preview_rows("ADR6_USR21")
            preview_rows = self._build_user_preview_rows(usr02_rows, combined_rows)

            if not preview_rows:
                self.user_preview_hint.setText(
                    self.format_ui_rtl_text(
                        "לא זוהו עדיין משתמשים להצגה. יש לטעון קבצי USR02 ו-ADR6 / USER_ADDR."
                    )
                )
                self._update_user_review_progress_summary(0, 0, 0)
                return

            filtered_rows, filter_note = self._filter_user_preview_rows(preview_rows)
            filter_mode = self._get_user_preview_filter_mode()

            if filter_note:
                self.user_preview_hint.setText(
                    self.format_ui_rtl_text(
                        f"{filter_note} הטבלה מציגה כעת את כל {len(preview_rows)} המשתמשים שנטענו לכלי."
                    )
                )
                rows_to_display = preview_rows
            else:
                rows_to_display = filtered_rows
                if filter_mode == "all":
                    self.user_preview_hint.setText(
                        self.format_ui_rtl_text(f"הטבלה מציגה כעת {len(rows_to_display)} משתמשים שנטענו לכלי.")
                    )
                else:
                    self.user_preview_hint.setText(
                        self.format_ui_rtl_text(
                            f"הטבלה מציגה כעת {len(rows_to_display)} משתמשים מתוך {len(preview_rows)} בהתאם לטווח התאריכים שנבחר."
                        )
                    )

            if not rows_to_display:
                self._update_user_review_progress_summary(0, 0, 0)
                return

            rows_to_display = sorted(rows_to_display, key=self._export_sort_key)

            for preview_row in rows_to_display:
                row_index = self.user_preview_table.rowCount()
                self.user_preview_table.insertRow(row_index)
                review_key = self._user_reviewer_state_key(preview_row.get("MANDT", ""), preview_row.get("BNAME", ""))
                for column, field_name in enumerate(self.user_preview_visible_columns):
                    value = preview_row.get(field_name, "") or ""

                    if field_name == "REVIEW_STATUS":
                        display_status = self._normalize_reviewer_status(value)
                        status_item = SortableTableWidgetItem(self.format_rtl_text(display_status))
                        status_item.setData(SortableTableWidgetItem.SORT_ROLE, display_status)
                        status_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        status_item.setToolTip(self.format_rtl_text(display_status))
                        status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        self.user_preview_table.setItem(row_index, column, status_item)
                        continue

                    display_value = self._format_user_preview_value_for_display(field_name, value)
                    item = SortableTableWidgetItem(self.format_rtl_text(display_value))
                    item.setData(SortableTableWidgetItem.SORT_ROLE, self._get_user_preview_sort_value(field_name, value))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    item.setToolTip(self.format_rtl_text(display_value))
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

                    if field_name in {"TECH_REVIEW_NOTES", "BUS_REVIEW_NOTES", "REVIEW_NOTES"}:
                        pass  # read-only for audit tool users

                    self.user_preview_table.setItem(row_index, column, item)

                self._update_review_row_highlight(row_index, preview_row)

            self.user_preview_table.resizeColumnsToContents()
            for column_index, field_name in enumerate(self.user_preview_visible_columns):
                default_width = int(self._get_user_preview_column_definition(field_name).get("width", 120))
                self.user_preview_table.setColumnWidth(
                    column_index,
                    max(self.user_preview_table.columnWidth(column_index), default_width),
                )
        finally:
            self.user_preview_table.blockSignals(False)
            self._refreshing_user_preview = False
            self.user_preview_table.setSortingEnabled(True)
            self._refresh_user_review_progress_summary_from_table()
            self._refresh_audit_summary_table()

    def run_validation(self) -> None:
        file_paths = self._current_file_paths()
        if not file_paths or not self.selected_slot_key:
            QMessageBox.warning(self, "חסר קובץ", "יש לבחור קובץ מתוך אחד ממשבצות הקלט לפני הרצת הבדיקה.")
            self.tabs.setCurrentIndex(0)
            return

        if self.validation_thread is not None:
            QMessageBox.information(self, "בדיקה פעילה", "בדיקה כבר רצה ברקע. יש להמתין לסיום.")
            return

        analysis_tab_index = self.tabs.indexOf(self.analysis_tab)
        if analysis_tab_index >= 0:
            self.tabs.setCurrentIndex(analysis_tab_index)
        self._start_slot_validation_async(self.selected_slot_key, file_paths)

    def _set_validation_running_state(self, is_running: bool, slot_key: str | None = None) -> None:
        self.audit_run_button.setEnabled(not is_running)
        self.audit_export_button.setEnabled(not is_running)
        if is_running:
            slot_text = slot_key or ""
            self.analysis_progress_label.setText(self.format_ui_rtl_text(f"מעבד כעת את המשבצת {slot_text}..."))
            self.analysis_progress_bar.setRange(0, 0)
            self.analysis_progress_container.show()
        else:
            self.analysis_progress_container.hide()

    def _start_slot_validation_async(self, slot_key: str, file_paths: list[str]) -> None:
        input_files_dict: dict[str, list[str | Path]] = {
            slot_key: [str(path) for path in file_paths]
        }
        required_columns = self._required_columns_for_slot(slot_key)
        authorized_users = self._get_authorized_stms_users()

        self.validation_thread = QThread(self)
        self.validation_worker = SlotValidationWorker(
            slot_key=slot_key,
            file_paths=list(file_paths),
            input_files_dict=input_files_dict,
            required_columns=required_columns,
            output_dir=self.config.output_dir,
            authorized_users=authorized_users,
        )
        self.validation_worker.moveToThread(self.validation_thread)

        self.validation_thread.started.connect(self.validation_worker.run)
        self.validation_worker.succeeded.connect(self._on_slot_validation_worker_succeeded)
        self.validation_worker.failed.connect(self._on_slot_validation_worker_failed)
        self.validation_worker.finished.connect(self.validation_thread.quit)
        self.validation_worker.finished.connect(self.validation_worker.deleteLater)
        self.validation_thread.finished.connect(self.validation_thread.deleteLater)
        self.validation_thread.finished.connect(self._on_slot_validation_worker_finished)

        self._set_validation_running_state(True, slot_key)
        self.validation_thread.start()

    @Slot(str, list, object)
    def _on_slot_validation_worker_succeeded(self, slot_key: str, file_paths: list[str], result: object) -> None:
        self._handle_slot_validation_success(slot_key, file_paths, result, show_feedback=True)

    @Slot(str, list, str)
    def _on_slot_validation_worker_failed(self, slot_key: str, file_paths: list[str], error_text: str) -> None:
        self._handle_slot_validation_error(slot_key, file_paths, error_text, show_feedback=True)

    @Slot()
    def _on_slot_validation_worker_finished(self) -> None:
        self.validation_worker = None
        self.validation_thread = None
        self._set_validation_running_state(False)

    def run_domain_validation(self, domain: str) -> None:
        if bool(self.DOMAIN_DEFINITIONS.get(domain, {}).get("in_development", False)):
            QMessageBox.information(
                self,
                "תחום בפיתוח",
                f"תחום '{domain}' נמצא בפיתוח ואינו כולל בדיקות אוטומטיות עדיין.\n\nבדיקות לתחום זה יתווספו בגרסאות הבאות.",
            )
            return

        domain_slots = self._get_domain_slots(domain)
        selected_slots: list[tuple[str, list[str]]] = []
        missing_required: list[str] = []

        for slot_key in domain_slots:
            file_paths = list(self.slot_widgets[slot_key].get("selected_paths", []))
            if file_paths:
                selected_slots.append((slot_key, file_paths))
            elif self.SLOT_DEFINITIONS[slot_key]["required"]:
                missing_required.append(slot_key)

        if not selected_slots:
            QMessageBox.warning(
                self,
                "לא נבחרו קבצים",
                f"לא נבחרו קבצים עבור תחום {domain}. יש לבחור לפחות קובץ אחד לפני הרצת הבדיקה.",
            )
            return

        if missing_required:
            QMessageBox.warning(
                self,
                "חסרים קבצי חובה",
                f"בתחום {domain} חסרים קבצי חובה עבור המשבצות: {', '.join(missing_required)}.\n\nהבדיקה תמשיך עבור הקבצים שנבחרו.",
            )

        processed_slots = 0
        processed_files = 0
        total_rows = 0
        total_invalid_rows = 0
        invalid_slots = 0
        failed_slots: list[str] = []

        # Keep the intake summary hidden while multiple slots are still processing.
        self.summary_group.hide()
        self.results_group.hide()

        for slot_key, file_paths in selected_slots:
            slot_summary = self._run_slot_validation(
                slot_key,
                file_paths,
                show_feedback=False,
                update_summary_ui=False,
            )
            processed_slots += 1
            processed_files += int(slot_summary["file_count"])
            total_rows += int(slot_summary["total_rows"])
            total_invalid_rows += int(slot_summary["invalid_rows"])

            if slot_summary["status"] == "error":
                failed_slots.append(slot_key)
            elif not bool(slot_summary["is_valid"]):
                invalid_slots += 1

        summary_lines = [
            f"בדיקת תחום {domain} הושלמה.",
            f"משבצות שנבדקו: {processed_slots}",
            f"קבצים שנבדקו: {processed_files}",
            f"שורות שנבדקו: {total_rows}",
        ]

        if invalid_slots:
            summary_lines.append(f"משבצות עם ממצאים: {invalid_slots}")
        if failed_slots:
            summary_lines.append(f"משבצות שנכשלו בעיבוד: {', '.join(failed_slots)}")
        summary_lines.append("ניתן לבצע לחיצה כפולה על הרשומה בלוג לצפייה בפירוט.")

        total_valid_rows = max(total_rows - total_invalid_rows, 0)
        self.summary_labels["total"].setText(str(total_rows))
        self.summary_labels["valid"].setText(str(total_valid_rows))
        self.summary_labels["invalid"].setText(str(total_invalid_rows))
        if failed_slots:
            self.summary_labels["status"].setText("הושלם עם שגיאות עיבוד")
        elif total_invalid_rows > 0:
            self.summary_labels["status"].setText("הושלם עם שגיאות קליטה")
        else:
            self.summary_labels["status"].setText("תקין")

        self.issues_table.setRowCount(0)
        self.issues_table.insertRow(0)
        for column, value in enumerate(["-", "-", "הסיכום מתייחס לכל המשבצות שנבדקו בריצה הנוכחית"]):
            item = QTableWidgetItem(self.format_rtl_text(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.issues_table.setItem(0, column, item)
        self.issues_table.resizeColumnsToContents()

        self.summary_group.show()
        self.results_group.show()

        if invalid_slots or failed_slots:
            QMessageBox.warning(self, "בדיקת תחום הושלמה עם ממצאים", "\n".join(summary_lines))
        else:
            QMessageBox.information(self, "בדיקת תחום הושלמה", "\n".join(summary_lines))

    def run_category_validation(self, category: str) -> None:
        selected_slots: list[tuple[str, list[str]]] = []
        missing_required: list[str] = []

        for slot_key in self._get_category_slots(category):
            file_paths = list(self.slot_widgets[slot_key].get("selected_paths", []))
            if file_paths:
                selected_slots.append((slot_key, file_paths))
            elif self.SLOT_DEFINITIONS[slot_key]["required"]:
                missing_required.append(slot_key)

        if not selected_slots:
            QMessageBox.warning(
                self,
                "לא נבחרו קבצים",
                f"לא נבחרו קבצים עבור הקבוצה {category}. יש לבחור לפחות קובץ אחד לפני הרצת הבדיקה.",
            )
            return

        if missing_required:
            QMessageBox.warning(
                self,
                "חסרים קבצי חובה",
                f"בקבוצה {category} חסרים קבצי חובה עבור המשבצות: {', '.join(missing_required)}.\n\nהבדיקה תמשיך עבור הקבצים שנבחרו.",
            )

        processed_slots = 0
        processed_files = 0
        total_rows = 0
        total_invalid_rows = 0
        invalid_slots = 0
        failed_slots: list[str] = []

        # Keep the intake summary hidden while multiple slots are still processing.
        self.summary_group.hide()
        self.results_group.hide()

        for slot_key, file_paths in selected_slots:
            slot_summary = self._run_slot_validation(
                slot_key,
                file_paths,
                show_feedback=False,
                update_summary_ui=False,
            )
            processed_slots += 1
            processed_files += int(slot_summary["file_count"])
            total_rows += int(slot_summary["total_rows"])
            total_invalid_rows += int(slot_summary["invalid_rows"])

            if slot_summary["status"] == "error":
                failed_slots.append(slot_key)
            elif not bool(slot_summary["is_valid"]):
                invalid_slots += 1

        summary_lines = [
            f"בדיקת הקבוצה {category} הושלמה.",
            f"משבצות שנבדקו: {processed_slots}",
            f"קבצים שנבדקו: {processed_files}",
            f"שורות שנבדקו: {total_rows}",
        ]

        if invalid_slots:
            summary_lines.append(f"משבצות עם ממצאים: {invalid_slots}")
        if failed_slots:
            summary_lines.append(f"משבצות שנכשלו בעיבוד: {', '.join(failed_slots)}")
        summary_lines.append("ניתן לבצע לחיצה כפולה על הרשומה בלוג לצפייה בפירוט.")

        total_valid_rows = max(total_rows - total_invalid_rows, 0)
        self.summary_labels["total"].setText(str(total_rows))
        self.summary_labels["valid"].setText(str(total_valid_rows))
        self.summary_labels["invalid"].setText(str(total_invalid_rows))
        if failed_slots:
            self.summary_labels["status"].setText("הושלם עם שגיאות עיבוד")
        elif total_invalid_rows > 0:
            self.summary_labels["status"].setText("הושלם עם שגיאות קליטה")
        else:
            self.summary_labels["status"].setText("תקין")

        self.issues_table.setRowCount(0)
        self.issues_table.insertRow(0)
        for column, value in enumerate(["-", "-", "הסיכום מתייחס לכל המשבצות שנבדקו בריצה הנוכחית"]):
            item = QTableWidgetItem(self.format_rtl_text(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.issues_table.setItem(0, column, item)
        self.issues_table.resizeColumnsToContents()

        self.summary_group.show()
        self.results_group.show()

        if invalid_slots or failed_slots:
            QMessageBox.warning(self, "בדיקת קבוצה הושלמה עם ממצאים", "\n".join(summary_lines))
        else:
            QMessageBox.information(self, "בדיקת קבוצה הושלמה", "\n".join(summary_lines))

    def _run_slot_validation(
        self,
        slot_key: str,
        file_paths: list[str],
        show_feedback: bool = True,
        update_summary_ui: bool = True,
    ) -> dict[str, Any]:
        try:
            result = self._process_slot_validation(slot_key, file_paths)
        except Exception as error:
            return self._handle_slot_validation_error(
                slot_key,
                file_paths,
                str(error),
                show_feedback,
                update_summary_ui,
            )

        return self._handle_slot_validation_success(
            slot_key,
            file_paths,
            result,
            show_feedback,
            update_summary_ui,
        )

    def _process_slot_validation(self, slot_key: str, file_paths: list[str]) -> Any:
        if slot_key == "AGR_1251":
            self.summary_labels["status"].setText("מעבד קובצי הרשאות גדולים במנות...")
            QApplication.processEvents()

        input_files_dict: dict[str, list[str | Path]] = {
            slot_key: [str(path) for path in file_paths]
        }

        return process_file(
            input_files=input_files_dict,
            required_columns=self._required_columns_for_slot(slot_key),
            output_dir=self.config.output_dir,
            source_name_override=slot_key,
            authorized_users=self._get_authorized_stms_users(),
        )

    def _handle_slot_validation_error(
        self,
        slot_key: str,
        file_paths: list[str],
        error_text: str,
        show_feedback: bool,
        update_summary_ui: bool = True,
    ) -> dict[str, Any]:
        if update_summary_ui:
            self.summary_labels["status"].setText(f"שגיאה בעיבוד {slot_key}")
            self.issues_table.setRowCount(0)
            error_message = f"אירעה שגיאה במהלך העיבוד של המשבצת {slot_key}: {error_text}"
            self.issues_table.insertRow(0)
            for column, value in enumerate(["מבנה", "SYSTEM", error_message]):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.issues_table.setItem(0, column, item)
            self.issues_table.resizeColumnsToContents()
        self._append_error_log_entries(slot_key, file_paths, str(error_text))
        if show_feedback:
            QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה במהלך העיבוד של המשבצת {slot_key}:\n{error_text}")
        return {
            "slot_key": slot_key,
            "status": "error",
            "file_count": len(file_paths),
            "total_rows": 0,
            "invalid_rows": 0,
            "is_valid": False,
        }

    def _handle_slot_validation_success(
        self,
        slot_key: str,
        file_paths: list[str],
        result: Any,
        show_feedback: bool,
        update_summary_ui: bool = True,
    ) -> dict[str, Any]:

        intake_issues = [iss for iss in result.issues if self._is_intake_issue(iss)]
        valid_rows, invalid_rows = self._compute_intake_summary(result.summary.total_rows, intake_issues)

        if update_summary_ui:
            self.summary_group.show()
            self.results_group.show()
            self.summary_labels["total"].setText(str(result.summary.total_rows))
            self.summary_labels["valid"].setText(str(valid_rows))
            self.summary_labels["invalid"].setText(str(invalid_rows))

            # Only intake-level issues (structural / missing required) surface in this tab.
            # Audit/analysis findings (e.g. RSPARAM policy) are deferred to the analysis tab.
            status_text = "תקין" if not intake_issues else f"שגיאות קליטה - {slot_key}"
            self.summary_labels["status"].setText(status_text)

            self.issues_table.setRowCount(0)
            if intake_issues:
                for issue in intake_issues:
                    row_index = self.issues_table.rowCount()
                    self.issues_table.insertRow(row_index)
                    values = [
                        str(issue.row_number if issue.row_number > 0 else "מבנה"),
                        self.format_rtl_text(issue.column_name),
                        self.format_rtl_text(issue.message),
                    ]
                    for column, value in enumerate(values):
                        item = QTableWidgetItem(value)
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        self.issues_table.setItem(row_index, column, item)
            else:
                self.issues_table.insertRow(0)
                for column, value in enumerate(["-", "-", "לא נמצאו שגיאות קליטה"]):
                    item = QTableWidgetItem(value)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    self.issues_table.setItem(0, column, item)
            self.issues_table.resizeColumnsToContents()

        audit_issues = [iss for iss in result.issues if not self._is_intake_issue(iss)]
        self._upsert_audit_control_data(
            slot_key=slot_key,
            result=result,
            audit_issues=audit_issues,
            extraction_date=self._get_slot_extraction_date(slot_key),
        )
        self._refresh_audit_summary_table()
        self._upsert_permissions_control_data(
            slot_key=slot_key,
            result=result,
            audit_issues=audit_issues,
            extraction_date=self._get_slot_extraction_date(slot_key),
        )
        self._refresh_permissions_summary_table()

        # Cache AGR_1251 / AGR_USERS rows for cross-join user-management permission check
        detected_profile = str(getattr(result, "detected_profile", "") or "").upper()
        if detected_profile == "AGR_1251":
            self.agr_1251_cached_rows = list(getattr(result, "rows", []))
            self._compute_user_mgmt_permissions()
            self._compute_auth_mgmt_permissions()
            self._compute_rscdok99_permissions()
            self._compute_data_mgmt_permissions()
            self._compute_transport_permissions()
            self._compute_debug_permissions()
            self._compute_job_mgmt_permissions()
        elif detected_profile == "AGR_USERS":
            self.agr_users_cached_rows = list(getattr(result, "rows", []))
            self._compute_user_mgmt_permissions()
            self._compute_auth_mgmt_permissions()
            self._compute_rscdok99_permissions()
            self._compute_data_mgmt_permissions()
            self._compute_transport_permissions()
            self._compute_debug_permissions()
            self._compute_job_mgmt_permissions()

        self._sync_permissions_findings_into_analysis_summary()
        self._refresh_audit_summary_table()

        self._append_run_log_entries(slot_key, file_paths, result)
        if result.report_path is not None:
            self.report_path = result.report_path
        self.report_button.setEnabled(self.report_path is not None)
        file_count = len(result.source_files) if result.source_files else len(file_paths)

        if show_feedback:
            if not intake_issues:
                QMessageBox.information(
                    self,
                    "הבדיקה הושלמה",
                    f"בדיקת המשבצת {slot_key} הסתיימה ללא שגיאות קליטה. נקלטו {file_count} קבצים.",
                )
            else:
                ordered_messages: list[str] = []
                structure_messages = [iss.message for iss in intake_issues if "אינו תואם למבנה" in iss.message]
                other_messages = [iss.message for iss in intake_issues if "אינו תואם למבנה" not in iss.message]
                for message in structure_messages + other_messages:
                    if message not in ordered_messages:
                        ordered_messages.append(message)
                    if len(ordered_messages) == 3:
                        break
                summary_text = "\n".join(f"• {message}" for message in ordered_messages)
                QMessageBox.warning(
                    self,
                    "נמצאו שגיאות קליטה",
                    f"בדיקת המשבצת {slot_key} הסתיימה עם שגיאות קליטה.\n\n{summary_text}\n\nממצאי ביקורת יוצגו בטאב 'ביצוע ניתוח לביקורת'.",
                )

        return {
            "slot_key": slot_key,
            "status": "ok",
            "file_count": file_count,
            "total_rows": result.summary.total_rows,
            "invalid_rows": invalid_rows,
            "is_valid": len(intake_issues) == 0,
        }

    def _append_run_log_entries(self, slot_key: str, file_paths: list[str], result) -> None:
        issues_by_file: dict[str, list] = {Path(path).name: [] for path in file_paths}
        row_counts_by_file = dict(getattr(result, "file_row_counts", {}))
        if not row_counts_by_file:
            for row in getattr(result, "rows", []):
                source_file = Path(str(row.get("__source_file", ""))).name
                if source_file:
                    row_counts_by_file[source_file] = row_counts_by_file.get(source_file, 0) + 1
        display_slot_name = self._get_slot_display_name(slot_key)
        report_group = self._get_slot_category(slot_key)
        extraction_date = self._get_slot_extraction_date(slot_key)

        for issue in result.issues:
            if issue.source_file:
                issue_name = Path(issue.source_file).name
                issues_by_file.setdefault(issue_name, []).append(issue)
            else:
                for issue_list in issues_by_file.values():
                    issue_list.append(issue)

        for path in file_paths:
            file_name = Path(path).name
            file_issues = issues_by_file.get(file_name, [])
            # Status and preview in the intake log are based only on intake-level issues.
            # All issues (including audit findings) are stored for drill-down details.
            intake_file_issues = [iss for iss in file_issues if self._is_intake_issue(iss)]
            status_text = "שגוי" if intake_file_issues else "תקין"
            checked_at = datetime.now()
            row_count = row_counts_by_file.get(file_name, 0)
            record = {
                "slot_key": display_slot_name,
                "report_group": report_group,
                "file_name": file_name,
                "extraction_date": extraction_date,
                "row_count": row_count,
                "status": status_text,
                "error_count": len(intake_file_issues),
                "error_preview": self._build_issue_preview(intake_file_issues),
                "date": checked_at.strftime("%Y-%m-%d"),
                "time": checked_at.strftime("%H:%M:%S"),
                "issues": list(file_issues),
            }
            self.run_log_records.append(record)

            row_index = self.run_log_table.rowCount()
            self.run_log_table.insertRow(row_index)
            values = [
                display_slot_name,
                report_group,
                file_name,
                extraction_date,
                str(row_count),
                status_text,
                str(len(intake_file_issues)),
                str(record["error_preview"]),
                str(record["date"]),
                str(record["time"]),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                item.setToolTip(self.format_rtl_text(value))
                if column == 5:
                    item.setBackground(QColor("#fdecec") if status_text in {"שגוי", "שגיאה"} else QColor("#eaf7ea"))
                self.run_log_table.setItem(row_index, column, item)
        self.run_log_table.resizeColumnsToContents()

    def _append_error_log_entries(self, slot_key: str, file_paths: list[str], error_text: str) -> None:
        checked_at = datetime.now()
        display_slot_name = self._get_slot_display_name(slot_key)
        report_group = self._get_slot_category(slot_key)
        extraction_date = self._get_slot_extraction_date(slot_key)

        for path in file_paths:
            file_name = Path(path).name
            issue = ValidationIssue(
                row_number=0,
                column_name="SYSTEM",
                message=f"אירעה שגיאה במהלך העיבוד: {error_text}",
                source_file=file_name,
            )
            record = {
                "slot_key": display_slot_name,
                "report_group": report_group,
                "file_name": file_name,
                "extraction_date": extraction_date,
                "row_count": 0,
                "status": "שגיאה",
                "error_count": 1,
                "error_preview": issue.message,
                "date": checked_at.strftime("%Y-%m-%d"),
                "time": checked_at.strftime("%H:%M:%S"),
                "issues": [issue],
            }
            self.run_log_records.append(record)

            row_index = self.run_log_table.rowCount()
            self.run_log_table.insertRow(row_index)
            values = [
                display_slot_name,
                report_group,
                file_name,
                extraction_date,
                "0",
                "שגיאה",
                "1",
                issue.message,
                str(record["date"]),
                str(record["time"]),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                item.setToolTip(self.format_rtl_text(value))
                if column == 5:
                    item.setBackground(QColor("#fdecec"))
                self.run_log_table.setItem(row_index, column, item)
        self.run_log_table.resizeColumnsToContents()

    @staticmethod
    def _count_stms_control_records(rows: list[dict[str, Any]]) -> int:
        count = 0
        for row in rows:
            normalized_row = {
                str(key).strip().upper(): value
                for key, value in row.items()
                if not str(key).startswith("__")
            }
            trkorr = ""
            for candidate in get_column_aliases("TRKORR"):
                if candidate in normalized_row and str(normalized_row.get(candidate, "")).strip():
                    trkorr = str(normalized_row.get(candidate, "")).strip()
                    break
            import_user = ""
            for candidate in get_column_aliases("IMPORT_USER"):
                if candidate in normalized_row and str(normalized_row.get(candidate, "")).strip():
                    import_user = str(normalized_row.get(candidate, "")).strip()
                    break
            if trkorr and import_user:
                count += 1
        return count

    @staticmethod
    def _find_row_column_by_alias(row: dict[str, Any], candidate: str) -> str | None:
        normalized_map = {str(column).strip().upper(): str(column) for column in row.keys()}
        for alias in get_column_aliases(candidate):
            if alias in normalized_map:
                return normalized_map[alias]
        return None

    @classmethod
    def _resolve_row_value_by_priority(cls, row: dict[str, Any], candidate: str) -> object | None:
        normalized_map = {str(column).strip().upper(): str(column) for column in row.keys()}
        fallback_value: object | None = None
        for alias in get_column_aliases(candidate):
            if alias not in normalized_map:
                continue
            value = row.get(normalized_map[alias])
            if fallback_value is None:
                fallback_value = value
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            return value
        return fallback_value

    @classmethod
    def _build_password_control_snapshots(cls, rows: list[dict[str, Any]]) -> dict[str, dict[str, str]]:
        param_map: dict[str, object] = {}
        for row in rows:
            parameter_column = cls._find_row_column_by_alias(row, "PARAMETER")
            if not parameter_column:
                continue

            parameter_name = str(row.get(parameter_column, "")).strip().casefold()
            if not parameter_name:
                continue

            value = cls._resolve_row_value_by_priority(row, "VALUE")
            if value is None:
                continue

            param_map[parameter_name] = value

        snapshots: dict[str, dict[str, str]] = {}
        for control_id, param_name, expected, _rule_type, _message in SAP_APP_RSPARAM_RULES:
            if param_name not in param_map:
                continue

            actual = param_map[param_name]
            snapshots[control_id] = {
                "actual_value": str(actual),
                "expected_value": str(expected),
                "status": "תקין",
                "full_description": f"הערך בפועל עבור {param_name} הוא {actual}, בעוד שהערך המצופה הוא {expected}. ההגדרה תקינה לפי דרישת הבקרה.",
            }

        return snapshots

    def _build_strong_profiles_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.permissions_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - משתמשים חזקים"))
        self.permissions_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        permissions_summary_layout = QVBoxLayout(self.permissions_summary_group)
        permissions_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.permissions_summary_table = QTableWidget(0, 6)
        self.permissions_summary_table.setItemDelegate(_RightAlignDelegate(self.permissions_summary_table))
        self.permissions_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("מזהה בקרה"),
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _perm_summary_hdr = self.permissions_summary_table.horizontalHeader()
        _perm_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _perm_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _perm_summary_hdr.setStretchLastSection(False)
        self.permissions_summary_table.setColumnWidth(2, 220)  # ממצא
        self.permissions_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.permissions_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.permissions_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.permissions_summary_table.setAlternatingRowColors(True)
        self.permissions_summary_table.setMinimumHeight(160)
        self.permissions_summary_table.itemSelectionChanged.connect(self._refresh_selected_permissions_users)
        permissions_summary_layout.addWidget(self.permissions_summary_table)
        parent_layout.addWidget(self.permissions_summary_group)

        self.permissions_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים הכלולים בממצא"))
        self.permissions_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        permissions_users_layout = QVBoxLayout(self.permissions_users_group)
        permissions_users_layout.setContentsMargins(8, 14, 8, 8)

        self.permissions_users_table = QTableWidget(0, 2)
        self.permissions_users_table.setItemDelegate(_RightAlignDelegate(self.permissions_users_table))
        self.permissions_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _perm_users_hdr = self.permissions_users_table.horizontalHeader()
        _perm_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _perm_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _perm_users_hdr.setStretchLastSection(True)  # משתמש
        self.permissions_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.permissions_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.permissions_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.permissions_users_table.setAlternatingRowColors(True)
        self.permissions_users_table.setMinimumHeight(180)
        self.permissions_users_table.setToolTip(self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הפרופילים החזקים שלו"))
        self.permissions_users_table.cellDoubleClicked.connect(self.show_permissions_user_profiles_dialog)
        permissions_users_layout.addWidget(self.permissions_users_table)
        parent_layout.addWidget(self.permissions_users_group)

        self._refresh_permissions_summary_table()

    def _build_user_mgmt_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.user_mgmt_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - ניהול משתמשים"))
        self.user_mgmt_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        user_mgmt_summary_layout = QVBoxLayout(self.user_mgmt_summary_group)
        user_mgmt_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.user_mgmt_summary_table = QTableWidget(0, 5)
        self.user_mgmt_summary_table.setItemDelegate(_RightAlignDelegate(self.user_mgmt_summary_table))
        self.user_mgmt_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _usrmgmt_summary_hdr = self.user_mgmt_summary_table.horizontalHeader()
        _usrmgmt_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _usrmgmt_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _usrmgmt_summary_hdr.setStretchLastSection(False)
        self.user_mgmt_summary_table.setColumnWidth(1, 280)  # ממצא
        self.user_mgmt_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.user_mgmt_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.user_mgmt_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.user_mgmt_summary_table.setAlternatingRowColors(True)
        self.user_mgmt_summary_table.setMinimumHeight(160)
        self.user_mgmt_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאת ניהול משתמשים")
        )
        self.user_mgmt_summary_table.itemSelectionChanged.connect(self._refresh_selected_user_mgmt_users)
        user_mgmt_summary_layout.addWidget(self.user_mgmt_summary_table)
        parent_layout.addWidget(self.user_mgmt_summary_group)

        self.user_mgmt_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאת ניהול משתמשים"))
        self.user_mgmt_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        user_mgmt_users_layout = QVBoxLayout(self.user_mgmt_users_group)
        user_mgmt_users_layout.setContentsMargins(8, 14, 8, 8)

        self.user_mgmt_users_table = QTableWidget(0, 2)
        self.user_mgmt_users_table.setItemDelegate(_RightAlignDelegate(self.user_mgmt_users_table))
        self.user_mgmt_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _usrmgmt_users_hdr = self.user_mgmt_users_table.horizontalHeader()
        _usrmgmt_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _usrmgmt_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _usrmgmt_users_hdr.setStretchLastSection(True)
        self.user_mgmt_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.user_mgmt_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.user_mgmt_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.user_mgmt_users_table.setAlternatingRowColors(True)
        self.user_mgmt_users_table.setMinimumHeight(180)
        self.user_mgmt_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאת ניהול משתמשים")
        )
        self.user_mgmt_users_table.cellDoubleClicked.connect(self.show_user_mgmt_user_dialog)
        user_mgmt_users_layout.addWidget(self.user_mgmt_users_table)
        parent_layout.addWidget(self.user_mgmt_users_group)

        self._refresh_user_mgmt_summary_table()

    def _upsert_permissions_control_data(
        self,
        slot_key: str,
        result: Any,
        audit_issues: list[ValidationIssue],
        extraction_date: str,
    ) -> None:
        del slot_key
        del extraction_date
        detected_profile = str(getattr(result, "detected_profile", "") or "").upper()
        control_id = "MA-PERM-01"
        if detected_profile != "UST04":
            return
        if control_id not in get_profile_audit_controls(detected_profile):
            return

        strong_issues = [issue for issue in audit_issues if issue.control_id == control_id]
        users_by_client: dict[str, dict[str, set[str]]] = {}
        rows = list(getattr(result, "rows", []))

        for issue in strong_issues:
            client_name = "-"
            user_name = str(issue.actual_value or "").strip().upper()
            profile_name = str(issue.expected_value or "").strip().upper()

            if issue.row_number > 0 and issue.row_number <= len(rows):
                row = rows[issue.row_number - 1]
                row_client = self._resolve_row_value_by_priority(row, "MANDT")
                if row_client is not None and str(row_client).strip():
                    client_name = str(row_client).strip()
                if not user_name:
                    row_user = self._resolve_row_value_by_priority(row, "BNAME")
                    if row_user is not None:
                        user_name = str(row_user).strip().upper()
                if not profile_name:
                    row_profile = self._resolve_row_value_by_priority(row, "PROFILE")
                    if row_profile is not None:
                        profile_name = str(row_profile).strip().upper()

            if not user_name:
                continue

            client_users = users_by_client.setdefault(client_name, {})
            client_users.setdefault(user_name, set())
            if profile_name:
                client_users[user_name].add(profile_name)

        keys_to_delete = [key for key in self.permissions_summary_records if str(key).startswith(f"{control_id}|")]
        for key in keys_to_delete:
            self.permissions_summary_records.pop(key, None)
            self.permissions_users_by_control.pop(key, None)

        control_meta = get_audit_control_definition(control_id)
        if not users_by_client:
            record_key = f"{control_id}|-"
            self.permissions_summary_records[record_key] = {
                "record_key": record_key,
                "control_id": control_id,
                "client": "-",
                "finding_text": "נמצאו 0 משתמשים בעלי פרופילים חזקים",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.permissions_users_by_control[record_key] = []
            return

        for client_name, client_users in sorted(users_by_client.items(), key=lambda item: item[0]):
            users_count = len(client_users)
            record_key = f"{control_id}|{client_name}"
            self.permissions_summary_records[record_key] = {
                "record_key": record_key,
                "control_id": control_id,
                "client": client_name,
                "finding_text": f"נמצאו {users_count} משתמשים בעלי פרופילים חזקים",
                "users_count": users_count,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "עם ממצא" if users_count > 0 else "תקין",
            }
            self.permissions_users_by_control[record_key] = [
                {
                    "client": client_name,
                    "user_name": user_name,
                    "profiles": sorted(profiles),
                }
                for user_name, profiles in sorted(client_users.items())
            ]

    @staticmethod
    def _extract_control_id_from_record_key(record_key: object, fallback_control_id: str) -> str:
        key_text = str(record_key or "").strip()
        if "|" in key_text:
            parsed_control_id = key_text.split("|", 1)[0].strip()
            if parsed_control_id:
                return parsed_control_id
        return fallback_control_id

    @staticmethod
    def _to_int(value: object) -> int:
        try:
            return int(str(value).strip())
        except (TypeError, ValueError):
            return 0

    def _permission_control_slots(self, control_id: str) -> list[str]:
        control_slots: dict[str, list[str]] = {
            "MA-PERM-01": ["UST04"],
            "MA-USRMGMT-01": ["AGR_1251", "AGR_USERS"],
            "MA-AUTHMGMT-01": ["AGR_1251", "AGR_USERS"],
            "MA-RSCDOK99-01": ["AGR_1251", "AGR_USERS"],
            "MA-DATAMGMT-01": ["AGR_1251", "AGR_USERS"],
            "MA-TRANSPORT-01": ["AGR_1251", "AGR_USERS"],
            "MA-DEBUG-01": ["AGR_1251", "AGR_USERS"],
            "MA-JOBMGMT-01": ["AGR_1251", "AGR_USERS"],
        }
        return control_slots.get(control_id, ["AGR_1251", "AGR_USERS"])

    def _permission_source_file_label(self, control_id: str) -> str:
        labels: list[str] = []
        for slot_key in self._permission_control_slots(control_id):
            label = self._get_slot_display_name(slot_key)
            if label not in labels:
                labels.append(label)
        return ", ".join(labels) if labels else "-"

    def _permission_extraction_date_label(self, control_id: str) -> str:
        extraction_dates: list[str] = []
        for slot_key in self._permission_control_slots(control_id):
            extraction_date = self._get_slot_extraction_date(slot_key)
            if extraction_date and extraction_date not in extraction_dates:
                extraction_dates.append(extraction_date)
        return ", ".join(extraction_dates) if extraction_dates else "-"

    def _permission_summary_sources(self) -> list[tuple[str, dict[str, dict[str, Any]]]]:
        return [
            ("MA-PERM-01", self.permissions_summary_records),
            ("MA-USRMGMT-01", self.user_mgmt_summary_records),
            ("MA-AUTHMGMT-01", self.auth_mgmt_summary_records),
            ("MA-RSCDOK99-01", self.rscdok99_summary_records),
            ("MA-DATAMGMT-01", self.data_mgmt_summary_records),
            ("MA-TRANSPORT-01", self.transport_summary_records),
            ("MA-DEBUG-01", self.debug_summary_records),
            ("MA-JOBMGMT-01", self.job_mgmt_summary_records),
        ]

    def _permission_user_sources(self) -> dict[str, dict[str, list[dict[str, Any]]]]:
        return {
            "MA-PERM-01": self.permissions_users_by_control,
            "MA-USRMGMT-01": self.user_mgmt_users_by_control,
            "MA-AUTHMGMT-01": self.auth_mgmt_users_by_control,
            "MA-RSCDOK99-01": self.rscdok99_users_by_control,
            "MA-DATAMGMT-01": self.data_mgmt_users_by_control,
            "MA-TRANSPORT-01": self.transport_users_by_control,
            "MA-DEBUG-01": self.debug_users_by_control,
            "MA-JOBMGMT-01": self.job_mgmt_users_by_control,
        }

    def _sync_permissions_findings_into_analysis_summary(self) -> None:
        permission_user_sources = self._permission_user_sources()
        permission_control_ids = [
            control_id
            for control_id, _records_map in self._permission_summary_sources()
        ]

        for control_id in permission_control_ids:
            self.audit_summary_records.pop(control_id, None)
            self.audit_details_by_control.pop(control_id, None)

        for fallback_control_id, records_map in self._permission_summary_sources():
            grouped_records: dict[str, list[dict[str, Any]]] = {}
            for row_data in records_map.values():
                control_id = str(row_data.get("control_id", "")).strip()
                if not control_id:
                    control_id = self._extract_control_id_from_record_key(
                        row_data.get("record_key", ""),
                        fallback_control_id,
                    )
                grouped_records.setdefault(control_id, []).append(row_data)

            for control_id, rows in grouped_records.items():
                if not rows:
                    continue

                control_meta = get_audit_control_definition(control_id)
                sorted_rows = sorted(rows, key=lambda item: str(item.get("client", "-")))
                user_map = permission_user_sources.get(control_id, {})
                all_users: set[tuple[str, str]] = set()
                finding_users: set[tuple[str, str]] = set()

                for row in sorted_rows:
                    record_key = str(row.get("record_key", "") or "")
                    users = user_map.get(record_key, [])
                    row_has_finding = (
                        self._to_int(row.get("users_count", 0)) > 0
                        or str(row.get("status", "")).strip() == "עם ממצא"
                    )
                    for user_data in users:
                        client_name = str(user_data.get("client", "-") or "-")
                        user_name = str(user_data.get("user_name", "-") or "-").strip()
                        if not user_name or user_name == "-":
                            continue
                        user_key = (client_name, user_name.upper())
                        all_users.add(user_key)
                        if row_has_finding:
                            finding_users.add(user_key)

                if all_users:
                    total_records = len(all_users)
                    finding_records = len(finding_users)
                else:
                    total_records = sum(max(self._to_int(row.get("users_count", 0)), 0) for row in sorted_rows)
                    finding_records = sum(
                        max(self._to_int(row.get("users_count", 0)), 0)
                        for row in sorted_rows
                        if str(row.get("status", "")).strip() == "עם ממצא"
                        or self._to_int(row.get("users_count", 0)) > 0
                    )
                valid_records = max(total_records - finding_records, 0)

                self.audit_summary_records[control_id] = {
                    "control_id": control_id,
                    "check_type": control_meta.get("check_type", "סקירת הרשאות"),
                    "source_file": self._permission_source_file_label(control_id),
                    "extraction_date": self._permission_extraction_date_label(control_id),
                    "work_environment": self._current_work_environment_label(),
                    "risk_level": control_meta.get("risk_level", "-"),
                    "description": control_meta.get("description", "-"),
                    "valid_records": valid_records,
                    "finding_records": finding_records,
                    "total_records": total_records,
                }
                source_file_label = self._permission_source_file_label(control_id)
                extraction_date_label = self._permission_extraction_date_label(control_id)
                detail_rows: list[dict[str, Any]] = []

                for row in sorted_rows:
                    record_key = str(row.get("record_key", "") or "")
                    users = user_map.get(record_key, [])
                    sorted_users = sorted(
                        users,
                        key=lambda item: (
                            str(item.get("client", "-")),
                            str(item.get("user_name", "-")),
                        ),
                    )
                    for user_data in sorted_users:
                        client_name = str(user_data.get("client", "-") or "-")
                        user_name = str(user_data.get("user_name", "-") or "-")
                        detail_rows.append(
                            {
                                "control_id": control_id,
                                "source_file": source_file_label,
                                "extraction_date": extraction_date_label,
                                "work_environment": self._current_work_environment_label(),
                                "category": control_meta.get("category", "-"),
                                "risk_level": row.get("risk_level", control_meta.get("risk_level", "-")),
                                "description": control_meta.get("description", "-"),
                                "check_type": control_meta.get("check_type", "סקירת הרשאות"),
                                "actual_value": user_name,
                                "expected_value": "לא נמצאו משתמשים",
                                "status": "עם ממצא",
                                "full_description": (
                                    f"משתמש: {user_name}. קליינט: {client_name}. "
                                    f"{row.get('finding_text', '-')}."
                                ),
                            }
                        )

                if not detail_rows:
                    detail_rows = [
                        {
                            "control_id": control_id,
                            "source_file": source_file_label,
                            "extraction_date": extraction_date_label,
                            "work_environment": self._current_work_environment_label(),
                            "category": control_meta.get("category", "-"),
                            "risk_level": control_meta.get("risk_level", "-"),
                            "description": control_meta.get("description", "-"),
                            "check_type": control_meta.get("check_type", "סקירת הרשאות"),
                            "actual_value": "-",
                            "expected_value": "לא נמצאו משתמשים",
                            "status": "תקין",
                            "full_description": "לא נמצאו משתמשים עם ממצא עבור בקרה זו.",
                        }
                    ]

                self.audit_details_by_control[control_id] = detail_rows

    def _refresh_permissions_summary_table(self) -> None:
        self.permissions_summary_table.setRowCount(0)
        self.permissions_users_table.setRowCount(0)
        if not self.permissions_summary_records:
            self.permissions_users_table.insertRow(0)
            item = QTableWidgetItem(self.format_rtl_text("אין משתמשים להצגה"))
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.permissions_users_table.setItem(0, 1, item)
            return

        for row_data in sorted(
            self.permissions_summary_records.values(),
            key=lambda item: (str(item.get("control_id", "")), str(item.get("client", ""))),
        ):
            row_index = self.permissions_summary_table.rowCount()
            self.permissions_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("control_id", "-")),
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.permissions_summary_table.setItem(row_index, column, item)

        if self.permissions_summary_table.rowCount() > 0:
            self.permissions_summary_table.selectRow(0)
            self._refresh_selected_permissions_users()
        self.permissions_summary_table.resizeColumnsToContents()

    def _refresh_selected_permissions_users(self) -> None:
        selected_items = self.permissions_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.permissions_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.permissions_users_by_control.get(record_key, [])
        self.permissions_users_table.setRowCount(0)

        if not user_rows:
            self.permissions_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.permissions_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.permissions_users_table.rowCount()
            self.permissions_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.permissions_users_table.setItem(row_index, 0, client_item)
            self.permissions_users_table.setItem(row_index, 1, user_item)
        self.permissions_users_table.resizeColumnsToContents()

    def show_permissions_user_profiles_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.permissions_users_table.rowCount():
            return

        user_item = self.permissions_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.permissions_summary_table.selectedItems()
        if not selected_items:
            return

        control_item = self.permissions_summary_table.item(selected_items[0].row(), 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.permissions_users_by_control.get(record_key, [])
        profiles: list[str] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                profiles = [str(profile) for profile in user_data.get("profiles", []) if str(profile).strip()]
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט פרופילים חזקים למשתמש")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 360)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [f"קליינט: {client_name}", f"משתמש: {user_name}", "", "פרופילים חזקים:"]
        if profiles:
            lines.extend(f"- {profile}" for profile in profiles)
        else:
            lines.append("- לא נמצאו פרופילים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # User-Management Permissions (MA-USRMGMT-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # ------------------------------------------------------------------

    def _compute_user_mgmt_permissions(self) -> None:
        """Recompute user-management permission findings from cached AGR_1251 + AGR_USERS rows."""
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-USRMGMT-01"
        control_meta = get_audit_control_definition(control_id)

        # Step A: find AGR_NAMEs that carry user-management permission objects.
        # qualifying_values: (OBJECT_upper, FIELD_upper) -> set of qualifying LOW/HIGH values
        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in USER_MGMT_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            # wildcard or value in qualifying set
            if low_str == "*" or high_str == "*":
                qualifies = True
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        # Step B: scan AGR_USERS; for each row whose AGR_NAME is in matching set,
        #         group by MANDT → UNAME → set of AGR_NAMEs.
        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            # Resolve MANDT with fallback to filename
            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        # Step C: rebuild summary/users dicts
        self.user_mgmt_summary_records.clear()
        self.user_mgmt_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.user_mgmt_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאות ניהול משתמשים",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.user_mgmt_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.user_mgmt_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאות ניהול משתמשים",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.user_mgmt_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_user_mgmt_summary_table()

    def _refresh_user_mgmt_summary_table(self) -> None:
        self.user_mgmt_summary_table.setRowCount(0)
        self.user_mgmt_users_table.setRowCount(0)
        if not self.user_mgmt_summary_records:
            self.user_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.user_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.user_mgmt_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.user_mgmt_summary_table.rowCount()
            self.user_mgmt_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.user_mgmt_summary_table.setItem(row_index, column, item)

        if self.user_mgmt_summary_table.rowCount() > 0:
            self.user_mgmt_summary_table.selectRow(0)
            self._refresh_selected_user_mgmt_users()
        self.user_mgmt_summary_table.resizeColumnsToContents()

    def _refresh_selected_user_mgmt_users(self) -> None:
        selected_items = self.user_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.user_mgmt_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.user_mgmt_users_by_control.get(record_key, [])
        self.user_mgmt_users_table.setRowCount(0)

        if not user_rows:
            self.user_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.user_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.user_mgmt_users_table.rowCount()
            self.user_mgmt_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.user_mgmt_users_table.setItem(row_index, 0, client_item)
            self.user_mgmt_users_table.setItem(row_index, 1, user_item)
        self.user_mgmt_users_table.resizeColumnsToContents()

    def show_user_mgmt_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.user_mgmt_users_table.rowCount():
            return

        user_item = self.user_mgmt_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.user_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.user_mgmt_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.user_mgmt_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאות ניהול משתמשים")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # Authorization-Management Permissions (MA-AUTHMGMT-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # ------------------------------------------------------------------

    def _build_auth_mgmt_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.auth_mgmt_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - ניהול הרשאות"))
        self.auth_mgmt_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        auth_mgmt_summary_layout = QVBoxLayout(self.auth_mgmt_summary_group)
        auth_mgmt_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.auth_mgmt_summary_table = QTableWidget(0, 5)
        self.auth_mgmt_summary_table.setItemDelegate(_RightAlignDelegate(self.auth_mgmt_summary_table))
        self.auth_mgmt_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _authmgmt_summary_hdr = self.auth_mgmt_summary_table.horizontalHeader()
        _authmgmt_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _authmgmt_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _authmgmt_summary_hdr.setStretchLastSection(False)
        self.auth_mgmt_summary_table.setColumnWidth(1, 280)  # ממצא
        self.auth_mgmt_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.auth_mgmt_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.auth_mgmt_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.auth_mgmt_summary_table.setAlternatingRowColors(True)
        self.auth_mgmt_summary_table.setMinimumHeight(160)
        self.auth_mgmt_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאת ניהול הרשאות")
        )
        self.auth_mgmt_summary_table.itemSelectionChanged.connect(self._refresh_selected_auth_mgmt_users)
        auth_mgmt_summary_layout.addWidget(self.auth_mgmt_summary_table)
        parent_layout.addWidget(self.auth_mgmt_summary_group)

        self.auth_mgmt_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאת ניהול הרשאות"))
        self.auth_mgmt_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        auth_mgmt_users_layout = QVBoxLayout(self.auth_mgmt_users_group)
        auth_mgmt_users_layout.setContentsMargins(8, 14, 8, 8)

        self.auth_mgmt_users_table = QTableWidget(0, 2)
        self.auth_mgmt_users_table.setItemDelegate(_RightAlignDelegate(self.auth_mgmt_users_table))
        self.auth_mgmt_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _authmgmt_users_hdr = self.auth_mgmt_users_table.horizontalHeader()
        _authmgmt_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _authmgmt_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _authmgmt_users_hdr.setStretchLastSection(True)
        self.auth_mgmt_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.auth_mgmt_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.auth_mgmt_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.auth_mgmt_users_table.setAlternatingRowColors(True)
        self.auth_mgmt_users_table.setMinimumHeight(180)
        self.auth_mgmt_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאת ניהול הרשאות")
        )
        self.auth_mgmt_users_table.cellDoubleClicked.connect(self.show_auth_mgmt_user_dialog)
        auth_mgmt_users_layout.addWidget(self.auth_mgmt_users_table)
        parent_layout.addWidget(self.auth_mgmt_users_group)

        self._refresh_auth_mgmt_summary_table()

    def _compute_auth_mgmt_permissions(self) -> None:
        """Recompute authorization-management permission findings from cached AGR_1251 + AGR_USERS rows."""
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-AUTHMGMT-01"
        control_meta = get_audit_control_definition(control_id)

        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in AUTH_MGMT_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            if low_str == "*" or high_str == "*":
                qualifies = True
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        self.auth_mgmt_summary_records.clear()
        self.auth_mgmt_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.auth_mgmt_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאות ניהול הרשאות",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.auth_mgmt_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.auth_mgmt_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאות ניהול הרשאות",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.auth_mgmt_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_auth_mgmt_summary_table()

    def _refresh_auth_mgmt_summary_table(self) -> None:
        self.auth_mgmt_summary_table.setRowCount(0)
        self.auth_mgmt_users_table.setRowCount(0)
        if not self.auth_mgmt_summary_records:
            self.auth_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.auth_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.auth_mgmt_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.auth_mgmt_summary_table.rowCount()
            self.auth_mgmt_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.auth_mgmt_summary_table.setItem(row_index, column, item)

        if self.auth_mgmt_summary_table.rowCount() > 0:
            self.auth_mgmt_summary_table.selectRow(0)
            self._refresh_selected_auth_mgmt_users()
        self.auth_mgmt_summary_table.resizeColumnsToContents()

    def _refresh_selected_auth_mgmt_users(self) -> None:
        selected_items = self.auth_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.auth_mgmt_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.auth_mgmt_users_by_control.get(record_key, [])
        self.auth_mgmt_users_table.setRowCount(0)

        if not user_rows:
            self.auth_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.auth_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.auth_mgmt_users_table.rowCount()
            self.auth_mgmt_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.auth_mgmt_users_table.setItem(row_index, 0, client_item)
            self.auth_mgmt_users_table.setItem(row_index, 1, user_item)
        self.auth_mgmt_users_table.resizeColumnsToContents()

    def show_auth_mgmt_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.auth_mgmt_users_table.rowCount():
            return

        user_item = self.auth_mgmt_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.auth_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.auth_mgmt_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.auth_mgmt_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאות ניהול הרשאות")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # RSCDOK99 Program Permissions (MA-RSCDOK99-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # AND logic: an AGR_NAME qualifies only when it satisfies ALL criteria
    # (S_PROGRAM/P_GROUP=RSCDOK99 AND S_PROGRAM/P_ACTION=SUB).
    # ------------------------------------------------------------------

    def _build_rscdok99_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.rscdok99_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - תוכנית RSCDOK99"))
        self.rscdok99_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        rscdok99_summary_layout = QVBoxLayout(self.rscdok99_summary_group)
        rscdok99_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.rscdok99_summary_table = QTableWidget(0, 5)
        self.rscdok99_summary_table.setItemDelegate(_RightAlignDelegate(self.rscdok99_summary_table))
        self.rscdok99_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _rscdok99_summary_hdr = self.rscdok99_summary_table.horizontalHeader()
        _rscdok99_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _rscdok99_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _rscdok99_summary_hdr.setStretchLastSection(False)
        self.rscdok99_summary_table.setColumnWidth(1, 300)  # ממצא
        self.rscdok99_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.rscdok99_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.rscdok99_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.rscdok99_summary_table.setAlternatingRowColors(True)
        self.rscdok99_summary_table.setMinimumHeight(160)
        self.rscdok99_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאה לתוכנית RSCDOK99")
        )
        self.rscdok99_summary_table.itemSelectionChanged.connect(self._refresh_selected_rscdok99_users)
        rscdok99_summary_layout.addWidget(self.rscdok99_summary_table)
        parent_layout.addWidget(self.rscdok99_summary_group)

        self.rscdok99_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאה לתוכנית RSCDOK99"))
        self.rscdok99_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        rscdok99_users_layout = QVBoxLayout(self.rscdok99_users_group)
        rscdok99_users_layout.setContentsMargins(8, 14, 8, 8)

        self.rscdok99_users_table = QTableWidget(0, 2)
        self.rscdok99_users_table.setItemDelegate(_RightAlignDelegate(self.rscdok99_users_table))
        self.rscdok99_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _rscdok99_users_hdr = self.rscdok99_users_table.horizontalHeader()
        _rscdok99_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _rscdok99_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _rscdok99_users_hdr.setStretchLastSection(True)
        self.rscdok99_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.rscdok99_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.rscdok99_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.rscdok99_users_table.setAlternatingRowColors(True)
        self.rscdok99_users_table.setMinimumHeight(180)
        self.rscdok99_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאה לתוכנית RSCDOK99")
        )
        self.rscdok99_users_table.cellDoubleClicked.connect(self.show_rscdok99_user_dialog)
        rscdok99_users_layout.addWidget(self.rscdok99_users_table)
        parent_layout.addWidget(self.rscdok99_users_group)

        self._refresh_rscdok99_summary_table()

    def _compute_rscdok99_permissions(self) -> None:
        """Recompute RSCDOK99 permission findings from cached AGR_1251 + AGR_USERS rows.

        An AGR_NAME qualifies only when ALL criteria in RSCDOK99_PERMISSION_CRITERIA
        are satisfied (AND logic): both S_PROGRAM/P_GROUP=RSCDOK99 and
        S_PROGRAM/P_ACTION=SUB must appear in the role's permission rows.
        """
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-RSCDOK99-01"
        control_meta = get_audit_control_definition(control_id)

        # Build per-AGR_NAME satisfied-criteria tracking.
        # criteria_count[agr_name] = set of criterion indexes satisfied
        # agr_name_objects[agr_name] = set of (OBJECT, FIELD, LOW_display) tuples that qualified
        criteria_count: dict[str, set[int]] = {}
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        total_criteria = len(RSCDOK99_PERMISSION_CRITERIA)

        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""

            for idx, (crit_obj, crit_fld, crit_values) in enumerate(RSCDOK99_PERMISSION_CRITERIA):
                if obj_upper != crit_obj.upper() or fld_upper != crit_fld.upper():
                    continue
                crit_upper = {v.upper() for v in crit_values}
                if low_str == "*" or high_str == "*":
                    qualifies = True
                else:
                    qualifies = bool(low_str and low_str in crit_upper) or bool(
                        high_str and high_str in crit_upper
                    )
                if not qualifies:
                    continue
                agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
                if agr_name_val is None or not str(agr_name_val).strip():
                    continue
                agr_name_upper = str(agr_name_val).strip().upper()
                criteria_count.setdefault(agr_name_upper, set()).add(idx)
                low_display = low_str if low_str else "-"
                agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        # Only keep roles that satisfy ALL criteria
        matching_agr_names: set[str] = {
            agr for agr, satisfied in criteria_count.items()
            if len(satisfied) >= total_criteria
        }

        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        self.rscdok99_summary_records.clear()
        self.rscdok99_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.rscdok99_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאה לתוכנית RSCDOK99",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.rscdok99_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.rscdok99_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאה לתוכנית RSCDOK99",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.rscdok99_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_rscdok99_summary_table()

    def _refresh_rscdok99_summary_table(self) -> None:
        self.rscdok99_summary_table.setRowCount(0)
        self.rscdok99_users_table.setRowCount(0)
        if not self.rscdok99_summary_records:
            self.rscdok99_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.rscdok99_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.rscdok99_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.rscdok99_summary_table.rowCount()
            self.rscdok99_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.rscdok99_summary_table.setItem(row_index, column, item)

        if self.rscdok99_summary_table.rowCount() > 0:
            self.rscdok99_summary_table.selectRow(0)
            self._refresh_selected_rscdok99_users()
        self.rscdok99_summary_table.resizeColumnsToContents()

    def _refresh_selected_rscdok99_users(self) -> None:
        selected_items = self.rscdok99_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.rscdok99_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.rscdok99_users_by_control.get(record_key, [])
        self.rscdok99_users_table.setRowCount(0)

        if not user_rows:
            self.rscdok99_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.rscdok99_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.rscdok99_users_table.rowCount()
            self.rscdok99_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.rscdok99_users_table.setItem(row_index, 0, client_item)
            self.rscdok99_users_table.setItem(row_index, 1, user_item)
        self.rscdok99_users_table.resizeColumnsToContents()

    def show_rscdok99_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.rscdok99_users_table.rowCount():
            return

        user_item = self.rscdok99_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.rscdok99_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.rscdok99_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.rscdok99_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאה לתוכנית RSCDOK99")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # Data Management Permissions (MA-DATAMGMT-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # OR logic: any single matching AGR_1251 row qualifies the AGR_NAME.
    # Qualifying objects: S_TCODE/TCD (SE16/SM30/SM31/SE16N/SE17/SM38/SE37),
    #   S_TABU_DIS/ACTVT (01/02/06), S_TABU_NAM/ACTVT (01/02/06),
    #   S_TABU_NAM/TABLE (any value), S_TABU_CLI/CLIIDMAINT (X),
    #   S_DATASET/ACTVT (06/34).
    # ------------------------------------------------------------------

    def _build_data_mgmt_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.data_mgmt_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - ניהול נתונים"))
        self.data_mgmt_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        data_mgmt_summary_layout = QVBoxLayout(self.data_mgmt_summary_group)
        data_mgmt_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.data_mgmt_summary_table = QTableWidget(0, 5)
        self.data_mgmt_summary_table.setItemDelegate(_RightAlignDelegate(self.data_mgmt_summary_table))
        self.data_mgmt_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _datamgmt_summary_hdr = self.data_mgmt_summary_table.horizontalHeader()
        _datamgmt_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _datamgmt_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _datamgmt_summary_hdr.setStretchLastSection(False)
        self.data_mgmt_summary_table.setColumnWidth(1, 300)  # ממצא
        self.data_mgmt_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.data_mgmt_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.data_mgmt_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.data_mgmt_summary_table.setAlternatingRowColors(True)
        self.data_mgmt_summary_table.setMinimumHeight(160)
        self.data_mgmt_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאת ניהול נתונים")
        )
        self.data_mgmt_summary_table.itemSelectionChanged.connect(self._refresh_selected_data_mgmt_users)
        data_mgmt_summary_layout.addWidget(self.data_mgmt_summary_table)
        parent_layout.addWidget(self.data_mgmt_summary_group)

        self.data_mgmt_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאת ניהול נתונים"))
        self.data_mgmt_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        data_mgmt_users_layout = QVBoxLayout(self.data_mgmt_users_group)
        data_mgmt_users_layout.setContentsMargins(8, 14, 8, 8)

        self.data_mgmt_users_table = QTableWidget(0, 2)
        self.data_mgmt_users_table.setItemDelegate(_RightAlignDelegate(self.data_mgmt_users_table))
        self.data_mgmt_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _datamgmt_users_hdr = self.data_mgmt_users_table.horizontalHeader()
        _datamgmt_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _datamgmt_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _datamgmt_users_hdr.setStretchLastSection(True)
        self.data_mgmt_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.data_mgmt_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.data_mgmt_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.data_mgmt_users_table.setAlternatingRowColors(True)
        self.data_mgmt_users_table.setMinimumHeight(180)
        self.data_mgmt_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאת ניהול נתונים")
        )
        self.data_mgmt_users_table.cellDoubleClicked.connect(self.show_data_mgmt_user_dialog)
        data_mgmt_users_layout.addWidget(self.data_mgmt_users_table)
        parent_layout.addWidget(self.data_mgmt_users_group)

        self._refresh_data_mgmt_summary_table()

    def _compute_data_mgmt_permissions(self) -> None:
        """Recompute data-management permission findings from cached AGR_1251 + AGR_USERS rows.

        OR logic: any single qualifying row in AGR_1251 makes the AGR_NAME a match.
        Special wildcard criterion: if the qualifying set for a (OBJECT, FIELD) pair
        contains "*", then any non-empty LOW or HIGH value qualifies.
        """
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-DATAMGMT-01"
        control_meta = get_audit_control_definition(control_id)

        # Step A: find AGR_NAMEs that carry data-management permission objects.
        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in DATA_MGMT_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            # SAP data has a wildcard → always qualifies
            if low_str == "*" or high_str == "*":
                qualifies = True
            elif "*" in qualifying_map[key]:
                # Criterion allows any value → any non-empty LOW or HIGH qualifies
                qualifies = bool(low_str) or bool(high_str)
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        # Step B: cross-join with AGR_USERS to collect users per client.
        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        # Step C+D: store results.
        self.data_mgmt_summary_records.clear()
        self.data_mgmt_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.data_mgmt_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאות ניהול נתונים",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.data_mgmt_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.data_mgmt_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאות ניהול נתונים",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.data_mgmt_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_data_mgmt_summary_table()

    def _refresh_data_mgmt_summary_table(self) -> None:
        self.data_mgmt_summary_table.setRowCount(0)
        self.data_mgmt_users_table.setRowCount(0)
        if not self.data_mgmt_summary_records:
            self.data_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.data_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.data_mgmt_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.data_mgmt_summary_table.rowCount()
            self.data_mgmt_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.data_mgmt_summary_table.setItem(row_index, column, item)

        if self.data_mgmt_summary_table.rowCount() > 0:
            self.data_mgmt_summary_table.selectRow(0)
            self._refresh_selected_data_mgmt_users()
        self.data_mgmt_summary_table.resizeColumnsToContents()

    def _refresh_selected_data_mgmt_users(self) -> None:
        selected_items = self.data_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.data_mgmt_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.data_mgmt_users_by_control.get(record_key, [])
        self.data_mgmt_users_table.setRowCount(0)

        if not user_rows:
            self.data_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.data_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.data_mgmt_users_table.rowCount()
            self.data_mgmt_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.data_mgmt_users_table.setItem(row_index, 0, client_item)
            self.data_mgmt_users_table.setItem(row_index, 1, user_item)
        self.data_mgmt_users_table.resizeColumnsToContents()

    def show_data_mgmt_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.data_mgmt_users_table.rowCount():
            return

        user_item = self.data_mgmt_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.data_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.data_mgmt_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.data_mgmt_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאות ניהול נתונים")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # Transport / Change-Management Permissions (MA-TRANSPORT-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # OR logic: any single qualifying AGR_1251 row makes the AGR_NAME a match.
    # Qualifying objects: S_TCODE/TCD (STMS/STMS_IMPORT/SCC4),
    #   S_TABU_DIS/DICBERCLS=SS, S_TABU_DIS/ACTVT=02,
    #   S_TRANSPORT/ACTVT (01/02/50/60/06/43),
    #   S_CTS_ADMI/CTS_ADMFCT (IMPT/IMPA/IMP*).
    # ------------------------------------------------------------------

    def _build_transport_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.transport_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - העברת שינויים"))
        self.transport_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        transport_summary_layout = QVBoxLayout(self.transport_summary_group)
        transport_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.transport_summary_table = QTableWidget(0, 5)
        self.transport_summary_table.setItemDelegate(_RightAlignDelegate(self.transport_summary_table))
        self.transport_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _transport_summary_hdr = self.transport_summary_table.horizontalHeader()
        _transport_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _transport_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _transport_summary_hdr.setStretchLastSection(False)
        self.transport_summary_table.setColumnWidth(1, 300)  # ממצא
        self.transport_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.transport_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.transport_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.transport_summary_table.setAlternatingRowColors(True)
        self.transport_summary_table.setMinimumHeight(160)
        self.transport_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאת העברת שינויים")
        )
        self.transport_summary_table.itemSelectionChanged.connect(self._refresh_selected_transport_users)
        transport_summary_layout.addWidget(self.transport_summary_table)
        parent_layout.addWidget(self.transport_summary_group)

        self.transport_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאת העברת שינויים"))
        self.transport_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        transport_users_layout = QVBoxLayout(self.transport_users_group)
        transport_users_layout.setContentsMargins(8, 14, 8, 8)

        self.transport_users_table = QTableWidget(0, 2)
        self.transport_users_table.setItemDelegate(_RightAlignDelegate(self.transport_users_table))
        self.transport_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _transport_users_hdr = self.transport_users_table.horizontalHeader()
        _transport_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _transport_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _transport_users_hdr.setStretchLastSection(True)
        self.transport_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.transport_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.transport_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.transport_users_table.setAlternatingRowColors(True)
        self.transport_users_table.setMinimumHeight(180)
        self.transport_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאת העברת שינויים")
        )
        self.transport_users_table.cellDoubleClicked.connect(self.show_transport_user_dialog)
        transport_users_layout.addWidget(self.transport_users_table)
        parent_layout.addWidget(self.transport_users_group)

        self._refresh_transport_summary_table()

    def _compute_transport_permissions(self) -> None:
        """Recompute transport/change-management permission findings from cached AGR_1251 + AGR_USERS rows.

        OR logic: any single qualifying row in AGR_1251 makes the AGR_NAME a match.
        """
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-TRANSPORT-01"
        control_meta = get_audit_control_definition(control_id)

        # Step A: find AGR_NAMEs that carry transport permission objects.
        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in TRANSPORT_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            if low_str == "*" or high_str == "*":
                qualifies = True
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        # Step B: cross-join with AGR_USERS to collect users per client.
        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        # Step C+D: store results.
        self.transport_summary_records.clear()
        self.transport_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.transport_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאת העברת שינויים",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.transport_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.transport_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאת העברת שינויים",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.transport_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_transport_summary_table()

    def _refresh_transport_summary_table(self) -> None:
        self.transport_summary_table.setRowCount(0)
        self.transport_users_table.setRowCount(0)
        if not self.transport_summary_records:
            self.transport_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.transport_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.transport_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.transport_summary_table.rowCount()
            self.transport_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.transport_summary_table.setItem(row_index, column, item)

        if self.transport_summary_table.rowCount() > 0:
            self.transport_summary_table.selectRow(0)
            self._refresh_selected_transport_users()
        self.transport_summary_table.resizeColumnsToContents()

    def _refresh_selected_transport_users(self) -> None:
        selected_items = self.transport_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.transport_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.transport_users_by_control.get(record_key, [])
        self.transport_users_table.setRowCount(0)

        if not user_rows:
            self.transport_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.transport_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.transport_users_table.rowCount()
            self.transport_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.transport_users_table.setItem(row_index, 0, client_item)
            self.transport_users_table.setItem(row_index, 1, user_item)
        self.transport_users_table.resizeColumnsToContents()

    def show_transport_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.transport_users_table.rowCount():
            return

        user_item = self.transport_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.transport_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.transport_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.transport_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאת העברת שינויים")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # DEBUG Permissions (MA-DEBUG-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # OR logic: any single qualifying AGR_1251 row makes the AGR_NAME a match.
    # Qualifying objects: S_TCODE/TCD (SE38/SA38/SE80/ST05),
    #   S_DEVELOP/OBJTYPE=DEBUG, S_DEVELOP/ACTVT (01/02),
    #   S_PROGRAM/P_ACTION=SUB, S_PROGRAM/P_GROUP (any value),
    #   S_ADMI_FCD/S_ADMI_FCD=PADM.
    # ------------------------------------------------------------------

    def _build_debug_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.debug_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - שימוש ב-DEBUG"))
        self.debug_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        debug_summary_layout = QVBoxLayout(self.debug_summary_group)
        debug_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.debug_summary_table = QTableWidget(0, 5)
        self.debug_summary_table.setItemDelegate(_RightAlignDelegate(self.debug_summary_table))
        self.debug_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _debug_summary_hdr = self.debug_summary_table.horizontalHeader()
        _debug_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _debug_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _debug_summary_hdr.setStretchLastSection(False)
        self.debug_summary_table.setColumnWidth(1, 300)  # ממצא
        self.debug_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.debug_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.debug_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.debug_summary_table.setAlternatingRowColors(True)
        self.debug_summary_table.setMinimumHeight(160)
        self.debug_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאות DEBUG")
        )
        self.debug_summary_table.itemSelectionChanged.connect(self._refresh_selected_debug_users)
        debug_summary_layout.addWidget(self.debug_summary_table)
        parent_layout.addWidget(self.debug_summary_group)

        self.debug_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאות DEBUG"))
        self.debug_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        debug_users_layout = QVBoxLayout(self.debug_users_group)
        debug_users_layout.setContentsMargins(8, 14, 8, 8)

        self.debug_users_table = QTableWidget(0, 2)
        self.debug_users_table.setItemDelegate(_RightAlignDelegate(self.debug_users_table))
        self.debug_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _debug_users_hdr = self.debug_users_table.horizontalHeader()
        _debug_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _debug_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _debug_users_hdr.setStretchLastSection(True)
        self.debug_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.debug_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.debug_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.debug_users_table.setAlternatingRowColors(True)
        self.debug_users_table.setMinimumHeight(180)
        self.debug_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאות DEBUG")
        )
        self.debug_users_table.cellDoubleClicked.connect(self.show_debug_user_dialog)
        debug_users_layout.addWidget(self.debug_users_table)
        parent_layout.addWidget(self.debug_users_group)

        self._refresh_debug_summary_table()

    def _compute_debug_permissions(self) -> None:
        """Recompute DEBUG permission findings from cached AGR_1251 + AGR_USERS rows.

        OR logic: any single qualifying row in AGR_1251 makes the AGR_NAME a match.
        Special wildcard criterion: if the qualifying set for a (OBJECT, FIELD) pair
        contains "*", then any non-empty LOW or HIGH value qualifies.
        """
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-DEBUG-01"
        control_meta = get_audit_control_definition(control_id)

        # Step A: find AGR_NAMEs that carry DEBUG permission objects.
        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in DEBUG_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            if low_str == "*" or high_str == "*":
                qualifies = True
            elif "*" in qualifying_map[key]:
                # Criterion allows any value → any non-empty LOW or HIGH qualifies
                qualifies = bool(low_str) or bool(high_str)
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        # Step B: cross-join with AGR_USERS to collect users per client.
        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        # Step C+D: store results.
        self.debug_summary_records.clear()
        self.debug_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.debug_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאות DEBUG",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.debug_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.debug_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאות DEBUG",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.debug_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_debug_summary_table()

    def _refresh_debug_summary_table(self) -> None:
        self.debug_summary_table.setRowCount(0)
        self.debug_users_table.setRowCount(0)
        if not self.debug_summary_records:
            self.debug_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.debug_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.debug_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.debug_summary_table.rowCount()
            self.debug_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.debug_summary_table.setItem(row_index, column, item)

        if self.debug_summary_table.rowCount() > 0:
            self.debug_summary_table.selectRow(0)
            self._refresh_selected_debug_users()
        self.debug_summary_table.resizeColumnsToContents()

    def _refresh_selected_debug_users(self) -> None:
        selected_items = self.debug_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.debug_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.debug_users_by_control.get(record_key, [])
        self.debug_users_table.setRowCount(0)

        if not user_rows:
            self.debug_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.debug_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.debug_users_table.rowCount()
            self.debug_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.debug_users_table.setItem(row_index, 0, client_item)
            self.debug_users_table.setItem(row_index, 1, user_item)
        self.debug_users_table.resizeColumnsToContents()

    def show_debug_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.debug_users_table.rowCount():
            return

        user_item = self.debug_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.debug_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.debug_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.debug_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאות DEBUG")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    # ------------------------------------------------------------------
    # Job-Management Permissions (MA-JOBMGMT-01)
    # Cross-join: AGR_1251 (permission objects) × AGR_USERS (role assignments)
    # OR logic: any single qualifying AGR_1251 row makes the AGR_NAME a match.
    # Qualifying objects: S_TCODE/TCD (SE37/SE36/SM36/SE35/SE30/SE34/SHDB/SM36WIZ),
    #   S_BTCH_ADM/BTCADMIN=Y, S_BTCH_JOB/JOBACTION (DELE/RELE/PROT),
    #   S_BTCH_NAM/BTCUNAME (any value), S_BTCH_MONI/BSCAKTI (DELE/RELE).
    # ------------------------------------------------------------------

    def _build_job_mgmt_permissions_section(self, parent_layout: QVBoxLayout) -> None:
        self.job_mgmt_summary_group = QGroupBox(self.format_ui_rtl_text("ממצאי הרשאות - עידכון ג'ובים"))
        self.job_mgmt_summary_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        job_mgmt_summary_layout = QVBoxLayout(self.job_mgmt_summary_group)
        job_mgmt_summary_layout.setContentsMargins(8, 14, 8, 8)

        self.job_mgmt_summary_table = QTableWidget(0, 5)
        self.job_mgmt_summary_table.setItemDelegate(_RightAlignDelegate(self.job_mgmt_summary_table))
        self.job_mgmt_summary_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("ממצא"),
            self.format_rtl_text("כמות משתמשים"),
            self.format_rtl_text("רמת סיכון"),
            self.format_rtl_text("סטטוס"),
        ])
        _job_mgmt_summary_hdr = self.job_mgmt_summary_table.horizontalHeader()
        _job_mgmt_summary_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _job_mgmt_summary_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _job_mgmt_summary_hdr.setStretchLastSection(False)
        self.job_mgmt_summary_table.setColumnWidth(1, 300)  # ממצא
        self.job_mgmt_summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.job_mgmt_summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.job_mgmt_summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.job_mgmt_summary_table.setAlternatingRowColors(True)
        self.job_mgmt_summary_table.setMinimumHeight(160)
        self.job_mgmt_summary_table.setToolTip(
            self.format_ui_rtl_text("לחיצה על שורה תציג את המשתמשים בעלי הרשאות עידכון ג'ובים")
        )
        self.job_mgmt_summary_table.itemSelectionChanged.connect(self._refresh_selected_job_mgmt_users)
        job_mgmt_summary_layout.addWidget(self.job_mgmt_summary_table)
        parent_layout.addWidget(self.job_mgmt_summary_group)

        self.job_mgmt_users_group = QGroupBox(self.format_ui_rtl_text("משתמשים בעלי הרשאות עידכון ג'ובים"))
        self.job_mgmt_users_group.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        job_mgmt_users_layout = QVBoxLayout(self.job_mgmt_users_group)
        job_mgmt_users_layout.setContentsMargins(8, 14, 8, 8)

        self.job_mgmt_users_table = QTableWidget(0, 2)
        self.job_mgmt_users_table.setItemDelegate(_RightAlignDelegate(self.job_mgmt_users_table))
        self.job_mgmt_users_table.setHorizontalHeaderLabels([
            self.format_rtl_text("קליינט"),
            self.format_rtl_text("משתמש"),
        ])
        _job_mgmt_users_hdr = self.job_mgmt_users_table.horizontalHeader()
        _job_mgmt_users_hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        _job_mgmt_users_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        _job_mgmt_users_hdr.setStretchLastSection(True)
        self.job_mgmt_users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.job_mgmt_users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.job_mgmt_users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.job_mgmt_users_table.setAlternatingRowColors(True)
        self.job_mgmt_users_table.setMinimumHeight(180)
        self.job_mgmt_users_table.setToolTip(
            self.format_ui_rtl_text("לחיצה כפולה על משתמש תציג את הרולים המעניקים לו הרשאות ג'ובים")
        )
        self.job_mgmt_users_table.cellDoubleClicked.connect(self.show_job_mgmt_user_dialog)
        job_mgmt_users_layout.addWidget(self.job_mgmt_users_table)
        parent_layout.addWidget(self.job_mgmt_users_group)

        self._refresh_job_mgmt_summary_table()

    def _compute_job_mgmt_permissions(self) -> None:
        """Recompute job-management permission findings from cached AGR_1251 + AGR_USERS rows.

        OR logic: any single qualifying row in AGR_1251 makes the AGR_NAME a match.
        Special wildcard criterion: if the qualifying set for a (OBJECT, FIELD) pair
        contains "*", then any non-empty LOW or HIGH value qualifies.
        """
        if not self.agr_1251_cached_rows or not self.agr_users_cached_rows:
            return

        control_id = "MA-JOBMGMT-01"
        control_meta = get_audit_control_definition(control_id)

        qualifying_map: dict[tuple[str, str], set[str]] = {
            (obj.upper(), fld.upper()): {v.upper() for v in vals}
            for (obj, fld), vals in JOB_MGMT_PERMISSION_CRITERIA.items()
        }

        # agr_name_objects: AGR_NAME -> set of (OBJECT, FIELD, LOW_display) tuples that qualified
        agr_name_objects: dict[str, set[tuple[str, str, str]]] = {}
        for row in self.agr_1251_cached_rows:
            obj_val = self._resolve_row_value_by_priority(row, "OBJECT")
            fld_val = self._resolve_row_value_by_priority(row, "FIELD")
            low_val = self._resolve_row_value_by_priority(row, "LOW")
            high_val = self._resolve_row_value_by_priority(row, "HIGH")
            if obj_val is None or fld_val is None:
                continue
            obj_upper = str(obj_val).strip().upper()
            fld_upper = str(fld_val).strip().upper()
            key = (obj_upper, fld_upper)
            if key not in qualifying_map:
                continue
            low_str = str(low_val).strip().upper() if low_val is not None else ""
            high_str = str(high_val).strip().upper() if high_val is not None else ""
            if low_str == "*" or high_str == "*":
                qualifies = True
            elif "*" in qualifying_map[key]:
                # Criterion allows any value → any non-empty LOW or HIGH qualifies
                qualifies = bool(low_str) or bool(high_str)
            else:
                qualifies = bool(low_str and low_str in qualifying_map[key]) or bool(
                    high_str and high_str in qualifying_map[key]
                )
            if not qualifies:
                continue
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None or not str(agr_name_val).strip():
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            low_display = low_str if low_str else "-"
            agr_name_objects.setdefault(agr_name_upper, set()).add((obj_upper, fld_upper, low_display))

        matching_agr_names: set[str] = set(agr_name_objects.keys())

        # Cross-join with AGR_USERS to collect users per client.
        users_by_client: dict[str, dict[str, set[str]]] = {}
        for row in self.agr_users_cached_rows:
            agr_name_val = self._resolve_row_value_by_priority(row, "AGR_NAME")
            if agr_name_val is None:
                continue
            agr_name_upper = str(agr_name_val).strip().upper()
            if agr_name_upper not in matching_agr_names:
                continue

            mandt_val = self._resolve_row_value_by_priority(row, "MANDT")
            if mandt_val is not None and str(mandt_val).strip():
                mandt = str(mandt_val).strip()
            else:
                source_file = str(row.get("__source_file", ""))
                digits_match = re.search(r"\d{3}", Path(source_file).name)
                mandt = digits_match.group(0) if digits_match else "-"

            uname_val = self._resolve_row_value_by_priority(row, "UNAME")
            if uname_val is None or not str(uname_val).strip():
                continue
            uname = str(uname_val).strip().upper()

            client_users = users_by_client.setdefault(mandt, {})
            client_users.setdefault(uname, set()).add(agr_name_upper)

        self.job_mgmt_summary_records.clear()
        self.job_mgmt_users_by_control.clear()

        if not users_by_client:
            record_key = f"{control_id}|-"
            self.job_mgmt_summary_records[record_key] = {
                "record_key": record_key,
                "client": "-",
                "finding_text": "לא נמצאו משתמשים בעלי הרשאות עידכון ג'ובים",
                "users_count": 0,
                "risk_level": control_meta.get("risk_level", "-"),
                "status": "תקין",
            }
            self.job_mgmt_users_by_control[record_key] = []
        else:
            for mandt, client_users in sorted(users_by_client.items()):
                users_count = len(client_users)
                record_key = f"{control_id}|{mandt}"
                self.job_mgmt_summary_records[record_key] = {
                    "record_key": record_key,
                    "client": mandt,
                    "finding_text": f"נמצאו {users_count} משתמשים בעלי הרשאות עידכון ג'ובים",
                    "users_count": users_count,
                    "risk_level": control_meta.get("risk_level", "-"),
                    "status": "עם ממצא" if users_count > 0 else "תקין",
                }
                self.job_mgmt_users_by_control[record_key] = [
                    {
                        "client": mandt,
                        "user_name": uname,
                        "roles": [
                            {
                                "agr_name": r,
                                "objects": sorted(agr_name_objects.get(r, set())),
                            }
                            for r in sorted(roles)
                        ],
                    }
                    for uname, roles in sorted(client_users.items())
                ]

        self._refresh_job_mgmt_summary_table()

    def _refresh_job_mgmt_summary_table(self) -> None:
        self.job_mgmt_summary_table.setRowCount(0)
        self.job_mgmt_users_table.setRowCount(0)
        if not self.job_mgmt_summary_records:
            self.job_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("יש לטעון קבצי AGR_1251 ו-AGR_USERS"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.job_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for row_data in sorted(
            self.job_mgmt_summary_records.values(),
            key=lambda item: str(item.get("client", "")),
        ):
            row_index = self.job_mgmt_summary_table.rowCount()
            self.job_mgmt_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("client", "-")),
                str(row_data.get("finding_text", "-")),
                str(row_data.get("users_count", 0)),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("status", "-")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("record_key", ""))
                self.job_mgmt_summary_table.setItem(row_index, column, item)

        if self.job_mgmt_summary_table.rowCount() > 0:
            self.job_mgmt_summary_table.selectRow(0)
            self._refresh_selected_job_mgmt_users()
        self.job_mgmt_summary_table.resizeColumnsToContents()

    def _refresh_selected_job_mgmt_users(self) -> None:
        selected_items = self.job_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.job_mgmt_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        record_key = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        user_rows = self.job_mgmt_users_by_control.get(record_key, [])
        self.job_mgmt_users_table.setRowCount(0)

        if not user_rows:
            self.job_mgmt_users_table.insertRow(0)
            empty_item = QTableWidgetItem(self.format_rtl_text("לא נמצאו משתמשים להצגה"))
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.job_mgmt_users_table.setItem(0, 1, empty_item)
            return

        for user_data in user_rows:
            row_index = self.job_mgmt_users_table.rowCount()
            self.job_mgmt_users_table.insertRow(row_index)
            client_name = str(user_data.get("client", "-") or "-")
            user_name = str(user_data.get("user_name", "-") or "-")
            client_item = QTableWidgetItem(self.format_rtl_text(client_name))
            client_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item = QTableWidgetItem(self.format_rtl_text(user_name))
            user_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            user_item.setData(Qt.ItemDataRole.UserRole, {"client": client_name, "user": user_name})
            self.job_mgmt_users_table.setItem(row_index, 0, client_item)
            self.job_mgmt_users_table.setItem(row_index, 1, user_item)
        self.job_mgmt_users_table.resizeColumnsToContents()

    def show_job_mgmt_user_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.job_mgmt_users_table.rowCount():
            return

        user_item = self.job_mgmt_users_table.item(row_index, 1)
        if user_item is None:
            return

        user_payload = user_item.data(Qt.ItemDataRole.UserRole)
        user_name = ""
        client_name = "-"
        if isinstance(user_payload, dict):
            user_name = str(user_payload.get("user", "") or "").strip()
            client_name = str(user_payload.get("client", "-") or "-").strip()
        if not user_name:
            user_name = str(user_item.text() or "").strip()
        if not user_name or user_name.startswith("לא נמצאו") or user_name.startswith("אין "):
            return

        selected_items = self.job_mgmt_summary_table.selectedItems()
        if not selected_items:
            return

        summary_item = self.job_mgmt_summary_table.item(selected_items[0].row(), 0)
        if summary_item is None:
            return

        record_key = str(summary_item.data(Qt.ItemDataRole.UserRole) or summary_item.text())
        user_rows = self.job_mgmt_users_by_control.get(record_key, [])
        roles: list[dict] = []
        for user_data in user_rows:
            if (
                str(user_data.get("user_name", "")).strip().upper() == user_name.upper()
                and str(user_data.get("client", "-") or "-").strip() == client_name
            ):
                roles = list(user_data.get("roles", []))
                break

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט הרשאות עידכון ג'ובים")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(640, 480)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        lines = [
            f"קליינט: {client_name}",
            f"משתמש: {user_name}",
            "",
            "רולים ואובייקטי הרשאה:",
        ]
        if roles:
            for role_entry in roles:
                lines.append(f"- {role_entry.get('agr_name', '')}")
                for obj, fld, low in role_entry.get("objects", []):
                    lines.append(f"    {obj} | {fld} | {low}")
        else:
            lines.append("- לא נמצאו רולים להצגה")
        details_box.setPlainText(self.format_rtl_text("\n".join(lines)))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _build_audit_detail_row(
        self,
        issue: ValidationIssue | None,
        control_id: str,
        source_file: str,
        extraction_date: str,
        control_meta: dict[str, str],
        control_snapshot: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        if issue is None:
            return {
                "control_id": control_id,
                "source_file": source_file,
                "extraction_date": extraction_date,
                "work_environment": self._current_work_environment_label(),
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
            "work_environment": self._current_work_environment_label(),
            "category": issue.category or control_meta.get("category", "-"),
            "risk_level": issue.risk_level or control_meta.get("risk_level", "-"),
            "description": issue.description or control_meta.get("description", "-"),
            "check_type": issue.check_type or control_meta.get("check_type", "-"),
            "actual_value": issue.actual_value or "-",
            "expected_value": issue.expected_value or "-",
            "status": issue.status or "עם ממצא",
            "full_description": issue.full_description or issue.message,
        }

    def _upsert_audit_control_data(
        self,
        slot_key: str,
        result: Any,
        audit_issues: list[ValidationIssue],
        extraction_date: str,
    ) -> None:
        control_ids = [issue.control_id for issue in audit_issues if issue.control_id]
        expected_controls = get_profile_audit_controls(getattr(result, "detected_profile", slot_key))
        all_control_ids = sorted(set(control_ids + expected_controls))
        if not all_control_ids:
            return

        source_file_label = ", ".join(getattr(result, "source_files", []) or [self._get_slot_display_name(slot_key)])
        detected_profile = str(getattr(result, "detected_profile", slot_key) or slot_key).upper()
        password_snapshots = self._build_password_control_snapshots(getattr(result, "rows", [])) if detected_profile in {"RSPARAM", "TPFET"} else {}
        for control_id in all_control_ids:
            control_meta = get_audit_control_definition(control_id)
            control_issues = [issue for issue in audit_issues if issue.control_id == control_id]
            finding_records = len(control_issues)

            if control_id == "44":
                total_records = self._count_stms_control_records(getattr(result, "rows", []))
            else:
                total_records = 1

            if total_records <= 0:
                total_records = max(finding_records, 1)
            valid_records = max(total_records - finding_records, 0)

            self.audit_summary_records[control_id] = {
                "control_id": control_id,
                "check_type": control_meta.get("check_type", "-"),
                "source_file": source_file_label,
                "extraction_date": extraction_date,
                "work_environment": self._current_work_environment_label(),
                "risk_level": control_meta.get("risk_level", "-"),
                "description": control_meta.get("description", "-"),
                "valid_records": valid_records,
                "finding_records": finding_records,
                "total_records": total_records,
            }

            detail_rows = [
                self._build_audit_detail_row(issue, control_id, source_file_label, extraction_date, control_meta)
                for issue in control_issues
            ]
            if not detail_rows:
                detail_rows = [
                    self._build_audit_detail_row(
                        None,
                        control_id,
                        source_file_label,
                        extraction_date,
                        control_meta,
                        password_snapshots.get(control_id),
                    )
                ]
            self.audit_details_by_control[control_id] = detail_rows

    def _sync_user_review_completion_finding(self) -> None:
        control_id = self.REVIEW_COMPLETION_CONTROL_ID
        self.audit_summary_records.pop(control_id, None)
        self.audit_details_by_control.pop(control_id, None)

        preview_rows, reviewed_rows, incomplete_rows = self._get_user_review_completion_snapshot()
        total_rows = len(preview_rows)
        if total_rows <= 0 or not incomplete_rows:
            return

        control_meta = get_audit_control_definition(control_id)
        source_file_label = self._get_slot_display_name("USR02")
        extraction_date = self._get_slot_extraction_date("USR02") or "-"

        self.audit_summary_records[control_id] = {
            "control_id": control_id,
            "check_type": control_meta.get("check_type", "השלמת סקירת משתמשים"),
            "source_file": source_file_label,
            "extraction_date": extraction_date,
            "work_environment": self._current_work_environment_label(),
            "risk_level": control_meta.get("risk_level", "בינוני"),
            "description": control_meta.get("description", "סקירת המשתמשים טרם הושלמה במלואה."),
            "valid_records": reviewed_rows,
            "finding_records": len(incomplete_rows),
            "total_records": total_rows,
        }

        self.audit_details_by_control[control_id] = [
            {
                "control_id": control_id,
                "source_file": source_file_label,
                "extraction_date": extraction_date,
                "work_environment": self._current_work_environment_label(),
                "category": control_meta.get("category", "MA - ניהול גישה"),
                "risk_level": control_meta.get("risk_level", "בינוני"),
                "description": control_meta.get("description", "סקירת המשתמשים טרם הושלמה במלואה."),
                "check_type": control_meta.get("check_type", "השלמת סקירת משתמשים"),
                "actual_value": str(preview_row.get("BNAME", "-")) or "-",
                "expected_value": "השלמת סקירה בהתאם לכלל ההשלמה",
                "status": "עם ממצא",
                "full_description": self._build_user_review_incomplete_reason(preview_row),
            }
            for preview_row in incomplete_rows
        ]

    def _refresh_audit_summary_table(self) -> None:
        self._sync_user_review_completion_finding()
        self.audit_summary_table.setRowCount(0)
        if not self.audit_summary_records:
            self.audit_detail_table.setRowCount(0)
            self.audit_detail_table.insertRow(0)
            for column, value in enumerate(["-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "אין ממצאים להצגה"]):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.audit_detail_table.setItem(0, column, item)
            return

        for row_data in sorted(self.audit_summary_records.values(), key=lambda item: str(item.get("control_id", ""))):
            row_index = self.audit_summary_table.rowCount()
            self.audit_summary_table.insertRow(row_index)
            values = [
                str(row_data.get("control_id", "-")),
                str(row_data.get("check_type", "-")),
                str(row_data.get("source_file", "-")),
                str(row_data.get("extraction_date", "-")),
                str(row_data.get("work_environment", "-")),
                str(row_data.get("risk_level", "-")),
                str(row_data.get("description", "-")),
                str(row_data.get("valid_records", 0)),
                str(row_data.get("finding_records", 0)),
                str(row_data.get("total_records", 0)),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_data.get("control_id", ""))
                self.audit_summary_table.setItem(row_index, column, item)

        if self.audit_summary_table.rowCount() > 0:
            self.audit_summary_table.selectRow(0)
            self._refresh_selected_audit_detail()
        self.audit_summary_table.resizeColumnsToContents()

    def _refresh_selected_audit_detail(self) -> None:
        selected_items = self.audit_summary_table.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        control_item = self.audit_summary_table.item(selected_row, 0)
        if control_item is None:
            return

        control_id = str(control_item.data(Qt.ItemDataRole.UserRole) or control_item.text())
        detail_rows = self.audit_details_by_control.get(control_id, [])
        self.audit_detail_table.setRowCount(0)

        for detail in detail_rows:
            row_index = self.audit_detail_table.rowCount()
            self.audit_detail_table.insertRow(row_index)
            values = [
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
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.audit_detail_table.setItem(row_index, column, item)
        self.audit_detail_table.resizeColumnsToContents()

    def export_audit_findings_to_excel(self, open_after_export: bool = False) -> Path | None:
        if not self.audit_summary_records:
            QMessageBox.warning(self, "אין נתונים לייצוא", "לא קיימים ממצאי ביקורת לייצוא.")
            return None

        summary_rows = sorted(self.audit_summary_records.values(), key=lambda item: str(item.get("control_id", "")))
        detail_rows: list[dict[str, Any]] = []
        for control_id in sorted(self.audit_details_by_control.keys()):
            detail_rows.extend(self.audit_details_by_control.get(control_id, []))

        export_path = self.config.output_dir / f"audit_findings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        ExcelReportWriter.write_audit_findings_report(summary_rows, detail_rows, export_path)
        self.audit_findings_export_path = export_path

        if open_after_export:
            QMessageBox.information(self, "הייצוא הושלם", f"קובץ הממצאים נשמר בהצלחה:\n{export_path}")

        return export_path

    def _build_audit_detail_dialog_text(self, row_index: int) -> str:
        if row_index < 0 or row_index >= self.audit_detail_table.rowCount():
            return self.format_rtl_text("לא נמצא פירוט עבור הרשומה שנבחרה.")

        field_labels = [
            "קובץ מקור",
            "תאריך הפקה",
            "סביבת עבודה",
            "קטגוריה",
            "רמת סיכון",
            "תיאור",
            "סוג בדיקה",
            "ערך בפועל",
            "ערך מצופה",
            "סטטוס",
            "תיאור מלא",
        ]
        lines = ["פירוט ממצא ביקורת:", ""]
        for column, field_label in enumerate(field_labels):
            item = self.audit_detail_table.item(row_index, column)
            field_value = item.text() if item is not None else "-"
            lines.append(f"{field_label}: {field_value}")
        return self.format_rtl_text("\n".join(lines))

    def show_audit_detail_dialog(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= self.audit_detail_table.rowCount():
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט ממצא ביקורת")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(760, 420)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        details_box.setPlainText(self._build_audit_detail_dialog_text(row_index))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _build_log_details(self, row_index: int) -> str:
        if row_index < 0 or row_index >= len(self.run_log_records):
            return "לא נמצא פירוט עבור הרשומה שנבחרה."

        record = self.run_log_records[row_index]
        lines = [
            f"משבצת: {record['slot_key']}",
            f"קבוצת דוחות: {record['report_group']}",
            f"קובץ: {record['file_name']}",
            f"תאריך הפקה: {record['extraction_date']}",
            f"מספר רשומות שנקלטו: {record['row_count']}",
            f"סטטוס: {record['status']}",
            f"מספר שגיאות: {record['error_count']}",
            f"תיאור קצר: {record['error_preview']}",
            f"תאריך בדיקה: {record['date']}",
            f"שעת בדיקה: {record['time']}",
            "",
            "פירוט:",
        ]

        issues = record.get("issues", [])
        intake = [iss for iss in issues if self._is_intake_issue(iss)]
        audit = [iss for iss in issues if not self._is_intake_issue(iss)]

        if not issues:
            lines.append("לא נמצאו שגיאות קליטה וממצאי ביקורת.")
        else:
            if intake:
                lines.append("--- שגיאות קליטה ---")
                for issue in intake:
                    row_label = issue.row_number if issue.row_number > 0 else "מבנה"
                    lines.append(f"- שורה {row_label} / {issue.column_name}: {issue.message}")
            else:
                lines.append("--- שגיאות קליטה: אין ---")

            if audit:
                lines.append("")
                lines.append("--- ממצאי ביקורת (לפירוט ראה טאב 'ביצוע ניתוח לביקורת') ---")
                for issue in audit:
                    row_label = issue.row_number if issue.row_number > 0 else "מבנה"
                    lines.append(f"- שורה {row_label} / {issue.column_name}: {issue.message}")

        return self.format_rtl_text("\n".join(lines))

    def show_log_details(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= len(self.run_log_records):
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט קובץ שנבדק")
        dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dialog.resize(760, 420)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        details_box.setPlainText(self._build_log_details(row_index))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def clear_results(self) -> None:
        self.selected_slot_key = None
        self.load_history = []
        for slot_key, widget_data in self.slot_widgets.items():
            widget_data["selected_paths"] = []
            self._update_slot_path_label(slot_key, [])
            date_edit = widget_data.get("extraction_date_edit")
            if isinstance(date_edit, QLineEdit):
                date_edit.setText(self._default_extraction_date())
        self.required_columns_edit.setText("")
        self.summary_labels["total"].setText("0")
        self.summary_labels["valid"].setText("0")
        self.summary_labels["invalid"].setText("0")
        self.summary_labels["status"].setText("ממתין להרצה")
        self.summary_group.hide()
        self.results_group.hide()
        self.report_path = None
        self.log_export_path = None
        self.audit_findings_export_path = None
        self.report_button.setEnabled(False)
        self.issues_table.setRowCount(0)
        self.audit_summary_records = {}
        self.audit_details_by_control = {}
        self.permissions_summary_records = {}
        self.permissions_users_by_control = {}
        self.audit_summary_table.setRowCount(0)
        self.audit_detail_table.setRowCount(0)
        self.permissions_summary_table.setRowCount(0)
        self.permissions_users_table.setRowCount(0)
        self.run_log_records = []
        self.run_log_table.setRowCount(0)
        self.refresh_user_preview()
        self.tabs.setCurrentIndex(0)

    def export_run_log_to_excel(self, open_after_export: bool = False) -> Path | None:
        if not self.run_log_records:
            QMessageBox.warning(self, "אין נתונים לייצוא", "טרם תועדו קבצים שנבדקו לייצוא לאקסל.")
            return None

        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = self.format_rtl_text("קבצים שנבדקו")
        headers = [
            "משבצת",
            "קבוצת דוחות",
            "קובץ",
            "תאריך הפקה",
            "רשומות שנקלטו",
            "סטטוס",
            "מספר שגיאות",
            "תיאור שגיאה",
            "תאריך בדיקה",
            "שעת בדיקה",
        ]
        sheet.append(headers)

        for record in self.run_log_records:
            sheet.append([
                record["slot_key"],
                record["report_group"],
                record["file_name"],
                record["extraction_date"],
                record["row_count"],
                record["status"],
                record["error_count"],
                record["error_preview"],
                record["date"],
                record["time"],
            ])

        export_path = self.config.output_dir / f"intake_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        workbook.save(export_path)
        self.log_export_path = export_path

        if open_after_export:
            QMessageBox.information(self, "הייצוא הושלם", f"קובץ התיעוד נשמר בהצלחה:\n{export_path}")

        return export_path

    def export_user_preview_to_excel(self, open_after_export: bool = False) -> Path | None:
        usr02_rows = self._load_preview_rows("USR02")
        combined_rows = self._load_preview_rows("ADR6_USR21")
        if not usr02_rows and not combined_rows:
            QMessageBox.warning(self, "אין נתונים לייצוא", "טרם נטענו משתמשים לסקירה לצורך ייצוא לאקסל.")
            return None

        all_preview_rows = self._build_user_preview_rows(usr02_rows, combined_rows)
        if not all_preview_rows:
            QMessageBox.warning(self, "אין נתונים לייצוא", "טרם נטענו משתמשים לסקירה לצורך ייצוא לאקסל.")
            return None

        field_to_col_def = {col["field"]: col for col in self.USER_PREVIEW_COLUMN_DEFINITIONS}
        export_field_names = [f for f in self.EXPORT_REVIEW_FIELDS if f in field_to_col_def]
        export_formal_names = [str(field_to_col_def[f]["formal"]) for f in export_field_names]

        sorted_rows = sorted(all_preview_rows, key=self._export_sort_key)

        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = self.format_rtl_text("סקירת משתמשים")
        sheet.append(export_formal_names)

        review_status_col_index: int | None = None
        technical_notes_col_index: int | None = None
        for idx, field in enumerate(export_field_names):
            if field == "REVIEW_STATUS":
                review_status_col_index = idx + 1  # 1-based Excel column
            elif field in {"TECH_REVIEW_NOTES", "REVIEW_NOTES"}:
                technical_notes_col_index = idx + 1  # 1-based Excel column

        total_data_rows = len(sorted_rows)
        for preview_row in sorted_rows:
            sheet.append([
                preview_row.get(field, "") or ""
                for field in export_field_names
            ])

        if review_status_col_index is not None and total_data_rows > 0:
            from openpyxl.utils import get_column_letter
            col_letter = get_column_letter(review_status_col_index)
            dv_formula = '"' + ",".join(self.REVIEW_STATUS_OPTIONS) + '"'
            dv = DataValidation(
                type="list",
                formula1=dv_formula,
                allow_blank=True,
                showDropDown=False,
            )
            dv.sqref = f"{col_letter}2:{col_letter}{total_data_rows + 1}"
            sheet.add_data_validation(dv)

        if total_data_rows > 0 and (review_status_col_index is not None or technical_notes_col_index is not None):
            from openpyxl.utils import get_column_letter  # noqa: F811
            warning_fill = PatternFill("solid", fgColor="FFF0C2")
            for excel_row_idx, preview_row in enumerate(sorted_rows, start=2):
                is_not_ok = preview_row.get("REVIEW_STATUS", "") == "נבדק - לא תקין"
                has_notes = bool((preview_row.get("TECH_REVIEW_NOTES", "") or "").strip())
                if is_not_ok and not has_notes:
                    if review_status_col_index is not None:
                        sheet.cell(row=excel_row_idx, column=review_status_col_index).fill = warning_fill
                    if technical_notes_col_index is not None:
                        sheet.cell(row=excel_row_idx, column=technical_notes_col_index).fill = warning_fill

        export_path = self.config.output_dir / f"users_review_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        workbook.save(export_path)
        self.user_preview_export_path = export_path

        if open_after_export:
            QMessageBox.information(self, "הייצוא הושלם", f"קובץ הסקירה נשמר בהצלחה:\n{export_path}")

        return export_path

    def import_user_review_from_excel(self) -> None:
        initial_directory = self._get_last_file_dialog_directory()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "ייבוא סקירת משתמשים מאקסל",
            initial_directory,
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not file_path:
            return

        try:
            workbook = load_workbook(file_path, read_only=True, data_only=True)
        except Exception as exc:
            QMessageBox.warning(self, "שגיאת ייבוא", f"לא ניתן לפתוח את קובץ האקסל:\n{exc}")
            return

        sheet = workbook.active
        if sheet is None:
            QMessageBox.warning(self, "שגיאת ייבוא", "הגיליון הפעיל בקובץ ריק או לא נמצא.")
            workbook.close()
            return

        rows_iter = iter(sheet.iter_rows(values_only=True))
        raw_headers = next(rows_iter, None)
        if raw_headers is None:
            QMessageBox.warning(self, "שגיאת ייבוא", "הקובץ ריק - לא נמצאו כותרות.")
            workbook.close()
            return

        headers = [str(h).strip() if h is not None else "" for h in raw_headers]

        FORMAL_TO_FIELD = {
            str(col["formal"]).strip(): str(col["field"])
            for col in self.USER_PREVIEW_COLUMN_DEFINITIONS
        }
        TECHNICAL_FIELDS = {str(col["field"]) for col in self.USER_PREVIEW_COLUMN_DEFINITIONS}

        def _resolve(header: str) -> str | None:
            if header in TECHNICAL_FIELDS:
                return header
            return FORMAL_TO_FIELD.get(header)

        col_map: dict[str, int] = {}
        for col_idx, header in enumerate(headers):
            field = _resolve(header)
            if field and field not in col_map:
                col_map[field] = col_idx

        expected_import_fields = list(self.EXPORT_REVIEW_FIELDS)
        missing = [field_name for field_name in expected_import_fields if field_name not in col_map]
        non_empty_headers = [header for header in headers if header]
        if missing or len(non_empty_headers) < len(expected_import_fields):
            missing_labels = [
                str(self._get_user_preview_column_definition(field_name).get("formal", field_name))
                for field_name in missing
            ]
            QMessageBox.warning(
                self,
                "שגיאת ייבוא",
                "קובץ הסקירה אינו תואם לתבנית המעודכנת של המערכת.\n"
                f"מספר עמודות מזוהות: {len(non_empty_headers)} מתוך {len(expected_import_fields)} נדרשות.\n"
                f"עמודות חסרות: {', '.join(missing_labels) if missing_labels else '-'}\n"
                "יש לייבא קובץ אקסל שיוצא מהכלי לאחר השינוי האחרון במבנה דוח הסקירה.",
            )
            workbook.close()
            return

        mandt_col = col_map.get("MANDT")
        bname_col = col_map["BNAME"]
        status_col = col_map["REVIEW_STATUS"]
        tech_notes_col = col_map.get("TECH_REVIEW_NOTES")
        if tech_notes_col is None:
            tech_notes_col = col_map.get("REVIEW_NOTES")
        business_notes_col = col_map.get("BUS_REVIEW_NOTES")

        imported_count = 0
        for row_values in rows_iter:
            bname = str(row_values[bname_col]).strip() if bname_col < len(row_values) and row_values[bname_col] is not None else ""
            if not bname:
                continue
            mandt = str(row_values[mandt_col]).strip() if mandt_col is not None and mandt_col < len(row_values) and row_values[mandt_col] is not None else ""
            review_key = self._user_reviewer_state_key(mandt, bname)

            raw_status = row_values[status_col] if status_col < len(row_values) else None
            status_value = self._normalize_reviewer_status(str(raw_status).strip() if raw_status is not None else "")

            raw_tech_notes = row_values[tech_notes_col] if tech_notes_col is not None and tech_notes_col < len(row_values) else None
            tech_notes_value = str(raw_tech_notes).strip() if raw_tech_notes is not None else ""
            raw_business_notes = row_values[business_notes_col] if business_notes_col is not None and business_notes_col < len(row_values) else None
            business_notes_value = str(raw_business_notes).strip() if raw_business_notes is not None else ""

            current = self.user_reviewer_state.setdefault(review_key, self._default_reviewer_values().copy())
            current["REVIEW_STATUS"] = status_value
            current["TECH_REVIEW_NOTES"] = tech_notes_value
            current["BUS_REVIEW_NOTES"] = business_notes_value
            imported_count += 1

        workbook.close()
        self._save_user_reviewer_state()
        self.refresh_user_preview()

        QMessageBox.information(
            self,
            "הייבוא הושלם",
            f"יובאו בהצלחה {imported_count} שורות מקובץ הסקירה.",
        )

    def _email_from_settings(self, settings_key: str) -> str:
        settings = self._current_system_settings()
        return str(settings.get(settings_key, "")).strip()

    @staticmethod
    def _validate_email_address(email_value: str) -> bool:
        normalized = email_value.strip()
        return bool(normalized and "@" in normalized and "." in normalized.split("@")[-1])

    def _create_outlook_review_draft(self, recipient_email: str, role_label: str) -> None:
        if not self._validate_email_address(recipient_email):
            QMessageBox.warning(self, "מייל לא מוגדר", f"לא הוגדרה כתובת מייל תקינה עבור {role_label} במסך ההגדרות.")
            return

        export_path = self.export_user_preview_to_excel(open_after_export=False)
        if export_path is None:
            return

        if not sys.platform.startswith("win"):
            QMessageBox.warning(self, "מערכת לא נתמכת", "יצירת טיוטת מייל נתמכת כרגע ב-Windows בלבד.")
            return

        try:
            import win32com.client  # type: ignore[import-not-found]
        except ModuleNotFoundError:
            install_command = f'"{sys.executable}" -m pip install pywin32'
            QMessageBox.warning(
                self,
                "רכיב חסר ל-Outlook",
                "לא ניתן ליצור טיוטת Outlook כי חסרה חבילת pywin32.\n\n"
                f"יש להריץ פעם אחת בסביבת העבודה:\n{install_command}",
            )
            return
        except Exception as error:
            QMessageBox.warning(self, "Outlook לא זמין", f"לא ניתן לטעון Outlook COM ליצירת טיוטה:\n{error}")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail_item = outlook.CreateItem(0)
            mail_item.To = recipient_email
            mail_item.Subject = self.format_rtl_text(f"סקירת דוח משתמשים - {datetime.now().strftime('%Y-%m-%d')}")
            mail_item.HTMLBody = (
                "<div dir='rtl' style='text-align:right; font-family:Arial, sans-serif; font-size:12pt;'>"
                "<p>שלום,</p>"
                "<p>מצורף דוח סקירת משתמשים עדכני מתוך המערכת.</p>"
                "<p>נא לעבור על הממצאים ולעדכן סטטוס/הערות בהתאם.</p>"
                "<p>בברכה,<br>מערכת בקרות ITGC</p>"
                "</div>"
            )
            mail_item.Attachments.Add(str(export_path))
            mail_item.Display(False)
        except Exception as error:
            QMessageBox.warning(self, "שגיאת מייל", f"יצירת טיוטת מייל נכשלה:\n{error}")
            return

        QMessageBox.information(
            self,
            "טיוטת מייל נוצרה",
            f"נוצרה טיוטה ל-{role_label} עם קובץ מצורף:\n{export_path}",
        )

    def draft_user_review_email_to_business(self) -> None:
        self._create_outlook_review_draft(
            recipient_email=self._email_from_settings("business_reviewer_email"),
            role_label="גורם עסקי",
        )

    def draft_user_review_email_to_technical(self) -> None:
        self._create_outlook_review_draft(
            recipient_email=self._email_from_settings("technical_reviewer_email"),
            role_label="גורם טכני",
        )

    def open_output_folder(self) -> None:
        self._open_path(self.config.output_dir)

    def open_report(self) -> None:
        if self.report_path and self.report_path.exists():
            self._open_path(self.report_path)
        else:
            QMessageBox.warning(self, "דוח לא זמין", "טרם נוצר דוח אקסל לפתיחה.")

    @staticmethod
    def _open_path(path: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return
        subprocess.run(["xdg-open", str(path)], check=False)


def launch_desktop_app() -> None:
    app = get_qt_app()
    window = ValidationDesktopApp()
    window.show()
    app.exec()
