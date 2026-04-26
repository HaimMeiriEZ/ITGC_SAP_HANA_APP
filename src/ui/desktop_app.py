import json
import os
import re
import subprocess
import sys
import copy
from datetime import datetime, date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
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
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QPlainTextEdit,
    QTabWidget,
    QTextEdit,
    QProgressBar,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from src.config import AppConfig
from src.models.validation_result import ValidationIssue
from src.pipeline import process_file
from src.validators.spec_rules import get_column_aliases


def get_qt_app() -> QApplication:
    if "unittest" in sys.modules and "QT_QPA_PLATFORM" not in os.environ:
        os.environ["QT_QPA_PLATFORM"] = "offscreen"

    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        app.setLayoutDirection(Qt.RightToLeft)
        app.setFont(QFont("Segoe UI", 10))
    return app


class SortableTableWidgetItem(QTableWidgetItem):
    SORT_ROLE = Qt.UserRole + 2

    def __lt__(self, other: object) -> bool:
        if isinstance(other, QTableWidgetItem):
            self_sort_value = self.data(self.SORT_ROLE)
            other_sort_value = other.data(self.SORT_ROLE)
            if self_sort_value is not None or other_sort_value is not None:
                left = "" if self_sort_value is None else str(self_sort_value)
                right = "" if other_sort_value is None else str(other_sort_value)
                return left < right
        return super().__lt__(other)


class ValidationDesktopApp(QMainWindow):
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
        {"field": "REVIEW_NOTES", "formal": "הערות סוקר", "technical": "REVIEW_NOTES", "source": "סוקר", "default": True, "width": 220},
        {"field": "UFLAG", "formal": "קוד נעילה", "technical": "UFLAG", "source": "USR02", "default": False, "width": 100},
    ]
    DEFAULT_USER_PREVIEW_COLUMNS = [
        column["field"]
        for column in USER_PREVIEW_COLUMN_DEFINITIONS
        if bool(column.get("default"))
    ]
    CURRENT_USER_PREVIEW_SETTINGS_VERSION = 5
    USER_PREVIEW_SETTINGS_MIGRATIONS = {
        2: ["PWDINITIAL", "PWDCHGDATE", "PWDSETDATE"],
        3: ["DEPARTMENT", "GLTGV", "GLTGB", "USTYP", "LOCNT", "OCOD1", "PASSCODE", "PWDSALTEDHASH", "SECURITY_POLICY"],
        4: ["REVIEW_STATUS", "REVIEW_NOTES"],
        5: ["FINDINGS_DESCRIPTION"],
    }
    USER_PREVIEW_FILTER_OPTIONS = [
        ("all", "כלל האוכלוסייה"),
        ("active", "פעילים בתקופה הנבדקת"),
        ("inactive", "לא פעילים בתקופה הנבדקת"),
    ]
    REVIEW_STATUS_OPTIONS = ["טרם נבדק", "נבדק - תקין", "נבדק - לא תקין"]
    DEFAULT_REVIEW_STATUS = "טרם נבדק"
    REVIEWED_STATUSES = {"נבדק - תקין", "נבדק - לא תקין"}
    USER_TYPE_RULES = {
        "Dialog": ["A"],
        "System": ["B"],
        "Communication": ["C"],
        "Service": ["S"],
        "Reference": ["L"],
    }
    USER_PREVIEW_DATE_FIELDS = {"TRDAT", "PWDCHGDATE", "PWDSETDATE", "GLTGV", "GLTGB"}
    EXPORT_REVIEW_FIELDS = [
        "MANDT", "BNAME", "NAME_TEXTC", "SMTP_ADDR", "STATUS", "USTYP",
        "GLTGV", "GLTGB", "TRDAT", "PWDSETDATE", "PWDCHGDATE",
        "FINDINGS_DESCRIPTION", "REVIEW_STATUS", "REVIEW_NOTES",
    ]

    SLOT_DEFINITIONS = {
        "USR02": {
            "category": "טבלאות משתמשים",
            "description": "משתמשים - מקור חובה לבדיקות גישה, סטטוס ותאריכי התחברות.",
            "expected_file": "usr02_100.txt",
            "required": True,
        },
        "ADR6_USR21": {
            "label": "ADR6 / USER_ADDR",
            "category": "טבלאות משתמשים",
            "description": "ניתן להזין קובצי ADR6 או USER_ADDR או את שניהם יחד לצורך העשרת נתוני המשתמשים מתוך USR02.",
            "expected_file": "adr6.txt או user_addr.txt",
            "required": False,
        },
        "AGR_USERS": {
            "category": "טבלאות הרשאות כלליות",
            "description": "רולים-משתמשים - מיפוי המשתמשים לרולים במערכת.",
            "expected_file": "agr_users_100.txt",
            "required": True,
        },
        "AGR_1251": {
            "category": "טבלאות הרשאות כלליות",
            "description": "רולים-אובייקטי הרשאה - זיהוי אובייקטי הרשאות רגישים.",
            "expected_file": "agr_1251_100.txt",
            "required": True,
        },
        "AGR_1252": {
            "category": "טבלאות הרשאות כלליות",
            "description": "רולים-טרנזקציות - זיהוי גישות עסקיות וטרנזקציות.",
            "expected_file": "agr_1252_100.txt",
            "required": False,
        },
        "AGR_DEFINE": {
            "category": "טבלאות הרשאות כלליות",
            "description": "רולים מורחב - מידע כללי על הגדרת הרול.",
            "expected_file": "agr_define.txt",
            "required": False,
        },
        "UST04": {
            "category": "טבלאות הרשאות כלליות",
            "description": "פרופילים-משתמשים - שיוך פרופילים ישיר למשתמשים.",
            "expected_file": "ust04.txt",
            "required": False,
        },
        "E070": {
            "category": "טבלאות שינויים",
            "description": "רשימת שינויים - נתוני transport requests ושינויים בסביבה.",
            "expected_file": "e070_100.txt",
            "required": True,
        },
        "T000": {
            "category": "טבלאות שינויים",
            "description": "לוג פעילות שינוי SCC4 - בקרות שינוי ברמת client.",
            "expected_file": "t000.txt",
            "required": False,
        },
        "STMS": {
            "category": "טבלאות שינויים",
            "description": "רשימת שינויים שהועברה דרך SCC4 או STMS.",
            "expected_file": "stms.txt",
            "required": False,
        },
        "RSPARAM": {
            "category": "מדיניות סיסמאות",
            "description": "פרמטרים סיסטמאיים - פרמטרי אבטחה והקשחת מערכת.",
            "expected_file": "rsparam.xlsx",
            "required": True,
        },
        "TPFET": {
            "label": "TPFET / RZ10",
            "category": "מדיניות סיסמאות",
            "description": "פרמטרים סיסטמאיים נוספים, כולל פרופילי login כגון RZ10.",
            "expected_file": "rz10.txt",
            "required": False,
        },
    }

    SETTINGS_SECTION_DEPENDENCIES = {
        "user_review_period": set(),
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
        self.slot_widgets: dict[str, dict[str, object]] = {}
        self.category_run_buttons: dict[str, QPushButton] = {}
        self.category_sections: dict[str, QGroupBox] = {}
        self.selected_slot_key: str | None = None
        self.load_history: list[str] = []
        self.summary_labels: dict[str, QLabel] = {}
        self.run_log_records: list[dict[str, object]] = []
        self._allow_user_preview_persistence = base_dir is not None or "unittest" not in sys.modules
        self.last_file_dialog_directory = self._load_last_file_dialog_directory()
        self._refreshing_user_preview = False
        self.user_preview_export_path: Path | None = None
        self.user_reviewer_state = self._load_user_reviewer_state()
        self.user_preview_visible_columns = self._load_user_preview_column_selection()
        self.system_settings_widgets: dict[str, object] = {}
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
        self.setLayoutDirection(Qt.RightToLeft)

    def _build_ui(self) -> None:
        central_widget = QWidget()
        central_widget.setLayoutDirection(Qt.RightToLeft)
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        _title_container = QWidget()
        _title_container.setLayoutDirection(Qt.LeftToRight)
        _title_row = QHBoxLayout(_title_container)
        _title_row.setContentsMargins(0, 0, 0, 0)
        _title_row.setSpacing(0)
        self.app_title_label = QLabel("כלי להערכת בקרות ITGC בסביבת SAP HANA APP")
        self.app_title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #16325c;")
        _title_row.addStretch(1)
        _title_row.addWidget(self.app_title_label)
        main_layout.addWidget(_title_container)

        _header_container = QWidget()
        _header_container.setLayoutDirection(Qt.LeftToRight)
        _header_row = QHBoxLayout(_header_container)
        _header_row.setContentsMargins(0, 0, 0, 0)
        _header_row.setSpacing(0)
        self.header_label = QLabel("מסך בדיקת קלטי SAP HANA DB")
        self.header_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #16325c;")
        _header_row.addStretch(1)
        _header_row.addWidget(self.header_label)

        self.hint_label = QTextEdit()
        self.hint_label.setReadOnly(True)
        self.hint_label.setHtml(
            '<p dir="rtl" style="color: #4f5d73; margin: 0; padding: 0;">'
            "בחר קבצים לפי המשבצת המתאימה. כוכבית מציינת משבצת חובה. חובה לציין את תאריך ההפקה של הקבצים. ניתן להריץ בדיקה נפרדת לכל קבוצת קבצים בלי להמתין לטעינת כל הדוחות."
            "</p>"
        )
        self.hint_label.setFixedHeight(46)
        self.hint_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.hint_label.setStyleSheet(
            "background: transparent; border: none; padding: 0;"
        )

        self.actions_row = QWidget()
        self.actions_row.setLayoutDirection(Qt.LeftToRight)
        self.actions_row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        buttons_layout = QHBoxLayout(self.actions_row)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        buttons_layout.addStretch(1)

        self.clear_last_load_button = QPushButton(self.format_ui_rtl_text("נקה טעינה אחרונה"))
        self.clear_last_load_button.clicked.connect(self.clear_last_loaded_slot)
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
        self.tabs.setLayoutDirection(Qt.RightToLeft)
        self.tabs.setDocumentMode(True)
        self.tabs.setTabPosition(QTabWidget.North)
        self.tabs.setMovable(False)
        self.tabs.setStyleSheet(
            """
            QTabBar::tab {
                background-color: #e9eef7;
                color: #16325c;
                border: 1px solid #b7c4d8;
                border-bottom: none;
                padding: 10px 18px;
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
        self.intake_tab.setLayoutDirection(Qt.RightToLeft)
        self.intake_layout = QVBoxLayout(self.intake_tab)
        self.intake_layout.setContentsMargins(8, 8, 8, 8)
        self.intake_layout.setSpacing(10)
        self.intake_layout.addWidget(_header_container)
        self.intake_layout.addWidget(self.hint_label)
        self.intake_layout.addWidget(self.actions_row)

        self.analysis_tab = QWidget()
        self.analysis_tab.setLayoutDirection(Qt.RightToLeft)
        self.analysis_layout = QVBoxLayout(self.analysis_tab)
        self.analysis_layout.setContentsMargins(12, 12, 12, 12)
        self.analysis_layout.setSpacing(10)
        self.analysis_hint_label = QLabel(
            self.format_ui_rtl_text("לאחר טעינת הקבצים ניתן לבצע ניתוח לביקורת ולסקור כאן את הממצאים המרכזיים.")
        )
        self.analysis_hint_label.setWordWrap(True)
        self.analysis_hint_label.setLayoutDirection(Qt.LeftToRight)
        self.analysis_hint_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.analysis_layout.addWidget(self.analysis_hint_label)
        self.audit_run_button = QPushButton(self.format_ui_rtl_text("בצע ניתוח לביקורת עבור המשבצת שנבחרה"))
        self.audit_run_button.clicked.connect(self.run_validation)
        self.analysis_layout.addWidget(self.audit_run_button, 0, Qt.AlignRight)

        self.review_tab = QWidget()
        self.review_tab.setLayoutDirection(Qt.RightToLeft)
        self.review_layout = QVBoxLayout(self.review_tab)
        self.review_layout.setContentsMargins(6, 6, 6, 6)
        self.review_layout.setSpacing(6)

        self.settings_tab = QWidget()
        self.settings_tab.setLayoutDirection(Qt.RightToLeft)
        self.settings_tab_layout = QVBoxLayout(self.settings_tab)
        self.settings_tab_layout.setContentsMargins(12, 12, 12, 12)
        self.settings_tab_layout.setSpacing(0)

        self.settings_scroll = QScrollArea()
        self.settings_scroll.setWidgetResizable(True)
        self.settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.settings_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.settings_scroll.setMinimumHeight(520)
        self.settings_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.settings_content = QWidget()
        self.settings_content.setLayoutDirection(Qt.RightToLeft)
        self.settings_layout = QVBoxLayout(self.settings_content)
        self.settings_layout.setContentsMargins(0, 0, 0, 0)
        self.settings_layout.setSpacing(10)

        self.settings_intro_label = QLabel(
            self.format_ui_rtl_text("בטאב זה ניתן לנהל את הגדרות הביקורת והעמודות הנדרשות לכל משבצת.")
        )
        self.settings_intro_label.setWordWrap(True)
        self.settings_intro_label.setLayoutDirection(Qt.LeftToRight)
        self.settings_intro_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.settings_layout.addWidget(self.settings_intro_label)
        self._build_system_settings_sections()

        self.settings_scroll.setWidget(self.settings_content)
        self.settings_tab_layout.addWidget(self.settings_scroll)

        self.tabs.addTab(self.intake_tab, self.format_rtl_text("קליטת קבצים"))
        self.tabs.addTab(self.settings_tab, self.format_rtl_text("הגדרות מערכת לביקורת"))
        self.tabs.addTab(self.review_tab, self.format_rtl_text("סקירת דוח משתמשים"))
        self.tabs.addTab(self.analysis_tab, self.format_rtl_text("ביצוע ניתוח לביקורת"))
        main_layout.addWidget(self.tabs)

        self.slots_group = QGroupBox(self.format_ui_rtl_text("מקורות קלט לבדיקת SAP HANA APP"))
        self.slots_group.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.slots_group.setLayoutDirection(Qt.RightToLeft)
        slots_group_layout = QVBoxLayout(self.slots_group)
        slots_group_layout.setContentsMargins(8, 18, 8, 8)

        self.slots_scroll = QScrollArea()
        self.slots_scroll.setWidgetResizable(True)
        self.slots_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.slots_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.slots_scroll.setMinimumHeight(280)
        self.slots_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        slots_container = QWidget()
        slots_container.setLayoutDirection(Qt.RightToLeft)
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
        slots_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)

        current_row = 0
        for category in self._ordered_categories():
            palette = self._category_palette(category)
            category_section = QGroupBox(self.format_ui_rtl_text(category))
            category_section.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            category_section.setStyleSheet(
                f"""
                QGroupBox {{
                    font-weight: bold;
                    border: 2px solid {palette['border']};
                    border-radius: 10px;
                    margin-top: 14px;
                    padding-top: 18px;
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
            self.category_sections[category] = category_section

            category_section.setLayoutDirection(Qt.RightToLeft)
            category_layout = QGridLayout(category_section)
            category_layout.setContentsMargins(12, 18, 12, 12)
            category_layout.setHorizontalSpacing(12)
            category_layout.setVerticalSpacing(10)
            category_layout.setColumnStretch(0, 0)
            category_layout.setColumnStretch(1, 1)
            category_layout.setColumnStretch(2, 2)
            category_layout.setColumnStretch(3, 0)
            category_layout.setColumnMinimumWidth(0, 140)
            category_layout.setColumnMinimumWidth(3, 120)
            category_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)

            category_button = QPushButton("הרץ בדיקה")
            category_button.setMinimumHeight(34)
            category_button.setToolTip(self.format_rtl_text(f"הרצת בדיקה עבור קבוצת {category}"))
            category_button.setStyleSheet(
                f"background-color: {palette['button']}; border: 2px solid {palette['border']}; color: white; font-weight: bold;"
            )
            category_button.clicked.connect(
                lambda _checked=False, cat=category: self.run_category_validation(cat)
            )
            self.category_run_buttons[category] = category_button

            section_row = 1
            for slot_key, metadata in self.SLOT_DEFINITIONS.items():
                if metadata["category"] != category:
                    continue

                display_name = metadata.get("label", slot_key)
                slot_title = QLabel(self.format_ui_rtl_text(f"{display_name}{' *' if metadata['required'] else ''}"))
                slot_title.setLayoutDirection(Qt.LeftToRight)
                slot_title.setAlignment(Qt.AlignRight | Qt.AlignTop)
                slot_title.setStyleSheet("font-weight: bold;")
                slot_title.setMinimumWidth(110)

                description = QLabel(self.format_ui_rtl_text(metadata["description"]))
                description.setLayoutDirection(Qt.LeftToRight)
                description.setAlignment(Qt.AlignRight | Qt.AlignTop)
                description.setWordWrap(True)
                description.setMinimumHeight(34)
                description.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

                sample = QLabel(self.format_ui_rtl_text(f"קובץ צפוי: {metadata['expected_file']}"))
                sample.setLayoutDirection(Qt.LeftToRight)
                sample.setAlignment(Qt.AlignRight | Qt.AlignTop)
                sample.setWordWrap(True)
                sample.setStyleSheet("color: #5b6573;")
                sample.setMinimumWidth(120)
                sample.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

                status_label = QLabel(self.format_ui_rtl_text("טרם נבחר קובץ"))
                status_label.setLayoutDirection(Qt.RightToLeft)
                status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                status_label.setWordWrap(True)
                status_label.setMinimumHeight(32)
                status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                status_label.setStyleSheet("padding: 6px; background: #ffffff; border: 1px solid #cfd6e4;")

                extraction_date_label = QLabel(self.format_ui_rtl_text("תאריך הפקה:"))
                extraction_date_label.setLayoutDirection(Qt.RightToLeft)
                extraction_date_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                extraction_date_label.setStyleSheet("color: #5b6573;")
                extraction_date_edit = QLineEdit(self._default_extraction_date())
                extraction_date_edit.setAlignment(Qt.AlignRight)
                extraction_date_edit.setPlaceholderText("YYYY-MM-DD")
                extraction_date_edit.setMinimumHeight(32)
                extraction_date_edit.setMaximumWidth(170)

                extraction_date_row = QWidget()
                extraction_date_row.setLayoutDirection(Qt.RightToLeft)
                extraction_date_layout = QHBoxLayout(extraction_date_row)
                extraction_date_layout.setContentsMargins(0, 0, 0, 0)
                extraction_date_layout.setSpacing(6)
                extraction_date_layout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                extraction_date_layout.addWidget(extraction_date_label, 0, Qt.AlignRight)
                extraction_date_layout.addWidget(extraction_date_edit, 0, Qt.AlignRight)
                extraction_date_layout.addStretch(1)

                select_button = QPushButton("בחירת קבצים" if slot_key in self.MULTI_FILE_SLOTS else "בחירת קובץ")
                select_button.setMinimumHeight(34)
                select_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                select_button.clicked.connect(lambda _checked=False, sk=slot_key: self.choose_file(sk))

                clear_slot_button = QPushButton("נקה")
                clear_slot_button.setMinimumHeight(34)
                clear_slot_button.setMinimumWidth(74)
                clear_slot_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                clear_slot_button.clicked.connect(lambda _checked=False, sk=slot_key: self.clear_slot_selection(sk))

                slot_buttons = QWidget()
                slot_buttons.setLayoutDirection(Qt.RightToLeft)
                slot_buttons_layout = QHBoxLayout(slot_buttons)
                slot_buttons_layout.setContentsMargins(0, 0, 0, 0)
                slot_buttons_layout.setSpacing(6)
                slot_buttons_layout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
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

            # Add the 'הרץ בדיקה' button at the last row, spanning all columns
            category_layout.setRowMinimumHeight(section_row, 40)
            category_layout.addWidget(category_button, section_row, 0, 1, 4, alignment=Qt.AlignRight)
            section_row += 1
            category_section.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            slots_layout.addWidget(category_section, current_row, 0, 1, 4)
            current_row += 1

        bottom_spacer = QLabel("")
        bottom_spacer.setMinimumHeight(120)
        slots_layout.addWidget(bottom_spacer, current_row, 0, 1, 4)
        current_row += 1
        slots_layout.setRowStretch(current_row, 1)
        slots_layout.setRowMinimumHeight(current_row, 20)
        self.slots_scroll.setWidget(slots_container)
        slots_group_layout.addWidget(self.slots_scroll)

        self.slots_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.intake_layout.addWidget(self.slots_group, 1)

        self.user_preview_group = QGroupBox(self.format_ui_rtl_text("רשימת משתמשים שנטענו"))
        self.user_preview_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.user_preview_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.user_preview_group.setMinimumHeight(460)
        user_preview_layout = QVBoxLayout(self.user_preview_group)
        user_preview_layout.setContentsMargins(8, 12, 8, 8)
        user_preview_layout.setSpacing(4)
        user_preview_layout.setAlignment(Qt.AlignTop)

        self.user_preview_actions_row = QWidget()
        self.user_preview_actions_row.setLayoutDirection(Qt.RightToLeft)
        self.user_preview_actions_row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.user_preview_actions_row.setMaximumHeight(40)
        user_preview_actions_layout = QHBoxLayout(self.user_preview_actions_row)
        user_preview_actions_layout.setContentsMargins(0, 0, 0, 0)
        user_preview_actions_layout.setSpacing(8)
        user_preview_actions_layout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.user_preview_export_button = QPushButton("ייצוא סקירה לאקסל")
        self.user_preview_export_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.user_preview_export_button.clicked.connect(lambda: self.export_user_preview_to_excel(open_after_export=True))
        user_preview_actions_layout.addWidget(self.user_preview_export_button, 0, Qt.AlignRight)

        self.user_preview_import_button = QPushButton("ייבוא סקירה מאקסל")
        self.user_preview_import_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.user_preview_import_button.clicked.connect(self.import_user_review_from_excel)
        user_preview_actions_layout.addWidget(self.user_preview_import_button, 0, Qt.AlignRight)

        self.user_preview_columns_button = QPushButton("הוסף / מחק עמודות")
        self.user_preview_columns_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.user_preview_columns_button.clicked.connect(self.show_user_preview_column_dialog)
        user_preview_actions_layout.addWidget(self.user_preview_columns_button, 0, Qt.AlignRight)
        user_preview_actions_layout.addStretch(1)
        user_preview_layout.addWidget(self.user_preview_actions_row, 0, Qt.AlignTop)

        self.user_preview_filter_row = QWidget()
        self.user_preview_filter_row.setLayoutDirection(Qt.RightToLeft)
        self.user_preview_filter_row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        user_preview_filter_layout = QHBoxLayout(self.user_preview_filter_row)
        user_preview_filter_layout.setContentsMargins(0, 0, 0, 0)
        user_preview_filter_layout.setSpacing(8)
        user_preview_filter_layout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.user_preview_filter_label = QLabel(self.format_ui_rtl_text("סינון משתמשים:"))
        self.user_preview_status_filter = QComboBox()
        self.user_preview_status_filter.setMinimumWidth(220)
        self.user_preview_status_filter.setLayoutDirection(Qt.RightToLeft)
        for filter_value, filter_label in self.USER_PREVIEW_FILTER_OPTIONS:
            self.user_preview_status_filter.addItem(self.format_rtl_text(filter_label), filter_value)

        self.audit_period_from_label = QLabel(self.format_ui_rtl_text("מתאריך:"))
        self.audit_period_from_edit = QLineEdit("")
        self.audit_period_from_edit.setAlignment(Qt.AlignRight)
        self.audit_period_from_edit.setPlaceholderText("YYYY-MM-DD")
        self.audit_period_from_edit.setMaximumWidth(130)

        self.audit_period_to_label = QLabel(self.format_ui_rtl_text("עד תאריך:"))
        self.audit_period_to_edit = QLineEdit("")
        self.audit_period_to_edit.setAlignment(Qt.AlignRight)
        self.audit_period_to_edit.setPlaceholderText("YYYY-MM-DD")
        self.audit_period_to_edit.setMaximumWidth(130)

        user_preview_filter_layout.addWidget(self.user_preview_filter_label, 0, Qt.AlignRight)
        user_preview_filter_layout.addWidget(self.user_preview_status_filter, 0, Qt.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_from_label, 0, Qt.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_from_edit, 0, Qt.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_to_label, 0, Qt.AlignRight)
        user_preview_filter_layout.addWidget(self.audit_period_to_edit, 0, Qt.AlignRight)
        user_preview_filter_layout.addStretch(1)
        user_preview_layout.addWidget(self.user_preview_filter_row, 0, Qt.AlignTop)

        self.user_preview_hint = QLabel(
            self.format_ui_rtl_text(
                "הטבלה מציגה את משתמשי USR02 עם העשרת נתונים מקובצי USER_ADDR ו-ADR6 בלבד."
            )
        )
        self.user_preview_hint.setWordWrap(True)
        self.user_preview_hint.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.user_preview_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.user_preview_hint.setMaximumHeight(44)
        user_preview_layout.addWidget(self.user_preview_hint, 0, Qt.AlignTop)

        self.user_review_progress_group = QGroupBox(self.format_ui_rtl_text("סיכום התקדמות סקירה"))
        self.user_review_progress_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.user_review_progress_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        user_review_progress_layout = QVBoxLayout(self.user_review_progress_group)
        user_review_progress_layout.setContentsMargins(8, 8, 8, 8)
        user_review_progress_layout.setSpacing(6)

        user_review_counts_row = QWidget()
        user_review_counts_row.setLayoutDirection(Qt.RightToLeft)
        user_review_counts_layout = QHBoxLayout(user_review_counts_row)
        user_review_counts_layout.setContentsMargins(0, 0, 0, 0)
        user_review_counts_layout.setSpacing(14)
        user_review_counts_layout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.user_review_total_label = QLabel(self.format_ui_rtl_text("סה\"כ משתמשים בדוח: 0"))
        self.user_review_total_label.setStyleSheet("font-weight: bold;")
        self.user_review_reviewed_label = QLabel(self.format_ui_rtl_text("משתמשים שנבדקו: 0"))
        self.user_review_reviewed_label.setStyleSheet("font-weight: bold; color: #2e7d32;")
        self.user_review_unreviewed_label = QLabel(self.format_ui_rtl_text("משתמשים שטרם נבדקו: 0"))
        self.user_review_unreviewed_label.setStyleSheet("font-weight: bold; color: #1565c0;")

        user_review_counts_layout.addWidget(self.user_review_total_label, 0, Qt.AlignRight)
        user_review_counts_layout.addWidget(self.user_review_reviewed_label, 0, Qt.AlignRight)
        user_review_counts_layout.addWidget(self.user_review_unreviewed_label, 0, Qt.AlignRight)
        user_review_counts_layout.addStretch(1)
        user_review_progress_layout.addWidget(user_review_counts_row)

        self.user_review_progress_bar = QProgressBar()
        self.user_review_progress_bar.setMinimum(0)
        self.user_review_progress_bar.setMaximum(100)
        self.user_review_progress_bar.setValue(0)
        self.user_review_progress_bar.setTextVisible(True)
        self.user_review_progress_bar.setAlignment(Qt.AlignCenter)
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
        self.user_review_progress_percent_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.user_review_progress_percent_label.setStyleSheet("font-weight: bold; color: #0d47a1;")
        user_review_progress_layout.addWidget(self.user_review_progress_percent_label)

        user_preview_layout.addWidget(self.user_review_progress_group, 0, Qt.AlignTop)

        self.user_preview_table = QTableWidget(0, 0)
        self.user_preview_table.setEditTriggers(
            QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed | QTableWidget.SelectedClicked
        )
        self.user_preview_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.user_preview_table.setAlternatingRowColors(True)
        self.user_preview_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.user_preview_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.user_preview_table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.user_preview_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.user_preview_table.setMinimumHeight(420)
        self.user_preview_table.setMaximumHeight(16777215)
        self._configure_user_preview_table()
        self.user_preview_table.itemChanged.connect(self._handle_user_preview_item_changed)
        self.user_preview_status_filter.currentIndexChanged.connect(self.refresh_user_preview)
        self.audit_period_from_edit.editingFinished.connect(self.refresh_user_preview)
        self.audit_period_to_edit.editingFinished.connect(self.refresh_user_preview)
        user_preview_layout.addWidget(self.user_preview_table, 1)
        self.review_layout.addWidget(self.user_preview_group, 1)

        self.run_log_group = QGroupBox(self.format_ui_rtl_text("לוג קבצים שנבדקו"))
        self.run_log_group.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        run_log_layout = QVBoxLayout(self.run_log_group)
        run_log_layout.setContentsMargins(12, 18, 12, 12)
        self.run_log_table = QTableWidget(0, 10)
        self.run_log_table.setHorizontalHeaderLabels(["משבצת", "קבוצת דוחות", "קובץ", "תאריך הפקה", "רשומות שנקלטו", "סטטוס", "מספר שגיאות", "תיאור שגיאה", "תאריך בדיקה", "שעת בדיקה"])
        self.run_log_table.horizontalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.run_log_table.verticalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.run_log_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.run_log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        self.run_log_table.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        self.run_log_table.setColumnWidth(1, 150)
        self.run_log_table.setColumnWidth(2, 180)
        self.run_log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.run_log_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.run_log_table.setAlternatingRowColors(True)
        self.run_log_table.setWordWrap(True)
        self.run_log_table.setTextElideMode(Qt.ElideMiddle)
        self.run_log_table.setMinimumHeight(160)
        self.run_log_table.setToolTip("לחיצה כפולה על שורה תפתח פירוט מלא עבור הקובץ")
        self.run_log_table.cellDoubleClicked.connect(self.show_log_details)
        run_log_layout.addWidget(self.run_log_table)
        self.run_log_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.run_log_group.setMaximumHeight(320)
        self.intake_layout.addWidget(self.run_log_group, 0)

        self.required_columns_group = QGroupBox(self.format_ui_rtl_text("עמודות חובה לבדיקה"))
        self.required_columns_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        required_layout = QHBoxLayout(self.required_columns_group)
        self.required_columns_edit = QLineEdit("")
        self.required_columns_edit.setAlignment(Qt.AlignRight)
        self.required_columns_edit.setPlaceholderText("יוזן אוטומטית לפי המשבצת שנבחרה")
        required_layout.addWidget(self.required_columns_edit)
        self.required_columns_group.show()
        self.settings_layout.addWidget(self.required_columns_group)

        self.summary_group = QGroupBox(self.format_ui_rtl_text("סיכום בדיקה"))
        self.summary_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
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
            title_label.setAlignment(Qt.AlignCenter)
            title_label.setStyleSheet("font-weight: bold;")
            value_label = QLabel(default_value)
            value_label.setAlignment(Qt.AlignCenter)
            value_label.setStyleSheet("font-size: 18px; padding: 6px;")
            summary_layout.addWidget(title_label, 0, column)
            summary_layout.addWidget(value_label, 1, column)
            self.summary_labels[key] = value_label
        self.summary_group.hide()
        self.analysis_layout.addWidget(self.summary_group)

        self.results_group = QGroupBox(self.format_ui_rtl_text("רשימת שגיאות"))
        self.results_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        results_layout = QVBoxLayout(self.results_group)
        results_layout.setContentsMargins(12, 18, 12, 12)
        self.issues_table = QTableWidget(0, 3)
        self.issues_table.setHorizontalHeaderLabels(["מספר שורה", "שם עמודה", "הודעת שגיאה"])
        self.issues_table.horizontalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.issues_table.verticalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.issues_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.issues_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.issues_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.issues_table.setAlternatingRowColors(True)
        results_layout.addWidget(self.issues_table)
        self.issues_table.setMinimumHeight(180)
        self.results_group.hide()
        self.analysis_layout.addWidget(self.results_group)

        central_widget.setStyleSheet(
            """
            QWidget {
                background-color: #f5f7fb;
                font-family: 'Segoe UI';
                font-size: 11pt;
            }
            QLabel {
                qproperty-alignment: 'AlignRight|AlignVCenter';
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #c7cfda;
                border-radius: 8px;
                margin-top: 16px;
                padding-top: 16px;
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
                padding: 8px 14px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #dbe7f8;
            }
            QLineEdit {
                background-color: white;
                border: 1px solid #cfd6e4;
                padding: 6px;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #cfd6e4;
                gridline-color: #d7deea;
            }
            """
        )

    def _ordered_categories(self) -> list[str]:
        categories: list[str] = []
        for metadata in self.SLOT_DEFINITIONS.values():
            if metadata["category"] not in categories:
                categories.append(metadata["category"])
        return categories

    def _category_palette(self, category: str) -> dict[str, str]:
        # Unified palette: active tab color (burgundy)
        return {"header": "#6d002f", "button": "#6d002f", "border": "#6d002f"}

    @staticmethod
    def _default_extraction_date() -> str:
        return datetime.now().strftime("%Y-%m-%d")

    def _build_system_settings_sections(self) -> None:
        buttons_row = QWidget()
        buttons_layout = QHBoxLayout(buttons_row)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        buttons_layout.addStretch(1)

        self.settings_reset_btn = QPushButton(self.format_ui_rtl_text("טען ברירות מחדל"))
        self.settings_reset_btn.clicked.connect(self._reset_system_settings_form)
        self.settings_save_btn = QPushButton(self.format_ui_rtl_text("שמור הגדרות"))
        self.settings_save_btn.clicked.connect(self._save_system_settings)
        buttons_layout.addWidget(self.settings_save_btn)
        buttons_layout.addWidget(self.settings_reset_btn)
        self.settings_layout.addWidget(buttons_row)

        review_group, review_layout, review_unavailable_label = self._build_settings_group(
            "טווח בחינה לסקירת משתמשים",
            "הגדרה מרכזית של טווח תקופת הבחינה, מסונכרנת עם מסך סקירת המשתמשים.",
        )
        review_form = QFormLayout()
        review_form.setLabelAlignment(Qt.AlignRight)
        review_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

        start_widget = QDateEdit()
        start_widget.setLayoutDirection(Qt.LeftToRight)
        start_widget.setDisplayFormat("yyyy-MM-dd")
        start_widget.setCalendarPopup(True)
        end_widget = QDateEdit()
        end_widget.setLayoutDirection(Qt.LeftToRight)
        end_widget.setDisplayFormat("yyyy-MM-dd")
        end_widget.setCalendarPopup(True)

        self.system_settings_widgets["user_review_period.start_date"] = start_widget
        self.system_settings_widgets["user_review_period.end_date"] = end_widget
        review_form.addRow("מתאריך", start_widget)
        review_form.addRow("עד תאריך", end_widget)
        review_layout.addLayout(review_form)

        self.settings_layout.addWidget(review_group)
        self.system_settings_sections["user_review_period"] = review_group
        self.system_settings_unavailable_labels["user_review_period"] = review_unavailable_label

        super_users_group, super_users_table, super_users_unavailable_label = self._build_super_users_section()
        self.settings_layout.addWidget(super_users_group)
        self.system_settings_widgets["super_users"] = super_users_table
        self.system_settings_sections["super_users"] = super_users_group
        self.system_settings_unavailable_labels["super_users"] = super_users_unavailable_label

        generic_users_group = self._add_settings_text_list_section("generic_users", "משתמשים גנריים", "רשימה מופרדת שורות")
        self.system_settings_sections["generic_users"] = generic_users_group

        self._add_settings_text_list_section("critical_roles", "פרופילים משתמשיי על", "רשימה מופרדת שורות")
        self._add_settings_text_list_section("critical_privileges", "הרשאות על", "רשימה מופרדת שורות")

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
                    widget.setLayoutDirection(Qt.LeftToRight)
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

        mapping_group, mapping_layout, mapping_unavailable_label = self._build_settings_group(
            "מיפוי קבצים",
            "התאמת שמות קבצים צפויים לכל משבצת, עבור וריאציות ייצוא בין כלים.",
        )
        mapping_form = QFormLayout()
        mapping_form.setLabelAlignment(Qt.AlignRight)
        mapping_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.system_settings_file_mapping_order = list(self.SLOT_DEFINITIONS.keys())
        for slot_key in self.system_settings_file_mapping_order:
            metadata = self.SLOT_DEFINITIONS[slot_key]
            display_name = str(metadata.get("label", slot_key))
            field_widget = QLineEdit()
            field_widget.setLayoutDirection(Qt.LeftToRight)
            self.system_settings_widgets[f"file_mappings.{slot_key}"] = field_widget
            mapping_form.addRow(self.format_ui_rtl_text(display_name), field_widget)
        mapping_layout.addLayout(mapping_form)
        self.settings_layout.addWidget(mapping_group)
        self.system_settings_sections["file_mappings"] = mapping_group
        self.system_settings_unavailable_labels["file_mappings"] = mapping_unavailable_label

        threshold_group, threshold_layout, threshold_unavailable_label = self._build_settings_group(
            "הגדרות נוספות",
            "סף חוסר פעילות משמש לבניית ממצאים אוטומטיים בסקירת משתמשים.",
        )
        threshold_form = QFormLayout()
        threshold_form.setLabelAlignment(Qt.AlignRight)
        threshold_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        threshold_widget = QLineEdit()
        self.system_settings_widgets["inactive_days_threshold"] = threshold_widget
        threshold_form.addRow("סף חוסר פעילות (ימים)", threshold_widget)
        threshold_layout.addLayout(threshold_form)
        self.settings_layout.addWidget(threshold_group)
        self.system_settings_sections["inactive_days_threshold"] = threshold_group
        self.system_settings_unavailable_labels["inactive_days_threshold"] = threshold_unavailable_label

    def _build_settings_group(self, title: str, description: str | None = None) -> tuple[QGroupBox, QVBoxLayout, QLabel]:
        group = QGroupBox(self.format_ui_rtl_text(title))
        group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 14, 10, 10)
        layout.setSpacing(8)
        if description:
            description_label = QLabel(self.format_ui_rtl_text(description))
            description_label.setWordWrap(True)
            description_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
            layout.addWidget(description_label)

        unavailable_label = QLabel(self.format_ui_rtl_text("הגדרה זו לא זמינה ללא קובץ מקור רלוונטי"))
        unavailable_label.setWordWrap(True)
        unavailable_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        unavailable_label.setStyleSheet("color: gray; font-style: italic;")
        unavailable_label.setVisible(False)
        layout.addWidget(unavailable_label)
        return group, layout, unavailable_label

    def _build_super_users_section(self) -> tuple[QGroupBox, QTableWidget, QLabel]:
        group, layout, unavailable_label = self._build_settings_group(
            "משתמשיי על",
            "רשימה של משתמשים בעלי גישה גבוהה. יש להזין CLIENT ו-BNAME.",
        )
        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels(["CLIENT", "משתמש"])
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
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

    def _add_settings_text_list_section(self, key: str, title: str, description: str) -> QGroupBox:
        group, group_layout, unavailable_label = self._build_settings_group(title, description)
        editor = QPlainTextEdit()
        editor.setMinimumHeight(90)
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

    def _default_system_settings(self) -> dict[str, object]:
        return {
            "generic_users": ["SAP", "DDIC", "TMSADM", "SAPCPIC"],
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

    def _current_system_settings(self) -> dict[str, object]:
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
        merged = copy.deepcopy(defaults)
        for key, value in loaded.items():
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key].update(value)
            else:
                merged[key] = value
        return merged

    def _sync_review_filters_from_settings(self, settings: dict[str, object]) -> None:
        period_cfg = settings.get("user_review_period", {}) if isinstance(settings, dict) else {}
        start_text = str(period_cfg.get("start_date", "")).strip()
        end_text = str(period_cfg.get("end_date", "")).strip()
        if hasattr(self, "audit_period_from_edit") and start_text:
            self.audit_period_from_edit.setText(start_text)
        if hasattr(self, "audit_period_to_edit") and end_text:
            self.audit_period_to_edit.setText(end_text)

    def _load_system_settings_into_form(self, settings: dict[str, object], load_review_period: bool = True) -> None:
        settings = settings or self._default_system_settings()

        def _fill_lines(key: str) -> None:
            editor = self.system_settings_widgets.get(key)
            values = settings.get(key, [])
            if isinstance(editor, QPlainTextEdit):
                editor.setPlainText("\n".join(str(item).strip() for item in values if str(item).strip()))

        _fill_lines("generic_users")
        _fill_lines("critical_roles")
        _fill_lines("critical_privileges")

        super_users_table = self.system_settings_widgets.get("super_users")
        super_users = settings.get("super_users", []) if isinstance(settings, dict) else []
        if isinstance(super_users_table, QTableWidget):
            super_users_table.setRowCount(0)
            if isinstance(super_users, list):
                for super_user in super_users:
                    if isinstance(super_user, dict):
                        mandt = str(super_user.get("MANDT", "")).strip()
                        bname = str(super_user.get("BNAME", "")).strip()
                        if mandt or bname:
                            row = super_users_table.rowCount()
                            super_users_table.insertRow(row)
                            super_users_table.setItem(row, 0, QTableWidgetItem(mandt))
                            super_users_table.setItem(row, 1, QTableWidgetItem(bname))

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

        password_defaults = settings.get("password_policy_defaults", {}) if isinstance(settings, dict) else {}
        if isinstance(password_defaults, dict):
            for key, value in password_defaults.items():
                widget = self.system_settings_widgets.get(f"password_policy_defaults.{key}")
                if isinstance(widget, QComboBox):
                    widget.setCurrentText(str(value))
                elif isinstance(widget, QLineEdit):
                    widget.setText(str(value))

    def _collect_system_settings_from_form(self) -> dict[str, object]:
        def _lines_from_editor(editor: object) -> list[str]:
            if not isinstance(editor, QPlainTextEdit):
                return []
            return [line.strip() for line in editor.toPlainText().splitlines() if line.strip()]

        settings = self._default_system_settings()
        settings["generic_users"] = _lines_from_editor(self.system_settings_widgets.get("generic_users"))
        settings["critical_roles"] = _lines_from_editor(self.system_settings_widgets.get("critical_roles"))
        settings["critical_privileges"] = _lines_from_editor(self.system_settings_widgets.get("critical_privileges"))

        super_users_table = self.system_settings_widgets.get("super_users")
        super_users: list[dict[str, str]] = []
        if isinstance(super_users_table, QTableWidget):
            for row_index in range(super_users_table.rowCount()):
                mandt_item = super_users_table.item(row_index, 0)
                bname_item = super_users_table.item(row_index, 1)
                mandt_text = str(mandt_item.text()).strip() if isinstance(mandt_item, QTableWidgetItem) else ""
                bname_text = str(bname_item.text()).strip() if isinstance(bname_item, QTableWidgetItem) else ""
                if mandt_text or bname_text:
                    super_users.append({"MANDT": mandt_text, "BNAME": bname_text})
        settings["super_users"] = super_users

        period_start_widget = self.system_settings_widgets.get("user_review_period.start_date")
        period_end_widget = self.system_settings_widgets.get("user_review_period.end_date")
        if isinstance(period_start_widget, QDateEdit) and isinstance(period_end_widget, QDateEdit):
            settings["user_review_period"] = {
                "start_date": period_start_widget.date().toPython().isoformat(),
                "end_date": period_end_widget.date().toPython().isoformat(),
            }

        file_mappings = {}
        for mapping_key in self.system_settings_file_mapping_order:
            widget = self.system_settings_widgets.get(f"file_mappings.{mapping_key}")
            if isinstance(widget, QLineEdit):
                file_mappings[mapping_key] = widget.text().strip()
        settings["file_mappings"] = file_mappings

        threshold_widget = self.system_settings_widgets.get("inactive_days_threshold")
        if isinstance(threshold_widget, QLineEdit):
            settings["inactive_days_threshold"] = self._safe_int(threshold_widget.text(), 90)

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
                default_value = self._default_system_settings()["password_policy_defaults"].get(field_name, 0)
                password_defaults[field_name] = self._safe_int(widget.text(), int(default_value))
        settings["password_policy_defaults"] = password_defaults

        return settings

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
        available_slots = self._available_selected_slots()
        for section_key, section_widget in self.system_settings_sections.items():
            required_slots = self.SETTINGS_SECTION_DEPENDENCIES.get(section_key, set())
            is_available = not required_slots or bool(available_slots.intersection(required_slots))
            section_widget.setVisible(True)
            section_widget.setEnabled(is_available)
            if section_key in self.system_settings_unavailable_labels:
                self.system_settings_unavailable_labels[section_key].setVisible(not is_available)
            section_widget.setToolTip(
                "" if is_available else "לא זמין ללא קובץ מקור רלוונטי עבור התיבה הזו"
            )

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

    def _get_slot_category(self, slot_key: str) -> str:
        return str(self.SLOT_DEFINITIONS.get(slot_key, {}).get("category", "לא סווג"))

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

    def _is_generic_user(self, bname: object, settings: dict[str, object]) -> bool:
        if not bname or not isinstance(settings, dict):
            return False
        normalized_name = str(bname).strip().casefold()
        generic_users = settings.get("generic_users", [])
        if not isinstance(generic_users, list):
            return False
        return any(str(item).strip().casefold() == normalized_name for item in generic_users)

    def _is_super_user(self, mandt: object, bname: object, settings: dict[str, object]) -> bool:
        if not bname or not isinstance(settings, dict):
            return False
        normalized_mandt = str(mandt).strip()
        normalized_bname = str(bname).strip().casefold()
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
        is_locked = preview_row.get("STATUS", "") == "נעול"
        has_findings = bool(preview_row.get("FINDINGS_DESCRIPTION", "").strip())
        is_reviewed = preview_row.get("REVIEW_STATUS", "") in ValidationDesktopApp.REVIEWED_STATUSES
        is_not_ok = preview_row.get("REVIEW_STATUS", "") == "נבדק - לא תקין"
        has_notes = bool(preview_row.get("REVIEW_NOTES", "").strip())
        if not is_locked and has_findings and (not is_reviewed or (is_not_ok and not has_notes)):
            return 1
        if not is_locked and not is_reviewed:
            return 2
        if not is_locked and is_reviewed and has_notes:
            return 3
        if not is_locked and is_reviewed and not has_findings and not has_notes:
            return 4
        return 5

    def _update_review_row_highlight(self, row_index: int) -> None:
        review_status_col: int | None = None
        notes_col: int | None = None
        for col_idx, field_name in enumerate(self.user_preview_visible_columns):
            if field_name == "REVIEW_STATUS":
                review_status_col = col_idx
            elif field_name == "REVIEW_NOTES":
                notes_col = col_idx

        review_status_text = ""
        if review_status_col is not None:
            combo = self.user_preview_table.cellWidget(row_index, review_status_col)
            if isinstance(combo, QComboBox):
                review_status_text = self.format_rtl_text(combo.currentText())

        notes_text = ""
        if notes_col is not None:
            notes_item = self.user_preview_table.item(row_index, notes_col)
            if notes_item is not None:
                notes_text = notes_item.text().strip()

        is_not_reviewed = review_status_text == "טרם נבדק"
        is_not_ok = review_status_text == "נבדק - לא תקין"
        needs_warning = is_not_ok and not notes_text

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

            if needs_warning and field_name in {"REVIEW_STATUS", "REVIEW_NOTES"}:
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
        total_rows = self.user_preview_table.rowCount()
        reviewed_rows = 0
        for row_index in range(total_rows):
            if self._get_user_preview_row_review_status(row_index) in self.REVIEWED_STATUSES:
                reviewed_rows += 1
        self._update_user_review_progress_summary(total_rows, reviewed_rows, total_rows - reviewed_rows)

    def _update_slot_path_label(self, slot_key: str, file_paths: list[str] | None = None) -> None:
        widget_data = self.slot_widgets.get(slot_key, {})
        label = widget_data.get("path_label")
        if not isinstance(label, QLabel):
            return

        paths = file_paths if file_paths is not None else list(widget_data.get("selected_paths", []))
        if not paths:
            label.setLayoutDirection(Qt.RightToLeft)
            label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            label.setText(self.format_ui_rtl_text("טרם נבחר קובץ"))
            return

        if len(paths) == 1:
            label.setLayoutDirection(Qt.LeftToRight)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setText(self.format_rtl_text(paths[0]))
            return

        label.setLayoutDirection(Qt.RightToLeft)
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
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

    def _get_category_slots(self, category: str) -> list[str]:
        return [
            slot_key
            for slot_key, metadata in self.SLOT_DEFINITIONS.items()
            if metadata["category"] == category
        ]

    def _current_file_paths(self) -> list[str]:
        if not self.selected_slot_key:
            return []
        return list(self.slot_widgets[self.selected_slot_key].get("selected_paths", []))

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
            "REVIEW_NOTES": "",
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
            normalized_state[str(review_key)] = {
                "REVIEW_STATUS": self._normalize_reviewer_status(review_values.get("REVIEW_STATUS")),
                "REVIEW_NOTES": str(review_values.get("REVIEW_NOTES", "")).strip(),
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
        return {
            "REVIEW_STATUS": self._normalize_reviewer_status(stored_values.get("REVIEW_STATUS")),
            "REVIEW_NOTES": str(stored_values.get("REVIEW_NOTES", "")).strip(),
        }

    def _update_reviewer_value(self, review_key: str, field_name: str, value: object) -> None:
        if not review_key or field_name not in {"REVIEW_STATUS", "REVIEW_NOTES"}:
            return

        current_values = self.user_reviewer_state.setdefault(review_key, self._default_reviewer_values().copy())
        if field_name == "REVIEW_STATUS":
            current_values[field_name] = self._normalize_reviewer_status(value)
        else:
            current_values[field_name] = "" if value is None else str(value).strip()
        self._save_user_reviewer_state()

    def _normalize_user_preview_columns(self, selected_columns: list[str] | None) -> list[str]:
        allowed_fields = [column["field"] for column in self.USER_PREVIEW_COLUMN_DEFINITIONS]
        if not selected_columns:
            return list(self.DEFAULT_USER_PREVIEW_COLUMNS)

        normalized = [field for field in allowed_fields if field in selected_columns]
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

    def _get_user_preview_column_definition(self, field_name: str) -> dict[str, object]:
        for column in self.USER_PREVIEW_COLUMN_DEFINITIONS:
            if column["field"] == field_name:
                return column
        return {"field": field_name, "formal": field_name, "technical": field_name, "source": "לא ידוע", "width": 120}

    def _handle_user_preview_item_changed(self, item: QTableWidgetItem) -> None:
        if self._refreshing_user_preview or item is None:
            return

        field_name = item.data(Qt.UserRole + 1)
        review_key = item.data(Qt.UserRole)
        if field_name != "REVIEW_NOTES" or not review_key:
            return

        normalized_text = self.format_rtl_text(item.text())
        item.setToolTip(normalized_text)
        self._update_reviewer_value(str(review_key), "REVIEW_NOTES", normalized_text)
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
        header.setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header.setSectionsMovable(False)
        header.setSectionsClickable(True)
        header.setMinimumSectionSize(70)
        self.user_preview_table.verticalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        for column_index, field_name in enumerate(self.user_preview_visible_columns):
            header.setSectionResizeMode(column_index, QHeaderView.Interactive)
            default_width = int(self._get_user_preview_column_definition(field_name).get("width", 120))
            self.user_preview_table.setColumnWidth(column_index, default_width)
        self.user_preview_table.setSortingEnabled(True)

    def _create_user_preview_columns_dialog(self) -> tuple[QDialog, QTableWidget]:
        dialog = QDialog(self)
        dialog.setWindowTitle(self.format_rtl_text("בחירת עמודות לסקירת משתמשים"))
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(720, 460)

        layout = QVBoxLayout(dialog)
        hint_label = QLabel(
            self.format_ui_rtl_text("סמן את העמודות שברצונך להציג. לחיצה על OK תרענן את הטבלה, ו-Cancel תשאיר את המצב הקיים.")
        )
        hint_label.setWordWrap(True)
        hint_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        layout.addWidget(hint_label)

        selection_table = QTableWidget(len(self.USER_PREVIEW_COLUMN_DEFINITIONS), 3)
        selection_table.setHorizontalHeaderLabels(["שם פורמלי", "שם טכני", "הצג"])
        selection_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        selection_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        selection_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        selection_table.horizontalHeader().setDefaultAlignment(Qt.AlignRight | Qt.AlignVCenter)
        selection_table.verticalHeader().setVisible(False)
        selection_table.setAlternatingRowColors(True)
        selection_table.setSelectionBehavior(QTableWidget.SelectRows)

        for row_index, column in enumerate(self.USER_PREVIEW_COLUMN_DEFINITIONS):
            formal_item = QTableWidgetItem(self.format_rtl_text(str(column["formal"])))
            formal_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            formal_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            technical_item = QTableWidgetItem(str(column["technical"]))
            technical_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            technical_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            checkbox_item = QTableWidgetItem("")
            checkbox_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            checkbox_item.setCheckState(Qt.Checked if column["field"] in self.user_preview_visible_columns else Qt.Unchecked)
            checkbox_item.setTextAlignment(Qt.AlignCenter)
            checkbox_item.setToolTip(self.format_ui_rtl_text(f"מקור נתון: {column['source']}"))
            selection_table.setItem(row_index, 0, formal_item)
            selection_table.setItem(row_index, 1, technical_item)
            selection_table.setItem(row_index, 2, checkbox_item)

        layout.addWidget(selection_table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        return dialog, selection_table

    def _get_selected_user_preview_columns(self, selection_table: QTableWidget) -> list[str]:
        selected_columns: list[str] = []
        for row_index, column in enumerate(self.USER_PREVIEW_COLUMN_DEFINITIONS):
            checkbox_item = selection_table.item(row_index, 2)
            if checkbox_item is not None and checkbox_item.checkState() == Qt.Checked:
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
        if dialog.exec() != QDialog.Accepted:
            return

        selected_columns = self._get_selected_user_preview_columns(selection_table)
        if not selected_columns:
            QMessageBox.warning(self, "בחירת עמודות", "יש לבחור לפחות עמודה אחת להצגה בטבלת הסקירה.")
            return

        self._apply_user_preview_columns(selected_columns)

    @staticmethod
    def _get_row_value(row: dict[str, object], *candidates: str) -> str:
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

    def _load_preview_rows(self, slot_key: str) -> list[dict[str, object]]:
        file_paths = list(self.slot_widgets.get(slot_key, {}).get("selected_paths", []))
        if not file_paths:
            return []

        try:
            result = process_file(
                file_paths,
                required_columns=[],
                source_name_override=slot_key,
            )
        except Exception:
            return []

        return list(result.rows)

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
        usr02_rows: list[dict[str, object]],
        combined_rows: list[dict[str, object]],
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
                    "REVIEW_NOTES": review_values.get("REVIEW_NOTES", ""),
                }
            )

        return preview_rows

    def refresh_user_preview(self) -> None:
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
                        combo_box = QComboBox()
                        combo_box.setLayoutDirection(Qt.RightToLeft)
                        combo_box.setMinimumWidth(150)
                        for status_value in self.REVIEW_STATUS_OPTIONS:
                            combo_box.addItem(self.format_rtl_text(status_value))
                        combo_box.setCurrentText(self._normalize_reviewer_status(value))
                        combo_box.currentTextChanged.connect(
                            lambda selected_text, current_key=review_key, r_idx=row_index: (
                                self._update_reviewer_value(current_key, "REVIEW_STATUS", selected_text),
                                self._update_review_row_highlight(r_idx) if not self._refreshing_user_preview else None,
                                self._refresh_user_review_progress_summary_from_table() if not self._refreshing_user_preview else None,
                            )
                        )
                        self.user_preview_table.setCellWidget(row_index, column, combo_box)
                        continue

                    display_value = self._format_user_preview_value_for_display(field_name, value)
                    item = SortableTableWidgetItem(self.format_rtl_text(display_value))
                    item.setData(SortableTableWidgetItem.SORT_ROLE, self._get_user_preview_sort_value(field_name, value))
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    item.setToolTip(self.format_rtl_text(display_value))
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

                    if field_name == "REVIEW_NOTES":
                        item.setFlags(item.flags() | Qt.ItemIsEditable)
                        item.setData(Qt.UserRole, review_key)
                        item.setData(Qt.UserRole + 1, field_name)

                    self.user_preview_table.setItem(row_index, column, item)

                self._update_review_row_highlight(row_index)

            self._refresh_user_review_progress_summary_from_table()

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

    def run_validation(self) -> None:
        file_paths = self._current_file_paths()
        if not file_paths or not self.selected_slot_key:
            QMessageBox.warning(self, "חסר קובץ", "יש לבחור קובץ מתוך אחד ממשבצות הקלט לפני הרצת הבדיקה.")
            self.tabs.setCurrentIndex(0)
            return

        self.tabs.setCurrentIndex(1)
        self._run_slot_validation(self.selected_slot_key, file_paths, show_feedback=True)

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
        invalid_slots = 0
        failed_slots: list[str] = []

        for slot_key, file_paths in selected_slots:
            slot_summary = self._run_slot_validation(slot_key, file_paths, show_feedback=False)
            processed_slots += 1
            processed_files += int(slot_summary["file_count"])
            total_rows += int(slot_summary["total_rows"])

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

        self.summary_group.show()
        self.results_group.show()

        if invalid_slots or failed_slots:
            QMessageBox.warning(self, "בדיקת קבוצה הושלמה עם ממצאים", "\n".join(summary_lines))
        else:
            QMessageBox.information(self, "בדיקת קבוצה הושלמה", "\n".join(summary_lines))

    def _run_slot_validation(self, slot_key: str, file_paths: list[str], show_feedback: bool = True) -> dict[str, object]:
        if slot_key == "AGR_1251":
            self.summary_labels["status"].setText("מעבד קובצי הרשאות גדולים במנות...")
            QApplication.processEvents()

        try:
            result = process_file(
                file_paths,
                required_columns=self._required_columns_for_slot(slot_key),
                output_dir=self.config.output_dir,
                source_name_override=slot_key,
            )
        except Exception as error:
            self.summary_labels["status"].setText(f"שגיאה בעיבוד {slot_key}")
            self.issues_table.setRowCount(0)
            error_message = f"אירעה שגיאה במהלך העיבוד של המשבצת {slot_key}: {error}"
            self.issues_table.insertRow(0)
            for column, value in enumerate(["מבנה", "SYSTEM", error_message]):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.issues_table.setItem(0, column, item)
            self._append_error_log_entries(slot_key, file_paths, str(error))
            if show_feedback:
                QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה במהלך העיבוד של המשבצת {slot_key}:\n{error}")
            return {
                "slot_key": slot_key,
                "status": "error",
                "file_count": len(file_paths),
                "total_rows": 0,
                "invalid_rows": 0,
                "is_valid": False,
            }

        self.summary_group.show()
        self.results_group.show()
        self.summary_labels["total"].setText(str(result.summary.total_rows))
        self.summary_labels["valid"].setText(str(result.summary.valid_rows))
        self.summary_labels["invalid"].setText(str(result.summary.invalid_rows))
        status_text = "תקין" if result.summary.is_valid else f"נמצאו שגיאות - {slot_key}"
        self.summary_labels["status"].setText(status_text)

        self.issues_table.setRowCount(0)
        if result.issues:
            for issue in result.issues:
                row_index = self.issues_table.rowCount()
                self.issues_table.insertRow(row_index)
                values = [
                    str(issue.row_number if issue.row_number > 0 else "מבנה"),
                    self.format_rtl_text(issue.column_name),
                    self.format_rtl_text(issue.message),
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.issues_table.setItem(row_index, column, item)
        else:
            self.issues_table.insertRow(0)
            for column, value in enumerate(["-", "-", "לא נמצאו שגיאות"]):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.issues_table.setItem(0, column, item)

        self._append_run_log_entries(slot_key, file_paths, result)
        if result.report_path is not None:
            self.report_path = result.report_path
        self.report_button.setEnabled(self.report_path is not None)
        file_count = len(result.source_files) if result.source_files else len(file_paths)

        if show_feedback:
            if result.summary.is_valid:
                QMessageBox.information(
                    self,
                    "הבדיקה הושלמה",
                    f"בדיקת המשבצת {slot_key} הסתיימה ללא שגיאות. נקלטו {file_count} קבצים.",
                )
            else:
                ordered_messages = []
                structure_messages = [issue.message for issue in result.issues if "אינו תואם למבנה" in issue.message]
                other_messages = [issue.message for issue in result.issues if "אינו תואם למבנה" not in issue.message]
                for message in structure_messages + other_messages:
                    if message not in ordered_messages:
                        ordered_messages.append(message)
                    if len(ordered_messages) == 3:
                        break
                summary_text = "\n".join(f"• {message}" for message in ordered_messages)
                QMessageBox.warning(
                    self,
                    "נמצאו שגיאות בבדיקה",
                    f"בדיקת המשבצת {slot_key} הסתיימה עם שגיאות.\n\n{summary_text}\n\nניתן לבצע לחיצה כפולה על הרשומה בלוג לצפייה בפירוט.",
                )

        return {
            "slot_key": slot_key,
            "status": "ok",
            "file_count": file_count,
            "total_rows": result.summary.total_rows,
            "invalid_rows": result.summary.invalid_rows,
            "is_valid": result.summary.is_valid,
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
            status_text = "שגוי" if file_issues else "תקין"
            checked_at = datetime.now()
            row_count = row_counts_by_file.get(file_name, 0)
            record = {
                "slot_key": display_slot_name,
                "report_group": report_group,
                "file_name": file_name,
                "extraction_date": extraction_date,
                "row_count": row_count,
                "status": status_text,
                "error_count": len(file_issues),
                "error_preview": self._build_issue_preview(file_issues),
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
                str(len(file_issues)),
                str(record["error_preview"]),
                str(record["date"]),
                str(record["time"]),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item.setToolTip(self.format_rtl_text(value))
                if column == 5:
                    item.setBackground(QColor("#fdecec") if status_text in {"שגוי", "שגיאה"} else QColor("#eaf7ea"))
                self.run_log_table.setItem(row_index, column, item)

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
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item.setToolTip(self.format_rtl_text(value))
                if column == 5:
                    item.setBackground(QColor("#fdecec"))
                self.run_log_table.setItem(row_index, column, item)

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
        if not issues:
            lines.append("לא נמצאו שגיאות בקובץ זה.")
        else:
            for issue in issues:
                row_label = issue.row_number if issue.row_number > 0 else "מבנה"
                lines.append(f"- שורה {row_label} / {issue.column_name}: {issue.message}")

        return self.format_rtl_text("\n".join(lines))

    def show_log_details(self, row_index: int, _column: int) -> None:
        if row_index < 0 or row_index >= len(self.run_log_records):
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("פירוט קובץ שנבדק")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(760, 420)

        layout = QVBoxLayout(dialog)
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        details_box.setPlainText(self._build_log_details(row_index))
        layout.addWidget(details_box)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
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
        self.report_button.setEnabled(False)
        self.issues_table.setRowCount(0)
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
        sheet.title = self.format_rtl_text("סקירת משתמשים")
        sheet.append(export_formal_names)

        review_status_col_index: int | None = None
        review_notes_col_index: int | None = None
        for idx, field in enumerate(export_field_names):
            if field == "REVIEW_STATUS":
                review_status_col_index = idx + 1  # 1-based Excel column
            elif field == "REVIEW_NOTES":
                review_notes_col_index = idx + 1  # 1-based Excel column

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

        if total_data_rows > 0 and (review_status_col_index is not None or review_notes_col_index is not None):
            from openpyxl.utils import get_column_letter  # noqa: F811
            warning_fill = PatternFill("solid", fgColor="FFF0C2")
            for excel_row_idx, preview_row in enumerate(sorted_rows, start=2):
                is_not_ok = preview_row.get("REVIEW_STATUS", "") == "נבדק - לא תקין"
                has_notes = bool((preview_row.get("REVIEW_NOTES", "") or "").strip())
                if is_not_ok and not has_notes:
                    if review_status_col_index is not None:
                        sheet.cell(row=excel_row_idx, column=review_status_col_index).fill = warning_fill
                    if review_notes_col_index is not None:
                        sheet.cell(row=excel_row_idx, column=review_notes_col_index).fill = warning_fill

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

        required_fields = {"BNAME", "REVIEW_STATUS"}
        if not required_fields.issubset(col_map.keys()):
            missing = required_fields - col_map.keys()
            QMessageBox.warning(
                self,
                "שגיאת ייבוא",
                f"הקובץ חסר עמודות נדרשות: {', '.join(sorted(missing))}\n"
                "ודא שהקובץ יוצא מהכלי ומכיל את עמודות הסקירה.",
            )
            workbook.close()
            return

        mandt_col = col_map.get("MANDT")
        bname_col = col_map["BNAME"]
        status_col = col_map["REVIEW_STATUS"]
        notes_col = col_map.get("REVIEW_NOTES")

        imported_count = 0
        for row_values in rows_iter:
            bname = str(row_values[bname_col]).strip() if bname_col < len(row_values) and row_values[bname_col] is not None else ""
            if not bname:
                continue
            mandt = str(row_values[mandt_col]).strip() if mandt_col is not None and mandt_col < len(row_values) and row_values[mandt_col] is not None else ""
            review_key = self._user_reviewer_state_key(mandt, bname)

            raw_status = row_values[status_col] if status_col < len(row_values) else None
            status_value = self._normalize_reviewer_status(str(raw_status).strip() if raw_status is not None else "")

            raw_notes = row_values[notes_col] if notes_col is not None and notes_col < len(row_values) else None
            notes_value = str(raw_notes).strip() if raw_notes is not None else ""

            current = self.user_reviewer_state.setdefault(review_key, self._default_reviewer_values().copy())
            current["REVIEW_STATUS"] = status_value
            current["REVIEW_NOTES"] = notes_value
            imported_count += 1

        workbook.close()
        self._save_user_reviewer_state()
        self.refresh_user_preview()

        QMessageBox.information(
            self,
            "הייבוא הושלם",
            f"יובאו בהצלחה {imported_count} שורות מקובץ הסקירה.",
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
