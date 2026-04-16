import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
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
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from src.config import AppConfig
from src.models.validation_result import ValidationIssue
from src.pipeline import process_file


def get_qt_app() -> QApplication:
    if "unittest" in sys.modules and "QT_QPA_PLATFORM" not in os.environ:
        os.environ["QT_QPA_PLATFORM"] = "offscreen"

    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        app.setLayoutDirection(Qt.RightToLeft)
        app.setFont(QFont("Segoe UI", 10))
    return app


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

    SLOT_DEFINITIONS = {
        "USR02": {
            "category": "טבלאות משתמשים",
            "description": "משתמשים - מקור חובה לבדיקות גישה, סטטוס ותאריכי התחברות.",
            "expected_file": "usr02_100.txt",
            "required": True,
        },
        "ADR6_USR21": {
            "label": "ADR6 / USR21",
            "category": "טבלאות משתמשים",
            "description": "ניתן להזין קובצי ADR6 או קובצי USR21 או את שניהם יחד לצורך הצלבת פרטי משתמש ואימייל.",
            "expected_file": "adr6.txt או usr21.txt",
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

    def __init__(self, base_dir: Path | None = None) -> None:
        super().__init__()
        self.config = AppConfig.default(base_dir or Path.cwd())
        self.config.input_dir.mkdir(parents=True, exist_ok=True)
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.report_path: Path | None = None
        self.slot_widgets: dict[str, dict[str, object]] = {}
        self.category_run_buttons: dict[str, QPushButton] = {}
        self.category_sections: dict[str, QGroupBox] = {}
        self.selected_slot_key: str | None = None
        self.summary_labels: dict[str, QLabel] = {}
        self.run_log_records: list[dict[str, object]] = []

        self._configure_window()
        self._build_ui()

    @staticmethod
    def format_rtl_text(text: object) -> str:
        raw_text = "" if text is None else str(text)
        return re.sub(r"[\u2066\u2067\u2068\u2069\u200e\u200f]", "", raw_text)

    @staticmethod
    def format_ui_rtl_text(text: object) -> str:
        normalized_text = ValidationDesktopApp.format_rtl_text(text).strip()
        if normalized_text and re.search(r"[\u0590-\u05FF]", normalized_text):
            return f"\u200f{normalized_text}"
        return normalized_text

    def _configure_window(self) -> None:
        self.setWindowTitle(self.format_rtl_text("מערכת בדיקות SAP HANA APP - ITGC"))
        self.setMinimumSize(1180, 760)
        self.resize(1280, 860)
        self.setLayoutDirection(Qt.RightToLeft)

    def _build_ui(self) -> None:
        self.page_scroll = QScrollArea()
        self.page_scroll.setWidgetResizable(True)
        self.page_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setCentralWidget(self.page_scroll)

        central_widget = QWidget()
        central_widget.setLayoutDirection(Qt.RightToLeft)
        self.page_scroll.setWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        self.header_label = QLabel(self.format_ui_rtl_text("מסך בדיקת קלטי SAP HANA APP"))
        self.header_label.setLayoutDirection(Qt.RightToLeft)
        self.header_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.header_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.header_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #16325c;")
        main_layout.addWidget(self.header_label)

        self.hint_label = QLabel(
            self.format_ui_rtl_text(
                "בחר קבצים לפי המשבצת המתאימה. כוכבית מציינת משבצת חובה. חובה לציין את תאריך ההפקה של הקבצים. ניתן להריץ בדיקה נפרדת לכל קבוצת קבצים בלי להמתין לטעינת כל הדוחות."
            )
        )
        self.hint_label.setLayoutDirection(Qt.RightToLeft)
        self.hint_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.hint_label.setWordWrap(True)
        self.hint_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.hint_label.setStyleSheet("color: #4f5d73;")
        main_layout.addWidget(self.hint_label)

        self.actions_row = QWidget()
        self.actions_row.setLayoutDirection(Qt.RightToLeft)
        self.actions_row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        buttons_layout = QHBoxLayout(self.actions_row)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        buttons_layout.addStretch(1)

        self.clear_button = QPushButton(self.format_ui_rtl_text("נקה מסך"))
        self.clear_button.clicked.connect(self.clear_results)
        buttons_layout.addWidget(self.clear_button)

        self.output_button = QPushButton(self.format_ui_rtl_text("פתח תיקיית פלט"))
        self.output_button.clicked.connect(self.open_output_folder)
        buttons_layout.addWidget(self.output_button)

        self.report_button = QPushButton(self.format_ui_rtl_text("פתח דוח אקסל"))
        self.report_button.setEnabled(False)
        self.report_button.clicked.connect(self.open_report)
        buttons_layout.addWidget(self.report_button)

        self.run_button = QPushButton(self.format_ui_rtl_text("הרץ בדיקה"))
        self.run_button.clicked.connect(self.run_validation)
        buttons_layout.addWidget(self.run_button)

        main_layout.addWidget(self.actions_row)

        self.slots_group = QGroupBox(self.format_ui_rtl_text("מקורות קלט לבדיקת SAP HANA APP"))
        self.slots_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.slots_group.setLayoutDirection(Qt.RightToLeft)
        slots_group_layout = QVBoxLayout(self.slots_group)
        slots_group_layout.setContentsMargins(8, 18, 8, 8)

        self.slots_scroll = QScrollArea()
        self.slots_scroll.setWidgetResizable(True)
        self.slots_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.slots_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.slots_scroll.setMinimumHeight(520)
        self.slots_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        slots_container = QWidget()
        slots_container.setLayoutDirection(Qt.RightToLeft)
        slots_layout = QGridLayout(slots_container)
        slots_layout.setContentsMargins(12, 12, 12, 12)
        slots_layout.setHorizontalSpacing(12)
        slots_layout.setVerticalSpacing(10)
        slots_layout.setColumnStretch(0, 0)
        slots_layout.setColumnStretch(1, 0)
        slots_layout.setColumnStretch(2, 1)
        slots_layout.setColumnStretch(3, 0)
        slots_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)

        current_row = 0
        for category in self._ordered_categories():
            palette = self._category_palette(category)
            category_section = QGroupBox(self.format_ui_rtl_text(category))
            category_section.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            category_section.setStyleSheet(
                f"""
                QGroupBox {{
                    font-weight: bold;
                    border: 1px solid {palette['border']};
                    border-radius: 10px;
                    margin-top: 14px;
                    padding-top: 18px;
                    background-color: #ffffff;
                }}
                QGroupBox::title {{
                    subcontrol-origin: margin;
                    subcontrol-position: top right;
                    padding: 4px 12px;
                    background-color: {palette['header']};
                    color: #16325c;
                    border-radius: 6px;
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
            category_layout.setColumnStretch(1, 0)
            category_layout.setColumnStretch(2, 1)
            category_layout.setColumnStretch(3, 0)
            category_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)

            category_button = QPushButton("הרץ בדיקה")
            category_button.setMinimumHeight(34)
            category_button.setToolTip(self.format_rtl_text(f"הרצת בדיקה עבור קבוצת {category}"))
            category_button.setStyleSheet(
                f"background-color: {palette['button']}; border: 1px solid {palette['border']};"
            )
            category_button.clicked.connect(
                lambda _checked=False, cat=category: self.run_category_validation(cat)
            )
            self.category_run_buttons[category] = category_button
            category_layout.addWidget(category_button, 0, 0)

            section_row = 1
            for slot_key, metadata in self.SLOT_DEFINITIONS.items():
                if metadata["category"] != category:
                    continue

                display_name = metadata.get("label", slot_key)
                slot_title = QLabel(self.format_ui_rtl_text(f"{display_name}{' *' if metadata['required'] else ''}"))
                slot_title.setLayoutDirection(Qt.RightToLeft)
                slot_title.setAlignment(Qt.AlignRight | Qt.AlignTop)
                slot_title.setStyleSheet("font-weight: bold;")
                slot_title.setMinimumWidth(110)

                description = QLabel(self.format_ui_rtl_text(metadata["description"]))
                description.setLayoutDirection(Qt.RightToLeft)
                description.setAlignment(Qt.AlignRight | Qt.AlignTop)
                description.setWordWrap(True)
                description.setMinimumHeight(34)
                description.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

                sample = QLabel(self.format_ui_rtl_text(f"קובץ צפוי: {metadata['expected_file']}"))
                sample.setLayoutDirection(Qt.RightToLeft)
                sample.setAlignment(Qt.AlignRight | Qt.AlignTop)
                sample.setStyleSheet("color: #5b6573;")
                sample.setMinimumWidth(180)

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

                category_layout.setRowMinimumHeight(section_row, 42)
                category_layout.addWidget(slot_title, section_row, 3)
                category_layout.addWidget(description, section_row, 2)
                category_layout.addWidget(sample, section_row, 1)
                category_layout.addWidget(select_button, section_row, 0)
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
                    "metadata": metadata,
                    "selected_paths": [],
                    "extraction_date_edit": extraction_date_edit,
                    "extraction_date_label": extraction_date_label,
                }

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

        main_layout.addWidget(self.slots_group)

        self.run_log_group = QGroupBox(self.format_ui_rtl_text("לוג קבצים שנבדקו"))
        self.run_log_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        run_log_layout = QVBoxLayout(self.run_log_group)
        run_log_layout.setContentsMargins(12, 18, 12, 12)
        self.run_log_table = QTableWidget(0, 10)
        self.run_log_table.setHorizontalHeaderLabels(["משבצת", "קבוצת דוחות", "קובץ", "תאריך הפקה", "רשומות שנקלטו", "סטטוס", "מספר שגיאות", "תיאור שגיאה", "תאריך בדיקה", "שעת בדיקה"])
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
        self.run_log_table.setMinimumHeight(260)
        self.run_log_table.setToolTip("לחיצה כפולה על שורה תפתח פירוט מלא עבור הקובץ")
        self.run_log_table.cellDoubleClicked.connect(self.show_log_details)
        run_log_layout.addWidget(self.run_log_table)
        main_layout.addWidget(self.run_log_group)

        self.required_columns_group = QGroupBox(self.format_ui_rtl_text("עמודות חובה לבדיקה"))
        self.required_columns_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        required_layout = QHBoxLayout(self.required_columns_group)
        self.required_columns_edit = QLineEdit("")
        self.required_columns_edit.setAlignment(Qt.AlignRight)
        self.required_columns_edit.setPlaceholderText("יוזן אוטומטית לפי המשבצת שנבחרה")
        required_layout.addWidget(self.required_columns_edit)
        self.required_columns_group.hide()
        main_layout.addWidget(self.required_columns_group)

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
        main_layout.addWidget(self.summary_group)

        self.results_group = QGroupBox(self.format_ui_rtl_text("רשימת שגיאות"))
        self.results_group.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        results_layout = QVBoxLayout(self.results_group)
        results_layout.setContentsMargins(12, 18, 12, 12)
        self.issues_table = QTableWidget(0, 3)
        self.issues_table.setHorizontalHeaderLabels(["מספר שורה", "שם עמודה", "הודעת שגיאה"])
        self.issues_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.issues_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.issues_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.issues_table.setAlternatingRowColors(True)
        results_layout.addWidget(self.issues_table)
        self.issues_table.setMinimumHeight(180)
        self.results_group.hide()
        main_layout.addWidget(self.results_group)

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
                subcontrol-position: top right;
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
        palettes = {
            "טבלאות משתמשים": {"header": "#e8f1ff", "button": "#dce9ff", "border": "#a9c2eb"},
            "טבלאות הרשאות כלליות": {"header": "#ede8ff", "button": "#e4dcff", "border": "#b8acec"},
            "טבלאות שינויים": {"header": "#e9faf1", "button": "#d9f3e5", "border": "#a8d8bd"},
            "מדיניות סיסמאות": {"header": "#fff3df", "button": "#ffe8bf", "border": "#e6c98b"},
        }
        return palettes.get(category, {"header": "#eef3fb", "button": "#e1eaf7", "border": "#bfd0e6"})

    @staticmethod
    def _default_extraction_date() -> str:
        return datetime.now().strftime("%Y-%m-%d")

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

    def _get_slot_extraction_date(self, slot_key: str) -> str:
        widget_data = self.slot_widgets.get(slot_key, {})
        date_edit = widget_data.get("extraction_date_edit")
        if isinstance(date_edit, QLineEdit):
            date_text = date_edit.text().strip()
            return date_text or "לא צוין"
        return "לא צוין"

    def choose_file(self, slot_key: str) -> None:
        if slot_key in self.MULTI_FILE_SLOTS:
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                f"בחירת קבצים עבור {slot_key}",
                str(self.config.input_dir),
                "Supported files (*.txt *.csv *.xlsx *.xlsm);;All files (*.*)",
            )
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                f"בחירת קובץ עבור {slot_key}",
                str(self.config.input_dir),
                "Supported files (*.txt *.csv *.xlsx *.xlsm);;All files (*.*)",
            )
            file_paths = [file_path] if file_path else []

        if file_paths:
            self.selected_slot_key = slot_key
            self.slot_widgets[slot_key]["selected_paths"] = file_paths
            self.slot_widgets[slot_key]["path_label"].setText(self._format_selected_files(file_paths))
            self.required_columns_edit.setText(self._suggest_required_columns(slot_key))

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

    def run_validation(self) -> None:
        file_paths = self._current_file_paths()
        if not file_paths or not self.selected_slot_key:
            QMessageBox.warning(self, "חסר קובץ", "יש לבחור קובץ מתוך אחד ממשבצות הקלט לפני הרצת הבדיקה.")
            return

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
                "slot_key": slot_key,
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
                slot_key,
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
                "slot_key": slot_key,
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
                slot_key,
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
        for widget_data in self.slot_widgets.values():
            widget_data["selected_paths"] = []
            widget_data["path_label"].setText("טרם נבחר קובץ")
            date_edit = widget_data.get("extraction_date_edit")
            if isinstance(date_edit, QLineEdit):
                date_edit.setText(self._default_extraction_date())
        self.required_columns_edit.setText("")
        self.summary_labels["total"].setText("0")
        self.summary_labels["valid"].setText("0")
        self.summary_labels["invalid"].setText("0")
        self.summary_labels["status"].setText("ממתין להרצה")
        self.report_path = None
        self.report_button.setEnabled(False)
        self.issues_table.setRowCount(0)
        self.run_log_records = []
        self.run_log_table.setRowCount(0)

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
