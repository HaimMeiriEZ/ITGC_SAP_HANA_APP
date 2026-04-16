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
        self.selected_slot_key: str | None = None
        self.summary_labels: dict[str, QLabel] = {}
        self.run_log_records: list[dict[str, object]] = []

        self._configure_window()
        self._build_ui()

    @staticmethod
    def format_rtl_text(text: object) -> str:
        raw_text = "" if text is None else str(text)
        return re.sub(r"[\u2066\u2067\u2068\u2069\u200e\u200f]", "", raw_text)

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
        self.page_scroll.setWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        header_label = QLabel(self.format_rtl_text("מסך בדיקת קלטי SAP HANA APP"))
        header_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #16325c;")
        main_layout.addWidget(header_label)

        hint_label = QLabel("בחר קבצים לפי הסלוט המתאים. כוכבית מציינת סלוט חובה. ניתן להריץ בדיקה נפרדת לכל קבוצת קבצים בלי להמתין לטעינת כל הדוחות.")
        hint_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        hint_label.setWordWrap(True)
        hint_label.setStyleSheet("color: #4f5d73;")
        main_layout.addWidget(hint_label)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)

        self.run_button = QPushButton("הרץ בדיקה")
        self.run_button.clicked.connect(self.run_validation)
        buttons_layout.addWidget(self.run_button)

        self.report_button = QPushButton("פתח דוח אקסל")
        self.report_button.setEnabled(False)
        self.report_button.clicked.connect(self.open_report)
        buttons_layout.addWidget(self.report_button)

        self.output_button = QPushButton("פתח תיקיית פלט")
        self.output_button.clicked.connect(self.open_output_folder)
        buttons_layout.addWidget(self.output_button)

        self.clear_button = QPushButton("נקה מסך")
        self.clear_button.clicked.connect(self.clear_results)
        buttons_layout.addWidget(self.clear_button)

        buttons_layout.addStretch()
        main_layout.addLayout(buttons_layout)

        self.slots_group = QGroupBox("מקורות קלט לבדיקת SAP HANA APP")
        self.slots_group.setAlignment(Qt.AlignRight)
        slots_group_layout = QVBoxLayout(self.slots_group)
        slots_group_layout.setContentsMargins(8, 18, 8, 8)

        self.slots_scroll = QScrollArea()
        self.slots_scroll.setWidgetResizable(True)
        self.slots_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.slots_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.slots_scroll.setMinimumHeight(520)
        self.slots_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        slots_container = QWidget()
        slots_layout = QGridLayout(slots_container)
        slots_layout.setContentsMargins(12, 12, 12, 12)
        slots_layout.setHorizontalSpacing(12)
        slots_layout.setVerticalSpacing(10)
        slots_layout.setColumnStretch(0, 0)
        slots_layout.setColumnStretch(1, 0)
        slots_layout.setColumnStretch(2, 1)
        slots_layout.setColumnStretch(3, 0)

        current_row = 0
        for category in self._ordered_categories():
            category_label = QLabel(category)
            category_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            category_label.setStyleSheet("font-weight: bold; color: #16325c; margin-top: 8px;")

            category_button = QPushButton("הרץ בדיקה")
            category_button.setMinimumHeight(32)
            category_button.setToolTip(self.format_rtl_text(f"הרצת בדיקה עבור קבוצת {category}"))
            category_button.clicked.connect(
                lambda _checked=False, cat=category: self.run_category_validation(cat)
            )
            self.category_run_buttons[category] = category_button

            slots_layout.addWidget(category_label, current_row, 1, 1, 3)
            slots_layout.addWidget(category_button, current_row, 0)
            current_row += 1

            for slot_key, metadata in self.SLOT_DEFINITIONS.items():
                if metadata["category"] != category:
                    continue

                display_name = metadata.get("label", slot_key)
                slot_title = QLabel(f"{display_name}{' *' if metadata['required'] else ''}")
                slot_title.setAlignment(Qt.AlignRight | Qt.AlignTop)
                slot_title.setStyleSheet("font-weight: bold;")
                slot_title.setMinimumWidth(110)

                description = QLabel(metadata["description"])
                description.setAlignment(Qt.AlignRight | Qt.AlignTop)
                description.setWordWrap(True)
                description.setMinimumHeight(34)
                description.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

                sample = QLabel(f"קובץ צפוי: {metadata['expected_file']}")
                sample.setAlignment(Qt.AlignRight | Qt.AlignTop)
                sample.setStyleSheet("color: #5b6573;")
                sample.setMinimumWidth(180)

                status_label = QLabel("טרם נבחר קובץ")
                status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                status_label.setWordWrap(True)
                status_label.setMinimumHeight(32)
                status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                status_label.setStyleSheet("padding: 6px; background: #ffffff; border: 1px solid #cfd6e4;")

                select_button = QPushButton("בחירת קבצים" if slot_key in self.MULTI_FILE_SLOTS else "בחירת קובץ")
                select_button.setMinimumHeight(34)
                select_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                select_button.clicked.connect(lambda _checked=False, sk=slot_key: self.choose_file(sk))

                slots_layout.setRowMinimumHeight(current_row, 42)
                slots_layout.addWidget(slot_title, current_row, 3)
                slots_layout.addWidget(description, current_row, 2)
                slots_layout.addWidget(sample, current_row, 1)
                slots_layout.addWidget(select_button, current_row, 0)
                current_row += 1
                slots_layout.setRowMinimumHeight(current_row, 36)
                slots_layout.addWidget(status_label, current_row, 0, 1, 4)
                current_row += 1

                self.slot_widgets[slot_key] = {
                    "path_label": status_label,
                    "button": select_button,
                    "metadata": metadata,
                    "selected_paths": [],
                }

        bottom_spacer = QLabel("")
        bottom_spacer.setMinimumHeight(120)
        slots_layout.addWidget(bottom_spacer, current_row, 0, 1, 4)
        current_row += 1
        slots_layout.setRowStretch(current_row, 1)
        slots_layout.setRowMinimumHeight(current_row, 20)
        self.slots_scroll.setWidget(slots_container)
        slots_group_layout.addWidget(self.slots_scroll)

        main_layout.addWidget(self.slots_group)

        self.run_log_group = QGroupBox("לוג קבצים שנבדקו")
        self.run_log_group.setAlignment(Qt.AlignRight)
        run_log_layout = QVBoxLayout(self.run_log_group)
        run_log_layout.setContentsMargins(12, 18, 12, 12)
        self.run_log_table = QTableWidget(0, 5)
        self.run_log_table.setHorizontalHeaderLabels(["סלוט", "קובץ", "סטטוס", "מספר שגיאות", "שעת בדיקה"])
        self.run_log_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.run_log_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.run_log_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.run_log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.run_log_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.run_log_table.setAlternatingRowColors(True)
        self.run_log_table.setMinimumHeight(260)
        self.run_log_table.setToolTip("לחיצה כפולה על שורה תפתח פירוט מלא עבור הקובץ")
        self.run_log_table.cellDoubleClicked.connect(self.show_log_details)
        run_log_layout.addWidget(self.run_log_table)
        main_layout.addWidget(self.run_log_group)

        self.required_columns_group = QGroupBox("עמודות חובה לבדיקה")
        self.required_columns_group.setAlignment(Qt.AlignRight)
        required_layout = QHBoxLayout(self.required_columns_group)
        self.required_columns_edit = QLineEdit("")
        self.required_columns_edit.setAlignment(Qt.AlignRight)
        self.required_columns_edit.setPlaceholderText("יוזן אוטומטית לפי הסלוט שנבחר")
        required_layout.addWidget(self.required_columns_edit)
        self.required_columns_group.hide()
        main_layout.addWidget(self.required_columns_group)

        self.summary_group = QGroupBox("סיכום בדיקה")
        self.summary_group.setAlignment(Qt.AlignRight)
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

        self.results_group = QGroupBox("רשימת שגיאות")
        self.results_group.setAlignment(Qt.AlignRight)
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
                right: 18px;
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
            "STMS": "TRKORR,STATUS",
            "RSPARAM": "PARAMETER,VALUE",
            "TPFET": "PARAMETER,VALUE",
        }
        return suggestions.get(slot_key, "")

    def run_validation(self) -> None:
        file_paths = self._current_file_paths()
        if not file_paths or not self.selected_slot_key:
            QMessageBox.warning(self, "חסר קובץ", "יש לבחור קובץ מתוך אחד מסלוטי הקלט לפני הרצת הבדיקה.")
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
                f"בקבוצה {category} חסרים קבצי חובה עבור הסלוטים: {', '.join(missing_required)}.\n\nהבדיקה תמשיך עבור הקבצים שנבחרו.",
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
            f"סלוטים שנבדקו: {processed_slots}",
            f"קבצים שנבדקו: {processed_files}",
            f"שורות שנבדקו: {total_rows}",
        ]

        if invalid_slots:
            summary_lines.append(f"סלוטים עם ממצאים: {invalid_slots}")
        if failed_slots:
            summary_lines.append(f"סלוטים שנכשלו בעיבוד: {', '.join(failed_slots)}")
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
            if show_feedback:
                QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה במהלך העיבוד של הסלוט {slot_key}:\n{error}")
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
                    f"בדיקת הסלוט {slot_key} הסתיימה ללא שגיאות. נקלטו {file_count} קבצים.",
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
                    f"בדיקת הסלוט {slot_key} הסתיימה עם שגיאות.\n\n{summary_text}\n\nניתן לבצע לחיצה כפולה על הרשומה בלוג לצפייה בפירוט.",
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
            record = {
                "slot_key": slot_key,
                "file_name": file_name,
                "status": status_text,
                "error_count": len(file_issues),
                "timestamp": datetime.now().strftime("%H:%M:%S"),
                "issues": list(file_issues),
            }
            self.run_log_records.append(record)

            row_index = self.run_log_table.rowCount()
            self.run_log_table.insertRow(row_index)
            values = [
                slot_key,
                file_name,
                status_text,
                str(len(file_issues)),
                str(record["timestamp"]),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(self.format_rtl_text(value))
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if column == 2:
                    item.setBackground(QColor("#fdecec") if status_text == "שגוי" else QColor("#eaf7ea"))
                self.run_log_table.setItem(row_index, column, item)

    def _build_log_details(self, row_index: int) -> str:
        if row_index < 0 or row_index >= len(self.run_log_records):
            return "לא נמצא פירוט עבור הרשומה שנבחרה."

        record = self.run_log_records[row_index]
        lines = [
            f"סלוט: {record['slot_key']}",
            f"קובץ: {record['file_name']}",
            f"סטטוס: {record['status']}",
            f"מספר שגיאות: {record['error_count']}",
            f"שעת בדיקה: {record['timestamp']}",
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
