import os
import re
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
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
    def __init__(self, base_dir: Path | None = None) -> None:
        super().__init__()
        self.config = AppConfig.default(base_dir or Path.cwd())
        self.config.input_dir.mkdir(parents=True, exist_ok=True)
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.report_path: Path | None = None

        self.summary_labels: dict[str, QLabel] = {}

        self._configure_window()
        self._build_ui()

    @staticmethod
    def format_rtl_text(text: object) -> str:
        raw_text = "" if text is None else str(text)
        return re.sub(r"[\u2066\u2067\u2068\u2069\u200e\u200f]", "", raw_text)

    def _configure_window(self) -> None:
        self.setWindowTitle(self.format_rtl_text("מערכת בדיקות קבצים - ITGC"))
        self.setMinimumSize(980, 680)
        self.resize(1100, 760)
        self.setLayoutDirection(Qt.RightToLeft)

    def _build_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        header_label = QLabel(self.format_rtl_text("מסך בדיקת קבצי TXT ו-Excel"))
        header_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #16325c;")
        main_layout.addWidget(header_label)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)

        self.select_button = QPushButton("בחירת קובץ")
        self.select_button.clicked.connect(self.choose_file)
        buttons_layout.addWidget(self.select_button)

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

        main_layout.addLayout(buttons_layout)

        self.source_group = QGroupBox("קובץ ופרמטרים")
        self.source_group.setAlignment(Qt.AlignRight)
        source_layout = QGridLayout(self.source_group)
        source_layout.setContentsMargins(12, 18, 12, 12)
        source_layout.setHorizontalSpacing(10)
        source_layout.setVerticalSpacing(10)

        source_layout.addWidget(QLabel("נתיב הקובץ הנבחר:"), 0, 1)
        self.file_display_label = QLabel("טרם נבחר קובץ")
        self.file_display_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.file_display_label.setWordWrap(True)
        self.file_display_label.setStyleSheet("padding: 6px; background: #ffffff; border: 1px solid #cfd6e4;")
        source_layout.addWidget(self.file_display_label, 0, 0)

        source_layout.addWidget(QLabel("עמודות חובה:"), 1, 1)
        self.required_columns_edit = QLineEdit("user_id,name,email")
        self.required_columns_edit.setAlignment(Qt.AlignRight)
        source_layout.addWidget(self.required_columns_edit, 1, 0)

        main_layout.addWidget(self.source_group)

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

        main_layout.addWidget(self.summary_group)

        self.results_group = QGroupBox("רשימת שגיאות")
        self.results_group.setAlignment(Qt.AlignRight)
        results_layout = QVBoxLayout(self.results_group)
        results_layout.setContentsMargins(12, 18, 12, 12)
        results_layout.setSpacing(10)
        self.issues_table = QTableWidget(0, 3)
        self.issues_table.setHorizontalHeaderLabels(["מספר שורה", "שם עמודה", "הודעת שגיאה"])
        self.issues_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.issues_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.issues_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.issues_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.issues_table.setAlternatingRowColors(True)
        results_layout.addWidget(self.issues_table)

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

    def choose_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "בחירת קובץ לבדיקה",
            str(self.config.input_dir),
            "Supported files (*.txt *.csv *.xlsx *.xlsm);;All files (*.*)",
        )
        if file_path:
            self.file_display_label.setText(self.format_rtl_text(file_path))

    def _parse_required_columns(self) -> list[str]:
        raw_value = self.required_columns_edit.text().replace(";", ",").replace("\n", ",")
        return [item.strip() for item in raw_value.split(",") if item.strip()]

    def _current_file_path(self) -> str:
        displayed = self.file_display_label.text().strip()
        return "" if displayed == "טרם נבחר קובץ" else displayed

    def run_validation(self) -> None:
        file_path = self._current_file_path()
        if not file_path:
            QMessageBox.warning(self, "חסר קובץ", "יש לבחור קובץ לפני הרצת הבדיקה.")
            return

        try:
            result = process_file(
                Path(file_path),
                required_columns=self._parse_required_columns(),
                output_dir=self.config.output_dir,
            )
        except Exception as error:
            QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה במהלך העיבוד:\n{error}")
            return

        self.summary_labels["total"].setText(str(result.summary.total_rows))
        self.summary_labels["valid"].setText(str(result.summary.valid_rows))
        self.summary_labels["invalid"].setText(str(result.summary.invalid_rows))
        self.summary_labels["status"].setText("תקין" if result.summary.is_valid else "נמצאו שגיאות")

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

        self.report_path = result.report_path
        self.report_button.setEnabled(self.report_path is not None)
        QMessageBox.information(self, "הבדיקה הושלמה", "הקובץ נבדק בהצלחה ודוח האקסל נוצר בתיקיית הפלט.")

    def clear_results(self) -> None:
        self.file_display_label.setText("טרם נבחר קובץ")
        self.required_columns_edit.setText("user_id,name,email")
        self.summary_labels["total"].setText("0")
        self.summary_labels["valid"].setText("0")
        self.summary_labels["invalid"].setText("0")
        self.summary_labels["status"].setText("ממתין להרצה")
        self.report_path = None
        self.report_button.setEnabled(False)
        self.issues_table.setRowCount(0)

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
