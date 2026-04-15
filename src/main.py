import argparse
from pathlib import Path

from src.config import AppConfig
from src.pipeline import process_file
from src.ui.desktop_app import launch_desktop_app


def run() -> None:
    parser = argparse.ArgumentParser(
        description="טעינת קבצי TXT/CSV/XLSX והרצת בדיקות תקינות.",
    )
    parser.add_argument("file_path", nargs="?", help="נתיב לקובץ הקלט")
    parser.add_argument(
        "--required",
        nargs="*",
        default=["user_id", "name"],
        help="עמודות חובה לבדיקה",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="הרצה במצב שורת פקודה במקום פתיחת הממשק הגרפי",
    )
    args = parser.parse_args()

    if not args.file_path and not args.cli:
        launch_desktop_app()
        return

    if not args.file_path:
        config = AppConfig.default()
        print("הפרויקט מוכן לקליטת קבצים ולבדיקות תקינות.")
        print(f"תיקיית קלט: {config.input_dir}")
        print(f"תיקיית פלט: {config.output_dir}")
        print(f"סיומות נתמכות: {', '.join(config.supported_extensions)}")
        return

    config = AppConfig.default()
    result = process_file(
        Path(args.file_path),
        required_columns=args.required,
        output_dir=config.output_dir,
    )
    print(f"שורות שנבדקו: {result.summary.total_rows}")
    print(f"שורות תקינות: {result.summary.valid_rows}")
    print(f"שורות שגויות: {result.summary.invalid_rows}")
    print(f"הקובץ תקין: {result.summary.is_valid}")

    for issue in result.issues:
        row_label = issue.row_number if issue.row_number > 0 else "מבנה"
        print(f"- שורה {row_label} / {issue.column_name}: {issue.message}")

    if result.report_path is not None:
        print(f"דוח אקסל: {result.report_path}")


if __name__ == "__main__":
    run()
