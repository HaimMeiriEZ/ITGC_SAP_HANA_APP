import argparse
from pathlib import Path

from src.config import AppConfig
from src.pipeline import process_file


def run() -> None:
    parser = argparse.ArgumentParser(
        description="Load TXT/CSV/XLSX files and run basic integrity checks.",
    )
    parser.add_argument("file_path", nargs="?", help="Path to the input file")
    parser.add_argument(
        "--required",
        nargs="*",
        default=["user_id", "name"],
        help="Required columns to validate",
    )
    args = parser.parse_args()

    if not args.file_path:
        config = AppConfig.default()
        print("Project initialized for file intake validation.")
        print(f"Input folder: {config.input_dir}")
        print(f"Output folder: {config.output_dir}")
        print(f"Supported files: {', '.join(config.supported_extensions)}")
        return

    result = process_file(Path(args.file_path), required_columns=args.required)
    print(f"Rows checked: {result.summary.total_rows}")
    print(f"Valid rows: {result.summary.valid_rows}")
    print(f"Invalid rows: {result.summary.invalid_rows}")
    print(f"Overall valid: {result.summary.is_valid}")

    for issue in result.issues:
        row_label = issue.row_number if issue.row_number > 0 else "schema"
        print(f"- Row {row_label} / {issue.column_name}: {issue.message}")


if __name__ == "__main__":
    run()
