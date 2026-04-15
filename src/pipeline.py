from pathlib import Path

from src.models.validation_result import ValidationResult
from src.readers.excel_reader import ExcelFileReader
from src.readers.text_reader import TextFileReader
from src.reporting.excel_report import ExcelReportWriter
from src.validators.engine import ValidationEngine


def process_file(
    file_path: str | Path,
    required_columns: list[str] | None = None,
    output_dir: str | Path | None = None,
) -> ValidationResult:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix in {".txt", ".csv"}:
        rows = TextFileReader().read(path)
    elif suffix in {".xlsx", ".xlsm"}:
        rows = ExcelFileReader().read(path)
    else:
        raise ValueError(f"Unsupported file type: {suffix}")

    engine = ValidationEngine(required_columns=required_columns or [])
    result = engine.validate(rows)

    if output_dir is not None:
        report_writer = ExcelReportWriter()
        result.report_path = report_writer.write(result, path, Path(output_dir))

    return result
