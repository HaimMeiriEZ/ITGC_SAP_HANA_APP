from pathlib import Path
from typing import Iterable

from src.models.validation_result import ValidationIssue, ValidationResult
from src.readers.excel_reader import ExcelFileReader
from src.readers.text_reader import TextFileReader
from src.reporting.excel_report import ExcelReportWriter
from src.validators.engine import ValidationEngine


MULTI_FILE_SAMPLE_LIMIT = 12000
AGR_1251_BATCH_SIZE = 20000


def process_file(
    file_path: str | Path | Iterable[str | Path],
    required_columns: list[str] | None = None,
    output_dir: str | Path | None = None,
    source_name_override: str | None = None,
) -> ValidationResult:
    paths = _normalize_paths(file_path)
    engine = ValidationEngine(required_columns=required_columns or [])
    source_name = source_name_override or paths[0].name

    if source_name_override == "AGR_1251":
        result = _process_agr1251_in_batches(paths, engine, source_name)
    else:
        rows: list[dict] = []
        file_row_counts: dict[str, int] = {}
        for path in paths:
            file_rows = _read_rows(path)
            file_row_counts[path.name] = len(file_rows)
            rows.extend(_attach_source(file_rows, path))
        result = engine.validate(rows, source_name=source_name)
        result.source_files = [path.name for path in paths]
        result.file_row_counts = file_row_counts
        result.total_rows_override = len(rows)

    if output_dir is not None:
        report_writer = ExcelReportWriter()
        result.report_path = report_writer.write(result, paths[0], Path(output_dir))

    return result


def _normalize_paths(file_path: str | Path | Iterable[str | Path]) -> list[Path]:
    raw_paths = list(file_path) if isinstance(file_path, (list, tuple, set)) else [file_path]
    paths = [Path(item) for item in raw_paths]
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")
    return paths


def _read_rows(path: Path) -> list[dict]:
    suffix = path.suffix.lower()
    if suffix in {".txt", ".csv"}:
        return TextFileReader().read(path)
    if suffix in {".xlsx", ".xlsm"}:
        return ExcelFileReader().read(path)
    raise ValueError(f"Unsupported file type: {suffix}")


def _attach_source(rows: list[dict], path: Path) -> list[dict]:
    return [{**row, "__source_file": path.name} for row in rows]


def _process_agr1251_in_batches(paths: list[Path], engine: ValidationEngine, source_name: str) -> ValidationResult:
    sample_rows: list[dict] = []
    issues: list[ValidationIssue] = []
    total_rows = 0
    row_offset = 0
    detected_profile: str | None = None
    seen_structure_issues: set[tuple[str, str, str]] = set()
    file_row_counts: dict[str, int] = {}

    for path in paths:
        suffix = path.suffix.lower()
        if suffix in {".txt", ".csv"}:
            batches = TextFileReader().read_in_batches(path, chunk_size=AGR_1251_BATCH_SIZE)
        elif suffix in {".xlsx", ".xlsm"}:
            batches = ExcelFileReader().read_in_batches(path, chunk_size=AGR_1251_BATCH_SIZE)
        else:
            raise ValueError(f"Unsupported file type: {suffix}")

        for batch in batches:
            annotated_batch = _attach_source(batch, path)
            if len(sample_rows) < MULTI_FILE_SAMPLE_LIMIT:
                sample_rows.extend(annotated_batch[: MULTI_FILE_SAMPLE_LIMIT - len(sample_rows)])

            batch_result = engine.validate(annotated_batch, source_name=source_name)
            if detected_profile is None:
                detected_profile = batch_result.detected_profile

            for issue in batch_result.issues:
                if issue.row_number == 0:
                    signature = (issue.column_name, issue.message, path.name)
                    if signature in seen_structure_issues:
                        continue
                    seen_structure_issues.add(signature)
                    issues.append(
                        ValidationIssue(
                            row_number=0,
                            column_name=issue.column_name,
                            message=issue.message,
                            source_file=path.name,
                        )
                    )
                    continue

                issues.append(
                    ValidationIssue(
                        row_number=issue.row_number + row_offset,
                        column_name=issue.column_name,
                        message=issue.message,
                        source_file=path.name,
                    )
                )

            total_rows += len(annotated_batch)
            file_row_counts[path.name] = file_row_counts.get(path.name, 0) + len(annotated_batch)
            row_offset += len(annotated_batch)

    return ValidationResult(
        rows=sample_rows,
        issues=issues,
        detected_profile=detected_profile,
        source_files=[path.name for path in paths],
        file_row_counts=file_row_counts,
        total_rows_override=total_rows,
    )
