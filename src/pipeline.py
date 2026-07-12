from pathlib import Path
from typing import Iterable

from src.models.validation_result import ValidationIssue, ValidationResult
from src.readers.excel_reader import ExcelFileReader
from src.readers.text_reader import TextFileReader
from src.readers.sap_transport_reader import SapTransportReader
from src.reporting.excel_report import ExcelReportWriter
from src.validators.engine import ValidationEngine
from src.validators.intake_rules import has_intake_issues


MULTI_FILE_SAMPLE_LIMIT = 12000
AGR_1251_BATCH_SIZE = 20000

_TRANSPORT_READER_SLOTS = {"STMS", "E070"}


def process_file(
    file_path: str | Path | Iterable[str | Path] | None = None,
    required_columns: list[str] | None = None,
    output_dir: str | Path | None = None,
    source_name_override: str | None = None,
    input_files: dict[str, list[str | Path]] | None = None,
    authorized_users: list[str] | None = None,
    strong_profiles: list[str] | None = None,
) -> ValidationResult:
    """Load SAP export file(s), run validation, and optionally write an intake report.

    Prefers ``input_files`` (slot_key → paths) when provided and non-empty.
    Otherwise uses the legacy ``file_path`` argument and wraps it as a one-key
    source map before processing.

    Args:
        file_path: Single path or iterable of paths (legacy entry point).
        required_columns: Optional required column names for validation.
        output_dir: When set, writes an Excel intake-errors report if needed.
        source_name_override: Slot/profile name override (e.g. ``AGR_1251``).
        input_files: Mapping of source/slot key to input paths.
        authorized_users: STMS authorized users for control checks.
        strong_profiles: Strong profile names for UST04/USH04 checks.

    Returns:
        ValidationResult with rows (or sample for large AGR_1251), issues,
        data_map, and optional report_path.
    """
    if input_files:
        resolved_source_map: dict[str, list[Path]] = {
            key: _normalize_paths(paths)
            for key, paths in input_files.items()
            if paths
        }
        if resolved_source_map:
            return _process_source_map(
                resolved_source_map,
                required_columns,
                output_dir,
                source_name_override,
                authorized_users,
                strong_profiles,
            )

    paths = _normalize_paths(file_path)  # type: ignore[arg-type]
    source_name = source_name_override or paths[0].name
    return _process_source_map(
        {source_name: paths},
        required_columns,
        output_dir,
        source_name_override,
        authorized_users,
        strong_profiles,
    )


def _process_source_map(
    source_map: dict[str, list[Path]],
    required_columns: list[str] | None,
    output_dir: str | Path | None,
    source_name_override: str | None,
    authorized_users: list[str] | None = None,
    strong_profiles: list[str] | None = None,
) -> ValidationResult:
    """Process a source_key → paths mapping and populate ``data_map`` per key."""
    first_key = next(iter(source_map))
    source_name = source_name_override or first_key
    all_paths = [p for paths in source_map.values() for p in paths]
    engine = _configure_engine(required_columns, authorized_users, strong_profiles)

    if source_name == "AGR_1251":
        result = _process_agr1251_in_batches(all_paths, engine, source_name)
        data_map: dict[str, list[dict]] = {}
        for key, paths in source_map.items():
            names = {p.name for p in paths}
            data_map[key] = [r for r in result.rows if r.get("__source_file") in names]
        result.data_map = data_map
    else:
        rows: list[dict] = []
        file_row_counts: dict[str, int] = {}
        data_map = {}
        for key, paths in source_map.items():
            key_rows: list[dict] = []
            for path in paths:
                file_rows = _read_rows(path, source_hint=key)
                file_row_counts[path.name] = len(file_rows)
                annotated = _attach_source(file_rows, path)
                rows.extend(annotated)
                key_rows.extend(annotated)
            data_map[key] = key_rows
        result = engine.run_all(data_map, source_name=source_name)
        result.source_files = [p.name for p in all_paths]
        result.file_row_counts = file_row_counts
        result.total_processed_rows = len(rows)
        result.data_map = data_map

    _write_intake_report_if_needed(result, all_paths[0], output_dir)
    return result


def _configure_engine(
    required_columns: list[str] | None,
    authorized_users: list[str] | None,
    strong_profiles: list[str] | None,
) -> ValidationEngine:
    """Create and configure a ValidationEngine for this run."""
    engine = ValidationEngine(required_columns=required_columns or [])
    if authorized_users:
        engine.set_authorized_users(authorized_users)
    if strong_profiles is not None:
        engine.set_strong_profiles(strong_profiles)
    return engine


def _write_intake_report_if_needed(
    result: ValidationResult,
    report_base_path: Path,
    output_dir: str | Path | None,
) -> None:
    """Write an Excel intake-errors report when output_dir is set and issues exist."""
    if output_dir is None or not has_intake_issues(result.issues):
        return
    report_writer = ExcelReportWriter()
    result.report_path = report_writer.write(result, report_base_path, Path(output_dir))


def _normalize_paths(file_path: str | Path | Iterable[str | Path]) -> list[Path]:
    """Normalize input path(s) to existing ``Path`` objects.

    Raises:
        FileNotFoundError: If any path does not exist on disk.
    """
    raw_paths = list(file_path) if isinstance(file_path, (list, tuple, set)) else [file_path]
    paths = [Path(item) for item in raw_paths]
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")
    return paths


def _read_rows(path: Path, source_hint: str | None = None) -> list[dict]:
    """Read rows from a supported export file, choosing the reader by type/slot.

    Raises:
        ValueError: If the file suffix is not supported.
    """
    suffix = path.suffix.lower()
    if source_hint in _TRANSPORT_READER_SLOTS and suffix in {".txt", ".csv"}:
        return SapTransportReader().read(path)
    if suffix in {".txt", ".csv"}:
        return TextFileReader().read(path)
    if suffix in {".xlsx", ".xlsm"}:
        return ExcelFileReader().read(path)
    raise ValueError(f"Unsupported file type: {suffix}")


def _attach_source(rows: list[dict], path: Path) -> list[dict]:
    """Annotate each row with ``__source_file`` set to the file name."""
    return [{**row, "__source_file": path.name} for row in rows]


def _process_agr1251_in_batches(
    paths: list[Path],
    engine: ValidationEngine,
    source_name: str,
) -> ValidationResult:
    """Validate large AGR_1251 files in batches; keep only a sample of rows in memory."""
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
        total_processed_rows=total_rows,
    )
