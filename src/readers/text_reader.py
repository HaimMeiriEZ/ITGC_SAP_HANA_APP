import csv
import re
from itertools import chain
from pathlib import Path
from typing import Any, Iterator


class TextFileReader:
    METADATA_PREFIXES = ("table:", "displayed fields:")

    def read(self, file_path: Path) -> list[dict[str, Any]]:
        return [row for batch in self.read_in_batches(file_path, chunk_size=50000) for row in batch]

    def read_in_batches(self, file_path: Path, chunk_size: int = 20000) -> Iterator[list[dict[str, Any]]]:
        with file_path.open("r", encoding="utf-8-sig", newline="") as handle:
            preview_lines: list[str] = []
            for _ in range(80):
                line = handle.readline()
                if not line:
                    break
                preview_lines.append(line)

            if not preview_lines:
                return

            delimiter = self._detect_delimiter("".join(preview_lines))
            header_index = self._find_header_line_index(preview_lines, delimiter)
            relevant_preview = preview_lines[header_index:]
            reader = csv.DictReader(
                chain(relevant_preview, handle),
                delimiter=delimiter,
                skipinitialspace=True,
                restkey="__extra__",
            )

            current_batch: list[dict[str, Any]] = []
            for row in reader:
                normalized_row = self._normalize_row(row)
                if not normalized_row:
                    continue
                current_batch.append(normalized_row)
                if len(current_batch) >= chunk_size:
                    yield current_batch
                    current_batch = []

            if current_batch:
                yield current_batch

    @classmethod
    def _find_header_line_index(cls, lines: list[str], delimiter: str) -> int:
        for index, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                continue

            if stripped.casefold().startswith(cls.METADATA_PREFIXES):
                continue

            cells = [cell.strip() for cell in line.split(delimiter) if cell.strip()]
            if len(cells) < 2:
                continue

            identifier_like = [cell for cell in cells if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_/@.-]*", cell)]
            if len(identifier_like) >= 2:
                return index

        return 0

    @staticmethod
    def _normalize_row(row: dict[str, Any]) -> dict[str, Any]:
        normalized: dict[str, Any] = {}
        for key, value in row.items():
            if key in (None, "__extra__"):
                continue

            key_text = str(key).strip()
            if not key_text:
                continue

            if isinstance(value, list):
                value = " ".join(str(item).strip() for item in value if str(item).strip())
            elif isinstance(value, str):
                value = value.strip()

            normalized[key_text] = value

        return normalized

    @staticmethod
    def _detect_delimiter(sample: str) -> str:
        candidates = ["\t", ";", "|", ","]
        delimiter_scores = {
            delimiter: sum(line.count(delimiter) for line in sample.splitlines())
            for delimiter in candidates
        }
        best_delimiter = max(candidates, key=lambda delimiter: delimiter_scores[delimiter])
        return best_delimiter if delimiter_scores[best_delimiter] > 0 else ","
