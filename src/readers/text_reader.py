import csv
import re
from itertools import chain
from pathlib import Path
from typing import Any, Iterator


class TextFileReader:
    METADATA_PREFIXES = ("table:", "displayed fields:")
    ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "cp1255", "cp1252", "latin-1")

    def read(self, file_path: Path) -> list[dict[str, Any]]:
        return [row for batch in self.read_in_batches(file_path, chunk_size=50000) for row in batch]

    def read_in_batches(self, file_path: Path, chunk_size: int = 20000) -> Iterator[list[dict[str, Any]]]:
        encoding = self._detect_encoding(file_path)

        with file_path.open("r", encoding=encoding, newline="") as handle:
            preview_lines: list[str] = []
            for _ in range(80):
                line = handle.readline()
                if not line:
                    break
                preview_lines.append(line)

            if not preview_lines:
                return

            if self._looks_like_key_value_text(preview_lines):
                current_batch: list[dict[str, Any]] = []
                for line in chain(preview_lines, handle):
                    parsed_row = self._parse_key_value_line(line)
                    if not parsed_row:
                        continue
                    current_batch.append(parsed_row)
                    if len(current_batch) >= chunk_size:
                        yield current_batch
                        current_batch = []

                if current_batch:
                    yield current_batch
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
    def _detect_encoding(cls, file_path: Path) -> str:
        sample = file_path.read_bytes()[:65536]
        if not sample:
            return "utf-8-sig"

        for encoding in cls.ENCODING_CANDIDATES:
            try:
                sample.decode(encoding)
                return encoding
            except UnicodeDecodeError:
                continue

        return "latin-1"

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

            identifier_like = [
                cell
                for cell in cells
                if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_/@.()\- %]*", cell)
            ]
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
    def _looks_like_key_value_text(lines: list[str]) -> bool:
        matches = 0
        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith(("#", ";", "[")):
                continue
            if re.match(r"^[A-Za-z0-9_./-]+\s*(=|:|\s{2,})\s*.+$", stripped):
                matches += 1
            if matches >= 2:
                return True
        return False

    @staticmethod
    def _parse_key_value_line(line: str) -> dict[str, Any]:
        stripped = line.strip()
        if not stripped or stripped.startswith(("#", ";", "[")):
            return {}

        match = re.match(r"^(?P<name>[A-Za-z0-9_./-]+)\s*(=|:|\s{2,})\s*(?P<value>.+?)\s*$", stripped)
        if not match:
            return {}

        return {
            "NAME": match.group("name").strip(),
            "VALUE": match.group("value").strip(),
        }

    @staticmethod
    def _detect_delimiter(sample: str) -> str:
        candidates = ["\t", ";", "|", ","]
        delimiter_scores = {
            delimiter: sum(line.count(delimiter) for line in sample.splitlines())
            for delimiter in candidates
        }
        best_delimiter = max(candidates, key=lambda delimiter: delimiter_scores[delimiter])
        return best_delimiter if delimiter_scores[best_delimiter] > 0 else ","
