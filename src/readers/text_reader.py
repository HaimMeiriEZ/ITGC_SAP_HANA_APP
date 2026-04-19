import csv
import re
from itertools import chain
from pathlib import Path
from typing import Any, Iterator


class TextFileReader:
    METADATA_PREFIXES = ("table:", "displayed fields:", "entries for", "time interval")
    ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "cp1255", "cp1252", "latin-1")
    HEADER_HINTS = {
        "TRKORR",
        "REQUEST",
        "REQUEST NUMBER",
        "TRANSPORT REQUEST",
        "NUMBER",
        "DATE",
        "TIME",
        "CLT",
        "CLIENT",
        "OWNER",
        "USER",
        "USER NAME",
        "PROJECT",
        "SHORT TEXT",
        "RC",
        "RETURN CODE",
        "TRFUNCTION",
        "TRSTATUS",
        "TARSYSTEM",
        "KORRDEV",
        "AS4USER",
        "AS4DATE",
        "AS4TIME",
        "STRKORR",
        "AS4TEXT",
        "MANDT",
        "BNAME",
        "UFLAG",
        "TRDAT",
        "LTIME",
        "AGR_NAME",
        "OBJECT",
        "FIELD",
        "LOW",
        "HIGH",
        "UNAME",
        "PROFILE",
        "PARAMETER",
        "PARAMETER NAME",
        "VALUE",
        "CURRENT VALUE",
        "CONFIGURED VALUE",
        "NAME",
        "SMTP_ADDR",
        "ADDRNUMBER",
        "PERSNUMBER",
    }

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

    @staticmethod
    def _clean_cell_text(value: object) -> str:
        return re.sub(r"\s+", " ", str(value).strip().strip('"'))

    @classmethod
    def _find_header_line_index(cls, lines: list[str], delimiter: str) -> int:
        best_index = 0
        best_score = float("-inf")
        fallback_index: int | None = None

        for index, line in enumerate(lines):
            stripped = cls._clean_cell_text(line)
            if not stripped:
                continue

            if stripped.casefold().startswith(cls.METADATA_PREFIXES):
                continue

            cells = [cls._clean_cell_text(cell) for cell in line.split(delimiter) if cls._clean_cell_text(cell)]
            if len(cells) < 2:
                continue

            normalized_cells = [cell.upper() for cell in cells]
            header_matches = sum(1 for cell in normalized_cells if cell in cls.HEADER_HINTS)
            identifier_like = sum(
                1
                for cell in normalized_cells
                if re.fullmatch(r"[A-Z_][A-Z0-9_/@.()\- %]*", cell)
            )
            data_like = sum(
                1
                for cell in normalized_cells
                if re.fullmatch(r"[0-9./:\-]+", cell)
            )
            score = (header_matches * 4) + min(identifier_like, len(cells)) - (data_like * 2)

            if header_matches >= 2 and len(cells) >= 3:
                return index

            if fallback_index is None and identifier_like >= 2:
                fallback_index = index

            if score > best_score:
                best_score = score
                best_index = index

        if best_score > 0:
            return best_index
        if fallback_index is not None:
            return fallback_index
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

    @classmethod
    def _looks_like_key_value_text(cls, lines: list[str]) -> bool:
        sample = "".join(lines)
        delimiter = cls._detect_delimiter(sample)

        for line in lines:
            stripped = line.strip().strip('"')
            if not stripped:
                continue
            if stripped.casefold().startswith(cls.METADATA_PREFIXES):
                continue

            cells = [cls._clean_cell_text(cell) for cell in line.split(delimiter) if cls._clean_cell_text(cell)]
            if len(cells) >= 3:
                identifier_like = [
                    cell
                    for cell in cells
                    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_/@.()\- %]*", cell)
                ]
                if len(identifier_like) >= min(3, len(cells)):
                    return False

        matches = 0
        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.casefold().startswith(cls.METADATA_PREFIXES):
                continue
            if stripped.startswith(("#", ";", "[")):
                continue
            if "\t" in stripped and stripped.count("\t") >= 2:
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
