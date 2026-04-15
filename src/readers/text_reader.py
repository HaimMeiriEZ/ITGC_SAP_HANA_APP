import csv
from pathlib import Path
from typing import Any


class TextFileReader:
    def read(self, file_path: Path) -> list[dict[str, Any]]:
        lines = file_path.read_text(encoding="utf-8-sig").splitlines()
        if not lines:
            return []

        sample = "\n".join(lines[:5])
        delimiter = self._detect_delimiter(sample)
        reader = csv.DictReader(lines, delimiter=delimiter)
        return [dict(row) for row in reader]

    @staticmethod
    def _detect_delimiter(sample: str) -> str:
        for delimiter in [";", ",", "\t", "|"]:
            if delimiter in sample:
                return delimiter
        return ","
