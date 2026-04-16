import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


class ExcelFileReader:
    def read(self, file_path: Path) -> list[dict[str, Any]]:
        workbook = load_workbook(file_path, data_only=True, read_only=True)
        best_rows: list[tuple[Any, ...]] = []
        best_header_index = 0
        best_score = -1

        for sheet in workbook.worksheets:
            all_rows = list(sheet.iter_rows(values_only=True))
            if not all_rows:
                continue

            header_index, score = self._find_header_row_index(all_rows)
            if score > best_score:
                best_rows = all_rows
                best_header_index = header_index
                best_score = score

        if not best_rows:
            workbook.close()
            return []

        header_row = best_rows[best_header_index]
        header_positions = [index for index, value in enumerate(header_row) if value not in (None, "")]
        headers = [str(header_row[index]).strip() for index in header_positions]
        data_rows: list[dict[str, Any]] = []

        for row in best_rows[best_header_index + 1 :]:
            if not any(value not in (None, "") for value in row):
                continue
            item = {
                headers[position]: row[index] if index < len(row) else None
                for position, index in enumerate(header_positions)
            }
            data_rows.append(item)

        workbook.close()
        return data_rows

    @staticmethod
    def _find_header_row_index(rows: list[tuple[Any, ...]]) -> tuple[int, int]:
        best_index = 0
        best_score = -1

        for index, row in enumerate(rows[:50]):
            cells = [str(value).strip() for value in row if value not in (None, "")]
            if len(cells) < 2:
                continue

            identifier_like = [cell for cell in cells if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_/@ .-]*", cell)]
            score = len(identifier_like)
            if score >= 2 and score > best_score:
                best_index = index
                best_score = score

        return best_index, best_score
