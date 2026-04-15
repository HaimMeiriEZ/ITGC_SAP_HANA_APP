from pathlib import Path
from typing import Any

from openpyxl import load_workbook


class ExcelFileReader:
    def read(self, file_path: Path) -> list[dict[str, Any]]:
        workbook = load_workbook(file_path, data_only=True, read_only=True)
        sheet = workbook.active
        row_iterator = sheet.iter_rows(values_only=True)
        header_row = next(row_iterator, None)

        if not header_row:
            return []

        headers = [str(value).strip() for value in header_row if value is not None]
        data_rows: list[dict[str, Any]] = []

        for row in row_iterator:
            if not any(value not in (None, "") for value in row):
                continue
            item = {
                headers[index]: row[index] if index < len(row) else None
                for index in range(len(headers))
            }
            data_rows.append(item)

        workbook.close()
        return data_rows
