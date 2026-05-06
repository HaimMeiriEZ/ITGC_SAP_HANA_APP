import re
from pathlib import Path
from typing import Any, Iterator

from src.readers.text_reader import TextFileReader


class SapTransportReader:
    """Dedicated reader for SAP transport management reports (STMS/E070).

    Handles two formats:
    1. STMS log text (no header row; data rows identified by transport-number
       pattern ``[SYS]K[6-digits]``, e.g. ``FPDK901838``).
    2. Standard SAP table export (tab-delimited with a header row, e.g. E070
       exported via SE16).  Delegated to :class:`TextFileReader`.

    Output uses SAP-standard column names so downstream validation rules
    (``spec_rules.py``) work without any mapping layer:

    +-----------------+----------------------------------------------+
    | Output column   | Source                                       |
    +=================+==============================================+
    | TRKORR          | Transport request number                     |
    | AS4DATE         | Import/release date                          |
    | AS4TIME         | Import/release time                          |
    | MANDT           | SAP client (Mandant)                         |
    | AS4USER         | Transport owner / original developer         |
    | IMPORT_USER     | User who performed the import to production  |
    | AS4TEXT         | Transport description                        |
    | RC              | STMS return code (import result)             |
    +-----------------+----------------------------------------------+
    """

    TRANSPORT_RE = re.compile(r"[A-Z0-9]{3}K\d{6}")
    # Column names that signal a standard SAP table export (not a log)
    STANDARD_HEADER_HINTS = {"TRKORR", "AS4USER", "TRSTATUS", "AS4DATE"}

    # --- Public interface (matches TextFileReader / ExcelFileReader) ----------

    def read(self, file_path: Path) -> list[dict[str, Any]]:
        return [row for batch in self.read_in_batches(file_path) for row in batch]

    def read_in_batches(
        self, file_path: Path, chunk_size: int = 20000
    ) -> Iterator[list[dict[str, Any]]]:
        if self._is_stms_log(file_path):
            yield from self._stms_log_batches(file_path, chunk_size)
        else:
            # Standard SAP table export — delegate to the generic reader
            yield from TextFileReader().read_in_batches(file_path, chunk_size)

    # --- Format detection ----------------------------------------------------

    def _is_stms_log(self, file_path: Path) -> bool:
        """Return True when the file looks like an STMS import log.

        Detection heuristic (first 30 lines):
        - At least one line contains a transport-number pattern
        - None of the first lines contain a recognisable SAP table header
        """
        encoding = TextFileReader._detect_encoding(file_path)
        transport_hits = 0
        header_hits = 0

        with file_path.open("r", encoding=encoding, errors="replace") as handle:
            for _ in range(30):
                line = handle.readline()
                if not line:
                    break
                upper = line.upper()
                if self.TRANSPORT_RE.search(line):
                    transport_hits += 1
                if any(hint in upper for hint in self.STANDARD_HEADER_HINTS):
                    header_hits += 1

        return transport_hits > 0 and header_hits == 0

    # --- STMS log parser -----------------------------------------------------

    def _stms_log_batches(
        self, file_path: Path, chunk_size: int
    ) -> Iterator[list[dict[str, Any]]]:
        encoding = TextFileReader._detect_encoding(file_path)
        batch: list[dict[str, Any]] = []

        with file_path.open("r", encoding=encoding, errors="replace", newline="") as handle:
            for line in handle:
                row = self._parse_stms_line(line)
                if row is None:
                    continue
                batch.append(row)
                if len(batch) >= chunk_size:
                    yield batch
                    batch = []

        if batch:
            yield batch

    def _parse_stms_line(self, line: str) -> dict[str, Any] | None:
        """Parse one STMS log line into a normalised dict.

        Expected column order (tab-separated, positional):
        ``[seq]  [date]  [time]  [TRKORR]  [client]  [owner]  [import_user]  [description…]``
        """
        if not self.TRANSPORT_RE.search(line):
            return None

        parts = [p.strip() for p in line.split("\t") if p.strip()]
        if len(parts) < 7:
            return None

        # Locate the transport number — it may not always be at index 3
        trkorr_index = next(
            (i for i, p in enumerate(parts) if self.TRANSPORT_RE.fullmatch(p)), None
        )
        if trkorr_index is None:
            return None

        # Derive surrounding fields relative to the transport number position
        # בפורמט STMS הייצוא אינו תמיד עם כותרות, לכן אנחנו "מעגנים" את הפענוח סביב TRKORR
        # ומחשבים אינדקסים יחסיים קדימה/אחורה לפי המבנה השכיח.
        date_index = trkorr_index - 2 if trkorr_index >= 2 else None
        time_index = trkorr_index - 1 if trkorr_index >= 1 else None
        client_index = trkorr_index + 1 if trkorr_index + 1 < len(parts) else None
        owner_index = trkorr_index + 2 if trkorr_index + 2 < len(parts) else None
        import_user_index = trkorr_index + 3 if trkorr_index + 3 < len(parts) else None
        desc_start_index = trkorr_index + 4 if trkorr_index + 4 < len(parts) else None

        def _get(index: int | None) -> str:
            if index is None or index >= len(parts):
                return ""
            return parts[index]

        tail_parts: list[str] = []
        if desc_start_index is not None:
            tail_parts = parts[desc_start_index:]

        rc = ""
        description_parts = tail_parts
        if tail_parts:
            last_token = tail_parts[-1]
            # כאשר הטוקן האחרון מספרי בלבד, נחשב אותו כ-RC ונשאיר את התיאור נקי ממנו.
            if re.fullmatch(r"\d+", last_token):
                rc = last_token
                description_parts = tail_parts[:-1]

        return {
            "TRKORR": parts[trkorr_index],
            "AS4DATE": _get(date_index),
            "AS4TIME": _get(time_index),
            "MANDT": _get(client_index),
            "AS4USER": _get(owner_index),
            "IMPORT_USER": _get(import_user_index),
            "AS4TEXT": " ".join(description_parts).strip(),
            "RC": rc,
        }
