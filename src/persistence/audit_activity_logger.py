from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path


class UserReviewActivityLogger:
    """Append-only CSV logger for user-review activity events."""

    FIELD_NAMES = [
        "timestamp",
        "actor",
        "action_type",
        "source",
        "review_key",
        "mandt",
        "bname",
        "field_name",
        "old_value",
        "new_value",
        "import_file",
        "import_mode",
        "warning_code",
        "warning_details",
        "session_id",
    ]

    def __init__(self, log_path: Path) -> None:
        self._log_path = log_path

    def append_event(
        self,
        *,
        actor: str,
        action_type: str,
        source: str,
        review_key: str = "",
        mandt: str = "",
        bname: str = "",
        field_name: str = "",
        old_value: object = "",
        new_value: object = "",
        import_file: str = "",
        import_mode: str = "",
        warning_code: str = "",
        warning_details: str = "",
        session_id: str = "",
    ) -> None:
        self._log_path.parent.mkdir(parents=True, exist_ok=True)
        write_header = not self._log_path.exists() or self._log_path.stat().st_size == 0

        row = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "actor": self._normalize_value(actor),
            "action_type": self._normalize_value(action_type),
            "source": self._normalize_value(source),
            "review_key": self._normalize_value(review_key),
            "mandt": self._normalize_value(mandt),
            "bname": self._normalize_value(bname),
            "field_name": self._normalize_value(field_name),
            "old_value": self._normalize_value(old_value),
            "new_value": self._normalize_value(new_value),
            "import_file": self._normalize_value(import_file),
            "import_mode": self._normalize_value(import_mode),
            "warning_code": self._normalize_value(warning_code),
            "warning_details": self._normalize_value(warning_details),
            "session_id": self._normalize_value(session_id),
        }

        with self._log_path.open("a", encoding="utf-8-sig", newline="") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=self.FIELD_NAMES)
            if write_header:
                writer.writeheader()
            writer.writerow(row)

    @staticmethod
    def _normalize_value(value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()
