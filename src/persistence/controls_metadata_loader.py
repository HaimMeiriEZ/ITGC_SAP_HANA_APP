"""Loader for control metadata CSV (filled by user).

Reads `data/output/controls_metadata.csv` and produces a dict keyed by control_id with
- description (override)
- process
- risk_description
- extra_notes
The fields are intended to enrich `AUDIT_CONTROL_DEFINITIONS` at application startup.

Column layout (Hebrew headers):
    מזהה בקרה, קטגוריה, רמת סיכון, סוג בדיקה,
    תיאור הבקרה (קיים),
    תהליך (למילוי),
    תיאור הסיכון (למילוי),
    צעדי טסט (אופציונלי - אם ריק יילקח טקסט גנרי),
    תיעוד נדרש (אופציונלי),     <-- ignored on purpose (we derive from profile)
    הערות נוספות (אופציונלי)
"""
from __future__ import annotations

import csv
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

CSV_FILENAME = "controls_metadata.csv"

_COL_ID = "מזהה בקרה"
_COL_DESC = "תיאור הבקרה (קיים)"
_COL_PROCESS = "תהליך (למילוי)"
_COL_RISK = "תיאור הסיכון (למילוי)"
_COL_TEST_STEPS = "צעדי טסט (אופציונלי - אם ריק יילקח טקסט גנרי)"
_COL_NOTES = "הערות נוספות (אופציונלי)"


def _clean(value: Any) -> str:
    return str(value or "").strip()


def load_controls_metadata_csv(output_dir: Path) -> dict[str, dict[str, str]]:
    """Read the controls metadata CSV. Returns dict[control_id, fields].

    Returns an empty dict if the file is missing or cannot be parsed.
    """
    csv_path = Path(output_dir) / CSV_FILENAME
    if not csv_path.exists():
        logger.warning("controls_metadata.csv not found at %s; skipping metadata override", csv_path)
        return {}

    result: dict[str, dict[str, str]] = {}
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                control_id = _clean(row.get(_COL_ID))
                if not control_id:
                    continue
                entry: dict[str, str] = {}
                desc = _clean(row.get(_COL_DESC))
                if desc:
                    entry["description"] = desc
                process = _clean(row.get(_COL_PROCESS))
                if process:
                    entry["process"] = process
                risk = _clean(row.get(_COL_RISK))
                if risk:
                    entry["risk_description"] = risk
                test_steps = _clean(row.get(_COL_TEST_STEPS))
                if test_steps:
                    entry["test_steps_override"] = test_steps
                notes = _clean(row.get(_COL_NOTES))
                if notes:
                    entry["extra_notes"] = notes
                if entry:
                    result[control_id] = entry
    except Exception as exc:  # pragma: no cover - defensive
        logger.exception("Failed to read controls metadata CSV %s: %s", csv_path, exc)
        return {}

    logger.info("Loaded controls metadata for %d controls from %s", len(result), csv_path)
    return result


def apply_metadata_to_definitions(
    metadata: dict[str, dict[str, str]],
    definitions: dict[str, dict[str, str]],
) -> None:
    """Merge metadata into AUDIT_CONTROL_DEFINITIONS in-place.

    - description: OVERWRITES existing value (per user decision)
    - process / risk_description / extra_notes / test_steps_override: added as new keys
    """
    for control_id, fields in metadata.items():
        target = definitions.get(control_id)
        if target is None:
            # Skip unknown ids silently; logged once
            logger.debug("CSV metadata references unknown control_id %s", control_id)
            continue
        for key, value in fields.items():
            if not value:
                continue
            if key == "description":
                target["description"] = value
            else:
                target[key] = value
