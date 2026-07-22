"""Persist compensating-control evidence files (one file per control)."""
from __future__ import annotations

import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any


class CompensatingControlRepository:
    _JSON_FILENAME = "compensating_controls.json"

    def __init__(self, output_dir: Path, base_dir: Path) -> None:
        self._output_dir = output_dir
        self._storage_dir = base_dir / "data" / "compensating_controls"

    def _json_path(self) -> Path:
        return self._output_dir / self._JSON_FILENAME

    def load(self) -> dict[str, dict[str, Any]]:
        path = self._json_path()
        if not path.exists():
            return {}
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}
        if not isinstance(raw, dict):
            return {}
        return {
            str(control_id): dict(entry)
            for control_id, entry in raw.items()
            if isinstance(entry, dict) and control_id
        }

    def save(self, data: dict[str, dict[str, Any]]) -> None:
        path = self._json_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def _delete_stored_file(self, entry: dict[str, Any] | None) -> None:
        if not entry:
            return
        stored = Path(str(entry.get("stored_path", "")))
        if stored.exists():
            try:
                stored.unlink()
            except OSError:
                pass

    def attach_file(
        self,
        control_id: str,
        source_path: Path,
        data: dict[str, dict[str, Any]],
    ) -> dict[str, Any]:
        """Copy *source_path* into storage and replace any prior file for *control_id*."""
        control_key = str(control_id).strip()
        if not control_key:
            raise ValueError("control_id is required")

        self._delete_stored_file(data.get(control_key))

        control_dir = self._storage_dir / control_key
        control_dir.mkdir(parents=True, exist_ok=True)

        file_id = str(uuid.uuid4())
        dest_filename = f"{file_id}_{source_path.name}"
        dest_path = control_dir / dest_filename
        shutil.copy2(source_path, dest_path)

        entry: dict[str, Any] = {
            "control_id": control_key,
            "original_filename": source_path.name,
            "stored_path": str(dest_path),
            "added_at": datetime.now().isoformat(timespec="seconds"),
        }
        data[control_key] = entry
        self.save(data)
        return entry

    def remove_file(self, control_id: str, data: dict[str, dict[str, Any]]) -> None:
        control_key = str(control_id).strip()
        if not control_key:
            return
        self._delete_stored_file(data.get(control_key))
        data.pop(control_key, None)
        self.save(data)
