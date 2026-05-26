from __future__ import annotations

import copy
import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


class UiStateRepository:
    def __init__(self, output_dir: Path, input_dir: Path) -> None:
        self._output_dir = output_dir
        self._input_dir = input_dir

    def system_settings_path(self) -> Path:
        return self._output_dir / "system_settings.json"

    def file_dialog_state_path(self) -> Path:
        return self._output_dir / "file_dialog_state.json"

    def user_preview_settings_path(self) -> Path:
        return self._output_dir / "user_preview_columns.json"

    def user_reviewer_state_path(self) -> Path:
        return self._output_dir / "user_preview_reviewer_state.json"

    def load_system_settings(self, defaults: dict[str, Any]) -> dict[str, Any]:
        settings_path = self.system_settings_path()
        if not settings_path.exists():
            return copy.deepcopy(defaults)

        try:
            loaded = json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception:
            return copy.deepcopy(defaults)
        if not isinstance(loaded, dict):
            return copy.deepcopy(defaults)

        # תאימות לאחור בין שמות שדות ישנים לחדשים.
        if "generic_users" not in loaded and "critical_users" in loaded:
            loaded["generic_users"] = loaded.get("critical_users", [])
        if "authorized_stms_users" not in loaded and "super_users" in loaded:
            loaded["authorized_stms_users"] = loaded.get("super_users", [])
        if "super_users" not in loaded and "authorized_stms_users" in loaded:
            loaded["super_users"] = loaded.get("authorized_stms_users", [])

        merged = copy.deepcopy(defaults)
        for key, value in loaded.items():
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key].update(value)
            else:
                merged[key] = value
        return merged

    def save_system_settings(self, settings: dict[str, Any]) -> None:
        settings_path = self.system_settings_path()
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_last_file_dialog_directory(self, allow_persistence: bool) -> Path:
        default_directory = self._input_dir
        if not allow_persistence:
            return default_directory

        state_path = self.file_dialog_state_path()
        if not state_path.exists():
            return default_directory

        try:
            raw_data = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            return default_directory

        saved_directory = ""
        if isinstance(raw_data, dict):
            saved_directory = str(raw_data.get("last_directory", "")).strip()

        candidate_directory = Path(saved_directory).expanduser() if saved_directory else default_directory
        if candidate_directory.exists() and candidate_directory.is_dir():
            return candidate_directory
        return default_directory

    def save_last_file_dialog_directory(self, directory_path: object, allow_persistence: bool) -> Path | None:
        if directory_path is None:
            return None

        candidate_directory = Path(str(directory_path)).expanduser()
        if candidate_directory.is_file():
            candidate_directory = candidate_directory.parent
        if not candidate_directory.exists() or not candidate_directory.is_dir():
            return None

        if not allow_persistence:
            return candidate_directory

        state_path = self.file_dialog_state_path()
        payload = {"last_directory": str(candidate_directory)}
        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return candidate_directory

    def load_user_reviewer_state(
        self,
        allow_persistence: bool,
        normalize_status: Callable[[object], str],
    ) -> dict[str, dict[str, str]]:
        if not allow_persistence:
            return {}

        state_path = self.user_reviewer_state_path()
        if not state_path.exists():
            return {}

        try:
            raw_data = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

        if not isinstance(raw_data, dict):
            return {}

        normalized_state: dict[str, dict[str, str]] = {}
        for review_key, review_values in raw_data.items():
            if not isinstance(review_values, dict):
                continue
            legacy_notes = str(review_values.get("REVIEW_NOTES", "")).strip()
            normalized_state[str(review_key)] = {
                "REVIEW_STATUS": normalize_status(review_values.get("REVIEW_STATUS")),
                "TECH_REVIEW_NOTES": str(review_values.get("TECH_REVIEW_NOTES", "")).strip() or legacy_notes,
                "BUS_REVIEW_NOTES": str(review_values.get("BUS_REVIEW_NOTES", "")).strip(),
            }
        return normalized_state

    def save_user_reviewer_state(self, allow_persistence: bool, reviewer_state: dict[str, dict[str, str]]) -> None:
        if not allow_persistence:
            return
        state_path = self.user_reviewer_state_path()
        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_path.write_text(json.dumps(reviewer_state, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_user_preview_column_selection(
        self,
        allow_persistence: bool,
        default_columns: list[str],
        current_version: int,
        migrations: dict[int, list[str]],
        normalize_columns: Callable[[list[str] | None], list[str]],
    ) -> list[str]:
        if not allow_persistence:
            return list(default_columns)

        settings_path = self.user_preview_settings_path()
        if not settings_path.exists():
            return list(default_columns)

        try:
            raw_data = json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception:
            return list(default_columns)

        loaded_columns = list(raw_data.get("visible_columns", [])) if isinstance(raw_data, dict) else []
        settings_version = int(raw_data.get("version", 0)) if isinstance(raw_data, dict) else 0

        for version in range(settings_version + 1, current_version + 1):
            for field_name in migrations.get(version, []):
                if field_name not in loaded_columns:
                    loaded_columns.append(field_name)

        return normalize_columns(loaded_columns)

    def save_user_preview_column_selection(
        self,
        allow_persistence: bool,
        current_version: int,
        visible_columns: list[str],
    ) -> None:
        if not allow_persistence:
            return

        settings_path = self.user_preview_settings_path()
        payload = {
            "version": current_version,
            "visible_columns": visible_columns,
        }
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


class IpeEvidenceRepository:
    """Persists IPE (Information Produced by Entity) screenshot metadata and manages stored image files."""

    _EVIDENCE_JSON = "ipe_evidence.json"

    def __init__(self, output_dir: Path, base_dir: Path) -> None:
        self._output_dir = output_dir
        self._evidence_dir = base_dir / "data" / "evidence"

    def _json_path(self) -> Path:
        return self._output_dir / self._EVIDENCE_JSON

    def load(self) -> dict[str, list[dict[str, Any]]]:
        path = self._json_path()
        if not path.exists():
            return {}
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}
        if not isinstance(raw, dict):
            return {}
        return raw

    def save(self, data: dict[str, list[dict[str, Any]]]) -> None:
        path = self._json_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def add_image(
        self,
        slot_key: str,
        source_path: Path,
        control_ids: list[str],
        data: dict[str, list[dict[str, Any]]],
    ) -> dict[str, Any]:
        """Copy *source_path* into the evidence folder and append a new entry to *data* in-place.

        Returns the new entry dict.
        """
        slot_dir = self._evidence_dir / slot_key
        slot_dir.mkdir(parents=True, exist_ok=True)

        image_id = str(uuid.uuid4())
        dest_filename = f"{image_id}_{source_path.name}"
        dest_path = slot_dir / dest_filename
        shutil.copy2(source_path, dest_path)

        entry: dict[str, Any] = {
            "id": image_id,
            "original_filename": source_path.name,
            "stored_path": str(dest_path),
            "control_ids": list(control_ids),
            "added_at": datetime.now().isoformat(timespec="seconds"),
        }
        data.setdefault(slot_key, []).append(entry)
        self.save(data)
        return entry

    def remove_image(
        self,
        slot_key: str,
        image_id: str,
        data: dict[str, list[dict[str, Any]]],
    ) -> None:
        """Remove the image entry from *data* and delete the stored file.  Saves afterwards."""
        entries = data.get(slot_key, [])
        to_remove = next((e for e in entries if e.get("id") == image_id), None)
        if to_remove is None:
            return
        stored = Path(to_remove.get("stored_path", ""))
        if stored.exists():
            try:
                stored.unlink()
            except OSError:
                pass
        data[slot_key] = [e for e in entries if e.get("id") != image_id]
        self.save(data)
