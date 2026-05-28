"""Loads and caches the compensating-controls YAML knowledge base."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

_DEFAULT_KB_PATH = (
    Path(__file__).parent.parent.parent / "data" / "knowledge_base" / "compensating_controls.yaml"
)


class CompensatingKBLoader:
    """Singleton-style loader.  Call ``get()`` to obtain the entry list."""

    _cache: list[dict[str, Any]] | None = None
    _loaded_path: Path | None = None

    @classmethod
    def get(cls, kb_path: Path | str | None = None) -> list[dict[str, Any]]:
        """Return list of compensating control dicts, loading from YAML if needed."""
        path = Path(kb_path) if kb_path else _DEFAULT_KB_PATH
        if cls._cache is not None and cls._loaded_path == path:
            return cls._cache

        try:
            import yaml  # lazy import — pyyaml optional at runtime
            with path.open(encoding="utf-8") as fh:
                data = yaml.safe_load(fh)
            cls._cache = data.get("compensating_controls", [])
            cls._loaded_path = path
            return cls._cache
        except Exception as exc:
            logger.warning("Could not load compensating controls KB from %s: %s", path, exc)
            return []

    @classmethod
    def filter_for_control(
        cls,
        control_id: str,
        risk_level: str = "high",
        kb_path: Path | str | None = None,
    ) -> list[dict[str, Any]]:
        """Return KB entries that apply to *control_id* and *risk_level*."""
        entries = cls.get(kb_path)
        results: list[dict[str, Any]] = []
        for entry in entries:
            applies_when: dict = entry.get("applies_when", {})
            ids: list[str] = applies_when.get("control_ids", [])
            risk_levels: list[str] = applies_when.get("risk_levels", [])

            # Match if control_id or its prefix matches any entry's applies_when list
            id_match = any(
                control_id.upper().startswith(allowed.upper()) or
                allowed.upper().startswith(control_id.upper())
                for allowed in ids
            )
            risk_match = not risk_levels or risk_level.lower() in [r.lower() for r in risk_levels]

            if id_match and risk_match:
                results.append(entry)
        return results
