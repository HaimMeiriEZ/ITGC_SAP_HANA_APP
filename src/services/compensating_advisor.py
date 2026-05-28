"""Compensating controls advisor — combines KB filtering with LLM ranking.

Usage:
    advisor = CompensatingAdvisor(client, work_environment="FPP")
    recommendations = advisor.recommend(
        control_id="MA3-3",
        control_description="...",
        risk_level="high",
        finding_count=5,
    )
    # recommendations: list[dict] with keys rank, title, rationale, evidence_needed, frameworks
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from .ai_service import OllamaClient
from .ai_prompts import SYSTEM_ITGC_AUDITOR, COMPENSATING_CONTROLS_ADVISOR
from .compensating_kb_loader import CompensatingKBLoader

logger = logging.getLogger(__name__)

_CACHE_FILE = (
    Path(__file__).parent.parent.parent / "data" / "output" / "ai_compensating_cache.json"
)


def _load_cache() -> dict[str, list[dict]]:
    try:
        if _CACHE_FILE.exists():
            return json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_cache(cache: dict[str, list[dict]]) -> None:
    try:
        _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        _CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.debug("Could not save compensating cache: %s", exc)


def _cache_key(control_id: str, risk_level: str, finding_count: int) -> str:
    return f"{control_id}|{risk_level}|{finding_count}"


class CompensatingAdvisor:
    def __init__(
        self,
        client: OllamaClient,
        work_environment: str = "",
        kb_path: Path | str | None = None,
    ) -> None:
        self._client = client
        self._work_environment = work_environment
        self._kb_path = kb_path

    def recommend(
        self,
        control_id: str,
        control_description: str,
        risk_level: str = "high",
        finding_count: int = 0,
    ) -> list[dict[str, Any]]:
        """Return up to 3 ranked compensating control recommendations.

        Returns plain KB entries (no LLM ranking) if Ollama is unavailable.
        Returns [] if no KB matches exist.
        """
        candidates = CompensatingKBLoader.filter_for_control(
            control_id, risk_level=risk_level, kb_path=self._kb_path
        )
        if not candidates:
            return []

        key = _cache_key(control_id, risk_level, finding_count)
        cache = _load_cache()
        if key in cache:
            return cache[key]

        # Build minimal JSON representation for the prompt (avoid huge payloads)
        kb_json = json.dumps(
            [
                {
                    "id": e.get("id", ""),
                    "title_he": e.get("title_he", ""),
                    "description_he": e.get("description_he", "")[:200],
                    "frameworks": e.get("frameworks", []),
                }
                for e in candidates
            ],
            ensure_ascii=False,
            indent=2,
        )

        prompt = COMPENSATING_CONTROLS_ADVISOR.format(
            control_id=control_id,
            control_description=control_description,
            risk_level=risk_level,
            finding_count=finding_count,
            work_environment=self._work_environment or "SAP",
            kb_candidates_json=kb_json,
        )

        result = self._client.generate_json(prompt, system=SYSTEM_ITGC_AUDITOR, max_tokens=768)

        if not isinstance(result, list):
            # LLM failed or returned non-list — fall back to KB order
            result = [
                {
                    "rank": i + 1,
                    "title": e.get("title_he", ""),
                    "rationale": e.get("description_he", ""),
                    "evidence_needed": "; ".join(e.get("evidence_needed", [])),
                    "frameworks": e.get("frameworks", []),
                }
                for i, e in enumerate(candidates[:3])
            ]

        # Ensure at most 3
        result = result[:3]
        cache[key] = result
        _save_cache(cache)
        return result
