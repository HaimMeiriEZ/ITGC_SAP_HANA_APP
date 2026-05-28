"""Narrates ITGC audit findings in Hebrew using a local Ollama LLM.

Fail-open design: if the LLM is unavailable or returns garbage, the
original rule-based text is returned unchanged.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
from pathlib import Path
from typing import Any

from .ai_service import OllamaClient, redact_user_payload
from .ai_prompts import SYSTEM_ITGC_AUDITOR, USER_FINDINGS_NARRATION, AUDIT_FINDINGS_NARRATION

logger = logging.getLogger(__name__)

_CACHE_FILE = Path(__file__).parent.parent.parent / "data" / "output" / "ai_narrations.json"


def _load_cache() -> dict[str, str]:
    try:
        if _CACHE_FILE.exists():
            return json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_cache(cache: dict[str, str]) -> None:
    try:
        _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        _CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.debug("Could not save narration cache: %s", exc)


def _cache_key(bname: str, mandt: str, raw_findings: str) -> str:
    h = hashlib.sha256(raw_findings.encode()).hexdigest()[:16]
    return f"{mandt}|{bname}|{h}"


def narrate_user_finding(
    row: dict[str, Any],
    raw_findings: str,
    client: OllamaClient,
    work_environment: str = "",
) -> str:
    """Return a Hebrew narration of the user-level finding.

    Falls back to *raw_findings* if LLM fails.
    Results are cached to avoid redundant inference.
    """
    if not raw_findings.strip():
        return raw_findings

    bname = str(row.get("BNAME", "")).strip()
    mandt = str(row.get("MANDT", "")).strip()
    key = _cache_key(bname, mandt, raw_findings)

    cache = _load_cache()
    if key in cache:
        return cache[key]

    safe_row = redact_user_payload(row)
    prompt = USER_FINDINGS_NARRATION.format(
        work_environment=work_environment or "SAP",
        bname=safe_row.get("BNAME", ""),
        ustyp=safe_row.get("USTYP", ""),
        status=safe_row.get("USTAT", ""),
        trdat=safe_row.get("TRDAT", ""),
        gltgb=safe_row.get("GLTGB", ""),
        security_policy=safe_row.get("SECURITY_POLICY", ""),
        raw_findings=raw_findings,
    )

    result = client.generate(prompt, system=SYSTEM_ITGC_AUDITOR, max_tokens=512)
    if not result:
        return raw_findings

    cache[key] = result
    _save_cache(cache)
    return result


def narrate_audit_finding(
    control_id: str,
    control_description: str,
    actual_value: str,
    expected_value: str,
    raw_finding: str,
    client: OllamaClient,
    work_environment: str = "",
) -> str:
    """Return a Hebrew narration of an audit-control-level finding.

    Falls back to *raw_finding* if LLM fails.
    """
    if not raw_finding.strip():
        return raw_finding

    key = _cache_key(control_id, work_environment, raw_finding)
    cache = _load_cache()
    if key in cache:
        return cache[key]

    prompt = AUDIT_FINDINGS_NARRATION.format(
        control_id=control_id,
        control_description=control_description,
        work_environment=work_environment or "SAP",
        actual_value=actual_value,
        expected_value=expected_value,
        raw_finding=raw_finding,
    )

    result = client.generate(prompt, system=SYSTEM_ITGC_AUDITOR, max_tokens=512)
    if not result:
        return raw_finding

    cache[key] = result
    _save_cache(cache)
    return result
