"""OllamaClient — local LLM integration for ITGC AI features.

All AI calls are local-only (Ollama). No data leaves the machine.
PII fields (PASSCODE, PWDSALTEDHASH, OCOD1) are stripped before any prompt.
Every public method degrades gracefully when Ollama is unavailable.
"""
from __future__ import annotations

import json
import logging
from typing import Any

logger = logging.getLogger(__name__)

# Fields that must never appear in any LLM prompt
_PII_FIELDS = frozenset({"PASSCODE", "PWDSALTEDHASH", "OCOD1", "USRPWD"})

# Recommended Ollama models ordered by Hebrew quality
RECOMMENDED_MODELS = [
    "aya-expanse:8b",
    "gemma2:9b",
    "llama3.1:8b-instruct",
    "qwen2.5:7b-instruct",
    "mistral:7b-instruct",
]

_DEFAULT_HOST = "http://localhost:11434"
_DEFAULT_MODEL = "aya-expanse:8b"
_DEFAULT_TEMPERATURE = 0.3
_DEFAULT_TIMEOUT = 60


def redact_user_payload(row: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of *row* with all PII fields removed."""
    return {k: v for k, v in row.items() if k.upper() not in _PII_FIELDS}


class OllamaClient:
    """Thin synchronous wrapper around the Ollama /api/generate endpoint.

    Reads configuration from the *ai_settings* dict passed at construction.
    Falls back to module-level defaults for any missing keys.
    """

    def __init__(self, ai_settings: dict[str, Any] | None = None) -> None:
        cfg = ai_settings or {}
        self._host: str = str(cfg.get("ollama_host", _DEFAULT_HOST)).rstrip("/")
        self._model: str = str(cfg.get("model", _DEFAULT_MODEL))
        self._temperature: float = float(cfg.get("temperature", _DEFAULT_TEMPERATURE))
        self._timeout: float = float(cfg.get("timeout_seconds", _DEFAULT_TIMEOUT))

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def is_available(self) -> bool:
        """Return True if the Ollama server responds on /api/tags."""
        try:
            import httpx
            with httpx.Client(timeout=5) as client:
                resp = client.get(f"{self._host}/api/tags")
                return resp.status_code == 200
        except Exception:
            return False

    def list_local_models(self) -> list[str]:
        """Return names of models already pulled in the local Ollama instance."""
        try:
            import httpx
            with httpx.Client(timeout=5) as client:
                resp = client.get(f"{self._host}/api/tags")
                if resp.status_code != 200:
                    return []
                data = resp.json()
                return [m.get("name", "") for m in data.get("models", []) if m.get("name")]
        except Exception:
            return []

    def generate(
        self,
        prompt: str,
        system: str = "",
        temperature: float | None = None,
        max_tokens: int = 512,
        timeout: float | None = None,
    ) -> str:
        """Send a generation request to Ollama and return the response text.

        Returns an empty string on any error so callers never crash.
        """
        try:
            import httpx
            payload: dict[str, Any] = {
                "model": self._model,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": temperature if temperature is not None else self._temperature,
                    "num_predict": max_tokens,
                },
            }
            if system:
                payload["system"] = system

            effective_timeout = timeout if timeout is not None else self._timeout
            with httpx.Client(timeout=effective_timeout) as client:
                resp = client.post(f"{self._host}/api/generate", json=payload)
                resp.raise_for_status()
                data = resp.json()
                return str(data.get("response", "")).strip()
        except Exception as exc:
            logger.debug("Ollama generate failed: %s", exc)
            return ""

    def generate_json(
        self,
        prompt: str,
        system: str = "",
        temperature: float | None = None,
        max_tokens: int = 768,
        timeout: float | None = None,
    ) -> dict[str, Any] | list[Any] | None:
        """Like *generate* but expects JSON output; returns None on parse failure."""
        raw = self.generate(prompt, system=system, temperature=temperature,
                            max_tokens=max_tokens, timeout=timeout)
        if not raw:
            return None
        # Strip markdown code fences if the model wrapped the JSON
        cleaned = raw.strip()
        if cleaned.startswith("```"):
            lines = cleaned.splitlines()
            cleaned = "\n".join(
                line for line in lines
                if not line.strip().startswith("```")
            ).strip()
        try:
            return json.loads(cleaned)
        except Exception:
            logger.debug("JSON parse failed for Ollama response: %s", cleaned[:200])
            return None
