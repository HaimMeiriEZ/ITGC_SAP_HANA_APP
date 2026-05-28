"""Statistical anomaly detection for SAP USR02 user population.

No LLM is used here — all logic is deterministic and explainable to auditors.
Each finding carries a Hebrew explanation string suitable for working-paper use.

Main entry-point:
    detector = UserAnomalyDetector(settings)
    cohort = detector.build_cohort_stats(all_rows)
    for row in all_rows:
        findings = detector.score_user(row, cohort)
        # findings is list[AnomalyFinding]
"""
from __future__ import annotations

import math
import re
import statistics
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class AnomalyFinding:
    code: str          # Short machine-readable code, e.g. "GENERIC_NAME"
    severity: str      # "low" | "medium" | "high"
    explanation_he: str  # Hebrew explanation for working-paper


@dataclass
class CohortStats:
    """Per-MANDT × USTYP statistics needed for Z-score outlier detection."""
    # Maps (mandt, ustyp) -> list of LOCNT values from that cohort
    locnt_by_cohort: dict[tuple[str, str], list[float]] = field(default_factory=dict)
    # Maps (mandt, ustyp) -> (mean, stdev)
    locnt_stats: dict[tuple[str, str], tuple[float, float]] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Regex patterns for generic / suspicious user names
# ---------------------------------------------------------------------------

_GENERIC_NAME_PATTERNS: list[re.Pattern] = [
    re.compile(r"^(test|temp|tmp|demo|dummy|user|generic|admin|service|sap|system)\d*$", re.IGNORECASE),
    re.compile(r"^(user|usr|u)\d{1,5}$", re.IGNORECASE),
    re.compile(r"^\d+$"),          # pure digits
    re.compile(r"^[a-z]{1,3}$", re.IGNORECASE),  # very short alphabetic (≤3 chars)
]


def _parse_date(value: object) -> date | None:
    """Try to parse common SAP date formats; return None on failure."""
    if not value:
        return None
    s = str(value).strip().replace("/", "-")
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _days_since(d: date | None, reference: date | None = None) -> int | None:
    if d is None:
        return None
    ref = reference or date.today()
    return (ref - d).days


# ---------------------------------------------------------------------------
# Main detector class
# ---------------------------------------------------------------------------

class UserAnomalyDetector:
    def __init__(self, settings: dict[str, Any] | None = None) -> None:
        cfg = settings or {}
        self._inactive_threshold: int = int(cfg.get("inactive_days_threshold", 90))
        self._locnt_zscore_threshold: float = float(cfg.get("anomaly_locnt_zscore", 2.5))
        self._enabled: bool = bool(cfg.get("anomaly_detection_enabled", True))

    # ------------------------------------------------------------------
    # Step 1: build population statistics (call once per preview load)
    # ------------------------------------------------------------------

    def build_cohort_stats(self, rows: list[dict[str, Any]]) -> CohortStats:
        stats = CohortStats()
        for row in rows:
            mandt = str(row.get("MANDT", "")).strip()
            ustyp = str(row.get("USTYP", "")).strip()
            key = (mandt, ustyp)
            try:
                locnt_val = float(str(row.get("LOCNT", 0) or 0))
            except (ValueError, TypeError):
                locnt_val = 0.0
            stats.locnt_by_cohort.setdefault(key, []).append(locnt_val)

        for key, values in stats.locnt_by_cohort.items():
            mean = statistics.mean(values) if values else 0.0
            stdev = statistics.stdev(values) if len(values) >= 2 else 0.0
            stats.locnt_stats[key] = (mean, stdev)

        return stats

    # ------------------------------------------------------------------
    # Step 2: score a single user row
    # ------------------------------------------------------------------

    def score_user(
        self,
        row: dict[str, Any],
        cohort: CohortStats,
        review_period_end: date | None = None,
    ) -> list[AnomalyFinding]:
        if not self._enabled:
            return []
        findings: list[AnomalyFinding] = []
        findings.extend(self._detect_generic_name(row))
        findings.extend(self._detect_temporal(row, review_period_end))
        findings.extend(self._detect_locnt_outlier(row, cohort))
        return findings

    # ------------------------------------------------------------------
    # Individual detectors
    # ------------------------------------------------------------------

    def _detect_generic_name(self, row: dict[str, Any]) -> list[AnomalyFinding]:
        bname = str(row.get("BNAME", "")).strip()
        name_textc = str(row.get("NAME_TEXTC", "")).strip()

        for pattern in _GENERIC_NAME_PATTERNS:
            if pattern.match(bname):
                return [AnomalyFinding(
                    code="GENERIC_NAME",
                    severity="medium",
                    explanation_he=(
                        f"[חריגה סטטיסטית] שם משתמש '{bname}' תואם דפוס של משתמש גנרי "
                        f"(כגון test/temp/admin). יש לבדוק האם מדובר במשתמש אישי מזוהה."
                    ),
                )]

        # Short bname with no full name and with login activity
        trdat = _parse_date(row.get("TRDAT"))
        if len(bname) <= 3 and not name_textc and trdat is not None:
            return [AnomalyFinding(
                code="SHORT_NAME_ACTIVE",
                severity="low",
                explanation_he=(
                    f"[חריגה סטטיסטית] שם משתמש '{bname}' קצר מאוד (עד 3 תווים) "
                    f"ואינו משויך לשם מלא, אך רשומה כניסה לאחרונה. "
                    f"יש לוודא שמדובר בחשבון מורשה."
                ),
            )]

        return []

    def _detect_temporal(
        self, row: dict[str, Any], review_period_end: date | None
    ) -> list[AnomalyFinding]:
        findings: list[AnomalyFinding] = []
        reference = review_period_end or date.today()
        bname = str(row.get("BNAME", "")).strip()
        ustyp = str(row.get("USTYP", "")).strip()

        trdat = _parse_date(row.get("TRDAT"))
        gltgb = _parse_date(row.get("GLTGB"))
        pwdchgdate = _parse_date(row.get("PWDCHGDATE"))
        pwdsetdate = _parse_date(row.get("PWDSETDATE"))

        # Login after account expiry
        if trdat and gltgb and trdat > gltgb:
            findings.append(AnomalyFinding(
                code="LOGIN_AFTER_EXPIRY",
                severity="high",
                explanation_he=(
                    f"[חריגה סטטיסטית] משתמש '{bname}' נכנס לאחרונה בתאריך {trdat} "
                    f"לאחר תאריך תפוגת התוקף {gltgb}. זהו ממצא קריטי המצריך בדיקה מיידית."
                ),
            ))

        # Long inactivity for Dialog users
        if ustyp == "A" and trdat:
            days_inactive = _days_since(trdat, reference)
            if days_inactive is not None and days_inactive > self._inactive_threshold:
                findings.append(AnomalyFinding(
                    code="LONG_INACTIVE_DIALOG",
                    severity="medium",
                    explanation_he=(
                        f"[חריגה סטטיסטית] משתמש Dialog '{bname}' לא התחבר במשך "
                        f"{days_inactive} ימים (מעל הסף של {self._inactive_threshold} יום). "
                        f"יש לשקול נעילה או מחיקת החשבון."
                    ),
                ))

        # Password change date before password set date (data anomaly)
        if pwdchgdate and pwdsetdate and pwdchgdate < pwdsetdate:
            findings.append(AnomalyFinding(
                code="PWD_DATE_INCONSISTENCY",
                severity="low",
                explanation_he=(
                    f"[חריגה סטטיסטית] עבור משתמש '{bname}': תאריך שינוי סיסמה "
                    f"({pwdchgdate}) קודם לתאריך הגדרת סיסמה ({pwdsetdate}). "
                    f"ייתכן שמדובר באי-עקביות בנתוני המקור."
                ),
            ))

        return findings

    def _detect_locnt_outlier(
        self, row: dict[str, Any], cohort: CohortStats
    ) -> list[AnomalyFinding]:
        mandt = str(row.get("MANDT", "")).strip()
        ustyp = str(row.get("USTYP", "")).strip()
        bname = str(row.get("BNAME", "")).strip()

        try:
            locnt = float(str(row.get("LOCNT", 0) or 0))
        except (ValueError, TypeError):
            return []

        if locnt <= 0:
            return []

        key = (mandt, ustyp)
        mean, stdev = cohort.locnt_stats.get(key, (0.0, 0.0))

        if stdev < 0.001:
            # All users in this cohort have the same LOCNT — no meaningful z-score
            return []

        z = (locnt - mean) / stdev
        if z < self._locnt_zscore_threshold:
            return []

        return [AnomalyFinding(
            code="HIGH_FAILED_LOGINS",
            severity="high",
            explanation_he=(
                f"[חריגה סטטיסטית] משתמש '{bname}' מציג {int(locnt)} ניסיונות כניסה "
                f"כושלים — חריגה של {z:.1f} סטיות תקן מעל ממוצע הקבוצה "
                f"(מ={mean:.1f}). ייתכן ניסיון פריצה או נעילת חשבון חוזרת."
            ),
        )]


# ---------------------------------------------------------------------------
# Helpers consumed by desktop_app / user_preview_service
# ---------------------------------------------------------------------------

def anomaly_findings_to_text(findings: list[AnomalyFinding]) -> str:
    """Concatenate Hebrew explanation strings separated by ' | '."""
    return " | ".join(f.explanation_he for f in findings)


def anomaly_score(findings: list[AnomalyFinding]) -> int:
    """Map a list of findings to a 0-100 severity score."""
    if not findings:
        return 0
    weights = {"low": 10, "medium": 40, "high": 80}
    total = sum(weights.get(f.severity, 0) for f in findings)
    return min(total, 100)


def anomaly_codes(findings: list[AnomalyFinding]) -> str:
    """Comma-separated list of finding codes."""
    return ", ".join(f.code for f in findings)
