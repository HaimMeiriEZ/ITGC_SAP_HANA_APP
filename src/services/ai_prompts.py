"""Hebrew prompt templates for all AI features.

Each constant is a template string.  Use str.format(**kwargs) or
str.format_map(mapping) to fill in the placeholders before sending to
the LLM.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# System prompt shared by all ITGC features
# ---------------------------------------------------------------------------

SYSTEM_ITGC_AUDITOR = (
    "אתה מומחה לביקורת IT ו-ITGC בסביבות SAP. "
    "אתה כותב בעברית בסגנון מקצועי-ביקורתי, תמציתי ומדויק. "
    "הימנע מפרשנות יצירתית — הסתמך רק על הנתונים שסופקו. "
    "אל תמציא פרטים שאינם בנתונים."
)

# ---------------------------------------------------------------------------
# Feature 2.1 — Findings narration (user level)
# ---------------------------------------------------------------------------

USER_FINDINGS_NARRATION = """\
להלן נתוני משתמש SAP מערכת {work_environment}:
- משתמש: {bname}
- סוג משתמש: {ustyp}
- סטטוס: {status}
- כניסה אחרונה: {trdat}
- תאריך תפוגת תוקף: {gltgb}
- מדיניות אבטחה: {security_policy}
- ממצאים טכניים שזוהו: {raw_findings}

נסח פסקה אחת (2-3 משפטים) המסכמת את ממצאי הסיכון בשפה מקצועית-ביקורתית בעברית,
המתאימה לנייר עבודה של מבקר. התמקד בסיכון העסקי ולא בנתונים הטכניים גרדם.
"""

# ---------------------------------------------------------------------------
# Feature 2.1 — Findings narration (audit control level)
# ---------------------------------------------------------------------------

AUDIT_FINDINGS_NARRATION = """\
להלן פרטי בקרת ITGC שנכשלה:
- מזהה בקרה: {control_id}
- תיאור בקרה: {control_description}
- סביבת עבודה: {work_environment}
- ערך בפועל: {actual_value}
- ערך מצופה: {expected_value}
- ממצא גולמי: {raw_finding}

נסח פסקה אחת (2-3 משפטים) המסכמת את הממצא בשפה ביקורתית-מקצועית בעברית,
מתאימה לנייר עבודה. ציין מה הסיכון העסקי ומהי ההשפעה האפשרית.
"""

# ---------------------------------------------------------------------------
# Feature 2.2 — Compensating controls advisor
# ---------------------------------------------------------------------------

COMPENSATING_CONTROLS_ADVISOR = """\
בקרת ITGC הבאה נכשלה:
- מזהה בקרה: {control_id}
- תיאור: {control_description}
- רמת סיכון: {risk_level}
- מספר רשומות עם ממצא: {finding_count}
- סביבת עבודה: {work_environment}

מועמדי בקרות מפצות מהמאגר:
{kb_candidates_json}

בחר עד 3 בקרות מפצות המתאימות ביותר לממצא זה, מדרג לפי עדיפות.
ענה אך ורק ב-JSON בפורמט הבא (מערך):
[
  {{
    "rank": 1,
    "title": "שם הבקרה המפצה בעברית",
    "rationale": "הסבר קצר מדוע היא מתאימה לממצא זה",
    "evidence_needed": "מה צריך לאסוף כעדות",
    "frameworks": ["SOX", "COBIT"]
  }}
]
"""
