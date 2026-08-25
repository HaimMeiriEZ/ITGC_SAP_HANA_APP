# Deployment — אריזת EXE (Windows)

## בניה

מתוך שורש הפרויקט (עם `.venv` פעיל מומלץ):

```powershell
.\scripts\build_exe.ps1
```

התוצאה: `dist/ITGC_SAP_APP/` — תיקייה להעתקה ללקוח (onedir, בסביבות ~150MB).

הפעלה אצל הלקוח: `ITGC_SAP_APP.exe`.

## מה נכלל בחבילה

- קוד האפליקציה + PySide6
- `data/knowledge_base/` (קטלוג בקרות + תוויות שדות)
- `data/output/system_settings.json` **אם קיים** בפרויקט בעת ה־build (הגדרות ביקורת ללקוח)

## מה לא נכלל

- `data/evidence/` (IPE)
- `data/compensating_controls/`
- ממצאים, ניירות עבודה, מצב סקירת משתמשים, לוגים, `file_dialog_state.json`

בהרצה ראשונה הכלי יוצר תיקיות `data/` ריקות לפי הצורך ומשכפל את ה־knowledge_base אם חסר.

## דרישות במחשב הלקוח

- Windows (x64)
- להטמעת OLE בנייר עבודה: Microsoft Excel
- לטיוטות מייל: Microsoft Outlook
- הרשאות כתיבה לתיקיית ההתקנה (תת־תיקיית `data/`)

## עדכון גרסה

העתיקו את תיקיית `dist/ITGC_SAP_APP` החדשה, ושמרו את `data/output` וקבצי הלקוח מהגרסה הקודמת (במיוחד `system_settings.json` ותיעוד שהועלה).
