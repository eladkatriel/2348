# Monday → Dropbox Report Automation (MVP)

## מה זה עושה
- מקבל webhook מ-Monday
- יוצר תיקייה בדרופבוקס
- מייצר דו"ח Word מטמפלייט
- ממלא placeholders

## איך להריץ

### 1. התקנת ספריות
pip install -r requirements.txt

### 2. הגדרת משתני סביבה
MONDAY_API_KEY=...
DROPBOX_TOKEN=...
BOARD_ID=...

### 3. הרצה
python app.py

## Deploy מומלץ
Render / Railway

## מבנה דרופבוקס
/Templates/Report_Template.docx
/Projects/

## Placeholders
{{Engineer}}
{{Street}}
{{Number}}
{{Apartment}}
{{Hit date}}
{{Date}}
{{months}}