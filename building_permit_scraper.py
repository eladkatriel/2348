"""
building_permit_scraper.py
===========================
מוריד אוטומטית PDF של תיק בניין מאתר העירייה הרלוונטית.

תומך ב:
  - קומפלוט (COMPLOT) — ~30 עיריות
  - Bartech (ערד)
  - תל אביב (TAVAPP)

שימוש:
    from building_permit_scraper import download_building_permit

    pdf_bytes, filename = download_building_permit(
        street="הארזים", number="29", city="בית שמש"
    )
    # pdf_bytes = bytes של ה-PDF, filename = שם הקובץ
    # מחזיר (None, None) אם לא נמצא
"""

import io
import os
import time
import asyncio
import httpx
from typing import Optional, Tuple
from israeli_municipalities_db import get_scraper_info


TIMEOUT = 20
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json, text/html, */*",
    "Accept-Language": "he-IL,he;q=0.9,en;q=0.8",
}


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def download_building_permit(
    street: str,
    number: str,
    city: str,
    prefer_oldest: bool = True,
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    מוריד PDF של תיק בניין (היתר/גרמושקא) לפי כתובת.

    מחזיר (pdf_bytes, filename) או (None, None) אם לא נמצא.
    prefer_oldest=True → מחזיר את התיק הישן ביותר (בנייה מקורית).
    """
    info = get_scraper_info(city)
    if not info["found"] or not info["can_scrape"]:
        print(f"[Scraper] {city} — אין מערכת דיגיטלית זמינה ({info.get('system')})")
        return None, None

    system = info["system"]
    print(f"[Scraper] חיפוש: {street} {number}, {city} | מערכת: {system}")

    try:
        if system == "COMPLOT":
            return _scrape_complot(info["api_base"], street, number, city, prefer_oldest)
        elif system == "BARTECH":
            return _scrape_bartech(info["api_base"], street, number, city)
        elif system == "TAVAPP":
            return _scrape_tel_aviv(street, number, prefer_oldest)
    except Exception as e:
        print(f"[Scraper] שגיאה: {e}")
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# COMPLOT scraper (~30 עיריות)
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_complot(
    api_base: str,
    street: str,
    number: str,
    city: str,
    prefer_oldest: bool,
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    קומפלוט API:
    1. GET /api/Buildings/GetBuildingsByAddress?street=X&houseNum=Y → רשימת בניינים
    2. GET /api/Buildings/GetBuildingFiles?buildingId=Z            → רשימת קבצים
    3. GET /api/Files/DownloadFile?fileId=W                        → הורדת PDF
    """
    with httpx.Client(headers=HEADERS, timeout=TIMEOUT, follow_redirects=True) as client:

        # שלב 1 — חיפוש בניין
        resp = client.get(
            f"{api_base}/api/Buildings/GetBuildingsByAddress",
            params={"street": street, "houseNum": number},
        )
        if resp.status_code != 200:
            # נסיון עם endpoint חלופי
            resp = client.get(
                f"{api_base}/newengine/api/Buildings/GetBuildingsByAddress",
                params={"street": street, "houseNum": number},
            )
        if resp.status_code != 200:
            print(f"[Complot] חיפוש נכשל: {resp.status_code}")
            return None, None

        buildings = resp.json()
        if not buildings:
            print(f"[Complot] לא נמצאו בניינים עבור {street} {number} ב{city}")
            return None, None

        print(f"[Complot] נמצאו {len(buildings)} בניינים")
        building = buildings[0]
        building_id = building.get("id") or building.get("BuildingId") or building.get("buildingId")
        if not building_id:
            print(f"[Complot] לא נמצא buildingId: {building}")
            return None, None

        # שלב 2 — רשימת קבצים
        resp2 = client.get(
            f"{api_base}/api/Buildings/GetBuildingFiles",
            params={"buildingId": building_id},
        )
        if resp2.status_code != 200:
            resp2 = client.get(
                f"{api_base}/newengine/api/Buildings/GetBuildingFiles",
                params={"buildingId": building_id},
            )
        if resp2.status_code != 200:
            print(f"[Complot] רשימת קבצים נכשלה: {resp2.status_code}")
            return None, None

        files = resp2.json()
        if not files:
            print(f"[Complot] לא נמצאו קבצים לבניין {building_id}")
            return None, None

        print(f"[Complot] נמצאו {len(files)} קבצים")

        # בחירת קובץ — העדפה: היתר/גרמושקא/תוכנית. ישן ביותר = בנייה מקורית
        target = _pick_best_file(files, prefer_oldest)
        if not target:
            print("[Complot] לא נמצא קובץ מתאים")
            return None, None

        file_id = target.get("id") or target.get("FileId") or target.get("fileId")
        filename = target.get("fileName") or target.get("FileName") or f"building_{building_id}.pdf"
        print(f"[Complot] מוריד: {filename} (id={file_id})")

        # שלב 3 — הורדת PDF
        resp3 = client.get(
            f"{api_base}/api/Files/DownloadFile",
            params={"fileId": file_id},
        )
        if resp3.status_code != 200:
            resp3 = client.get(
                f"{api_base}/newengine/api/Files/DownloadFile",
                params={"fileId": file_id},
            )
        if resp3.status_code == 200 and len(resp3.content) > 1000:
            print(f"[Complot] ✓ הורד {len(resp3.content):,} bytes")
            return resp3.content, filename
        else:
            print(f"[Complot] הורדה נכשלה: {resp3.status_code}")
            return None, None


# ─────────────────────────────────────────────────────────────────────────────
# BARTECH scraper (ערד)
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_bartech(
    api_base: str,
    street: str,
    number: str,
    city: str,
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    Bartech API (ערד):
    POST /api/PermitApplication/Search  { streetName, houseNumber }
    → רשימת בקשות
    GET  /api/PermitApplication/{id}/Documents → רשימת מסמכים
    GET  /api/Document/{docId}/Download        → PDF
    """
    with httpx.Client(headers=HEADERS, timeout=TIMEOUT, follow_redirects=True) as client:

        # שלב 1 — חיפוש
        resp = client.post(
            f"{api_base}/api/PermitApplication/Search",
            json={"streetName": street, "houseNumber": number},
        )
        if resp.status_code != 200:
            # נסיון GET
            resp = client.get(
                f"{api_base}/api/PermitApplication/Search",
                params={"streetName": street, "houseNumber": number},
            )
        if resp.status_code != 200:
            print(f"[Bartech] חיפוש נכשל: {resp.status_code}")
            return None, None

        results = resp.json()
        if not results:
            print(f"[Bartech] לא נמצאו תוצאות עבור {street} {number}")
            return None, None

        app_id = (results[0].get("id") or results[0].get("applicationId")
                  or results[0].get("Id"))
        if not app_id:
            print(f"[Bartech] לא נמצא ID: {results[0]}")
            return None, None

        print(f"[Bartech] נמצאה בקשה: {app_id}")

        # שלב 2 — מסמכים
        resp2 = client.get(f"{api_base}/api/PermitApplication/{app_id}/Documents")
        if resp2.status_code != 200:
            print(f"[Bartech] רשימת מסמכים נכשלה: {resp2.status_code}")
            return None, None

        docs = resp2.json()
        if not docs:
            print(f"[Bartech] לא נמצאו מסמכים")
            return None, None

        target = _pick_best_file(docs, prefer_oldest=True)
        if not target:
            target = docs[0]

        doc_id = target.get("id") or target.get("documentId") or target.get("Id")
        filename = target.get("fileName") or target.get("name") or f"arad_{app_id}.pdf"

        # שלב 3 — הורדה
        resp3 = client.get(f"{api_base}/api/Document/{doc_id}/Download")
        if resp3.status_code == 200 and len(resp3.content) > 1000:
            print(f"[Bartech] ✓ הורד {len(resp3.content):,} bytes")
            return resp3.content, filename
        else:
            print(f"[Bartech] הורדה נכשלה: {resp3.status_code}")
            return None, None


# ─────────────────────────────────────────────────────────────────────────────
# Tel Aviv scraper
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_tel_aviv(
    street: str,
    number: str,
    prefer_oldest: bool,
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    תל אביב — ארכיון הנדסי:
    GET https://archive-binyan.tel-aviv.gov.il/api/folders/search?street=X&houseNum=Y
    → רשימת תיקים
    GET /api/folders/{folderId}/files → רשימת קבצים
    GET /api/files/{fileId}/download  → PDF
    """
    base = "https://archive-binyan.tel-aviv.gov.il"
    with httpx.Client(headers=HEADERS, timeout=TIMEOUT, follow_redirects=True) as client:

        resp = client.get(
            f"{base}/api/folders/search",
            params={"street": street, "houseNum": number},
        )
        if resp.status_code != 200:
            print(f"[TelAviv] חיפוש נכשל: {resp.status_code}")
            return None, None

        folders = resp.json()
        if not folders:
            print(f"[TelAviv] לא נמצאו תיקים עבור {street} {number}")
            return None, None

        folder = folders[0]
        folder_id = folder.get("id") or folder.get("folderId")
        print(f"[TelAviv] תיק: {folder_id} — {folder.get('address', '')}")

        resp2 = client.get(f"{base}/api/folders/{folder_id}/files")
        if resp2.status_code != 200:
            print(f"[TelAviv] רשימת קבצים נכשלה: {resp2.status_code}")
            return None, None

        files = resp2.json()
        if not files:
            print(f"[TelAviv] לא נמצאו קבצים")
            return None, None

        target = _pick_best_file(files, prefer_oldest)
        if not target:
            target = files[0]

        file_id = target.get("id") or target.get("fileId")
        filename = target.get("fileName") or target.get("name") or f"telaviv_{folder_id}.pdf"

        resp3 = client.get(f"{base}/api/files/{file_id}/download")
        if resp3.status_code == 200 and len(resp3.content) > 1000:
            print(f"[TelAviv] ✓ הורד {len(resp3.content):,} bytes")
            return resp3.content, filename
        else:
            print(f"[TelAviv] הורדה נכשלה: {resp3.status_code}")
            return None, None


# ─────────────────────────────────────────────────────────────────────────────
# File picker
# ─────────────────────────────────────────────────────────────────────────────

PRIORITY_KEYWORDS = [
    "היתר", "גרמושקא", "גרמושקה", "הגשה", "תכנית", "קומה", "חתך",
    "permit", "plan", "drawing",
]

def _pick_best_file(files: list, prefer_oldest: bool) -> Optional[dict]:
    """בוחר את הקובץ הטוב ביותר מרשימה."""
    if not files:
        return None

    # מסנן PDF בלבד
    pdf_files = [
        f for f in files
        if (f.get("fileName") or f.get("name") or "").lower().endswith(".pdf")
        or (f.get("fileType") or f.get("type") or "").lower() == "pdf"
    ]
    if not pdf_files:
        pdf_files = files  # fallback — קח הכל

    # מחפש קבצים עם מילות מפתח בשם
    priority = [
        f for f in pdf_files
        if any(
            kw in (f.get("fileName") or f.get("name") or "").lower()
            for kw in PRIORITY_KEYWORDS
        )
    ]
    candidates = priority if priority else pdf_files

    # מיון לפי תאריך
    def get_date(f):
        for key in ("date", "uploadDate", "createdDate", "fileDate", "Date"):
            if f.get(key):
                return str(f[key])
        return ""

    candidates_sorted = sorted(candidates, key=get_date, reverse=not prefer_oldest)
    return candidates_sorted[0] if candidates_sorted else candidates[0]
