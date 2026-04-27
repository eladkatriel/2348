"""
building_permit_scraper.py
===========================
מוריד אוטומטית PDF של תיק בניין מאתר העירייה — דרך Playwright (דפדפן אמיתי).

הגדרת Render — Build Command:
  pip install -r requirements.txt && playwright install chromium --with-deps

תומך ב: קומפלוט (~30 עיריות), Bartech (ערד), תל אביב
"""

import os
import time
from typing import Optional, Tuple
from israeli_municipalities_db import get_scraper_info

PRIORITY_KEYWORDS = [
    "היתר", "גרמושקא", "גרמושקה", "הגשה", "תכנית", "קומה", "חתך",
    "permit", "plan", "drawing",
]


def download_building_permit(
    street: str,
    number: str,
    city: str,
    prefer_oldest: bool = True,
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    מוריד PDF של תיק בניין לפי כתובת.
    מחזיר (pdf_bytes, filename) או (None, None).
    """
    info = get_scraper_info(city)
    if not info["found"] or not info["can_scrape"]:
        print(f"[Scraper] {city} — אין מערכת דיגיטלית ({info.get('system')})")
        return None, None

    system   = info["system"]
    api_base = info["api_base"]
    print(f"[Scraper] מחפש: {street} {number}, {city} | מערכת: {system}")

    try:
        if system == "COMPLOT":
            return _scrape_complot(api_base, street, number, prefer_oldest)
        elif system == "BARTECH":
            return _scrape_bartech(api_base, street, number)
        elif system == "TAVAPP":
            return _scrape_tel_aviv(street, number, prefer_oldest)
    except Exception as e:
        print(f"[Scraper] שגיאה: {type(e).__name__}: {e}")
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# Playwright browser
# ─────────────────────────────────────────────────────────────────────────────

def _launch():
    """מאתחל Playwright ומחזיר (pw, browser)."""
    from playwright.sync_api import sync_playwright
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-setuid-sandbox",
              "--disable-dev-shm-usage", "--disable-gpu"],
    )
    return pw, browser


def _page(browser):
    """פותח context + page עם locale עברי."""
    ctx = browser.new_context(
        locale="he-IL",
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        accept_downloads=True,
    )
    p = ctx.new_page()
    p.set_default_timeout(20000)
    return p, ctx


def _try_fill(page, selectors: list, value: str) -> bool:
    for sel in selectors:
        try:
            page.fill(sel, value, timeout=3000)
            return True
        except Exception:
            continue
    return False


def _try_click(page, selectors: list) -> bool:
    for sel in selectors:
        try:
            page.click(sel, timeout=3000)
            return True
        except Exception:
            continue
    return False


def _download_pdf(page) -> Tuple[Optional[bytes], Optional[str]]:
    """מנסה למצוא כפתור הורדה ולהוריד PDF."""
    selectors = [
        "a[href*='.pdf']",
        "button:has-text('הורד')",
        "a:has-text('הורד')",
        "button:has-text('PDF')",
        "a:has-text('PDF')",
        "[title*='הורד']",
        ".download-btn",
        "[class*='download']",
    ]
    for sel in selectors:
        try:
            with page.expect_download(timeout=15000) as dl_info:
                page.click(sel, timeout=4000)
            dl = dl_info.value
            tmp = dl.path()
            if tmp:
                with open(tmp, "rb") as f:
                    data = f.read()
                if len(data) > 500:
                    name = dl.suggested_filename or "building.pdf"
                    print(f"[Scraper] ✓ הורד PDF: {name} ({len(data):,} bytes)")
                    return data, name
        except Exception:
            continue
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# COMPLOT — ~30 עיריות
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_complot(api_base, street, number, prefer_oldest):
    search_url = f"{api_base}/newengine/Pages/buildings2.aspx"
    print(f"[Complot] {search_url}")
    pw, browser = _launch()
    try:
        page, ctx = _page(browser)

        # ── נסיון 1: API ישיר דרך playwright request (עוקף DNS של Render) ──
        for ep in [
            f"{api_base}/api/Buildings/GetBuildingsByAddress",
            f"{api_base}/newengine/api/Buildings/GetBuildingsByAddress",
        ]:
            try:
                r = page.request.get(ep, params={"street": street, "houseNum": number})
                if r.ok:
                    buildings = r.json()
                    if buildings:
                        print(f"[Complot] API: {len(buildings)} בניינים")
                        result = _complot_api_download(page, api_base, buildings[0], prefer_oldest)
                        if result[0]:
                            return result
            except Exception as e:
                print(f"[Complot] API {ep}: {e}")

        # ── נסיון 2: דפדפן מלא ────────────────────────────────────────────
        page.on("download", lambda d: None)  # מאפשר downloads
        page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(2)

        _try_fill(page, [
            "input[placeholder*='רחוב']", "input[placeholder*='street']",
            "#streetInput", "#street", "input[name='street']",
        ], street)

        _try_fill(page, [
            "input[placeholder*='מספר']", "input[placeholder*='number']",
            "#houseNumInput", "#houseNum", "input[name='houseNum']",
        ], number)

        _try_click(page, [
            "button:has-text('חיפוש')", "input[type='submit']",
            "#searchBtn", "button[type='submit']",
        ])
        time.sleep(3)

        # לחיצה על תוצאה ראשונה
        _try_click(page, [
            ".building-row:first-child", "tr:nth-child(2) td",
            ".result-row:first-child", "tbody tr:first-child",
        ])
        time.sleep(2)

        return _download_pdf(page)

    finally:
        browser.close()
        pw.stop()


def _complot_api_download(page, api_base, building, prefer_oldest):
    """הורדת קובץ ספציפי דרך API אחרי שמצאנו building."""
    bid = (building.get("id") or building.get("BuildingId")
           or building.get("buildingId") or building.get("ID"))
    if not bid:
        return None, None

    for ep in [f"{api_base}/api/Buildings/GetBuildingFiles",
               f"{api_base}/newengine/api/Buildings/GetBuildingFiles"]:
        try:
            r = page.request.get(ep, params={"buildingId": bid})
            if not r.ok:
                continue
            files = r.json()
            if not files:
                continue
            target = _pick_best_file(files, prefer_oldest)
            if not target:
                continue
            fid = (target.get("id") or target.get("FileId") or target.get("fileId"))
            fname = target.get("fileName") or target.get("FileName") or "building.pdf"
            for dl_ep in [f"{api_base}/api/Files/DownloadFile",
                          f"{api_base}/newengine/api/Files/DownloadFile"]:
                try:
                    dr = page.request.get(dl_ep, params={"fileId": fid})
                    body = dr.body()
                    if dr.ok and len(body) > 500:
                        print(f"[Complot] ✓ {fname} ({len(body):,} bytes)")
                        return body, fname
                except Exception:
                    pass
        except Exception as e:
            print(f"[Complot] files API {ep}: {e}")
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# BARTECH — ערד
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_bartech(api_base, street, number):
    search_url = f"{api_base}/SearchPermitApplication"
    print(f"[Bartech] {search_url}")
    pw, browser = _launch()
    try:
        page, ctx = _page(browser)
        page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(2)

        _try_fill(page, [
            "input[name*='treet']", "input[placeholder*='רחוב']",
            "input[id*='treet']",
        ], street)

        _try_fill(page, [
            "input[name*='umber']", "input[placeholder*='מספר']",
            "input[id*='umber']",
        ], number)

        _try_click(page, [
            "button[type='submit']", "input[type='submit']",
            "button:has-text('חיפוש')", "button:has-text('Search')",
        ])
        time.sleep(3)

        _try_click(page, ["tr:nth-child(2)", ".result-row:first-child", "tbody tr:first-child td"])
        time.sleep(2)

        return _download_pdf(page)
    finally:
        browser.close()
        pw.stop()


# ─────────────────────────────────────────────────────────────────────────────
# תל אביב
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_tel_aviv(street, number, prefer_oldest):
    search_url = "https://archive-binyan.tel-aviv.gov.il/"
    print(f"[TelAviv] {search_url}")
    pw, browser = _launch()
    try:
        page, ctx = _page(browser)
        page.goto(search_url, wait_until="networkidle", timeout=30000)
        time.sleep(2)

        _try_fill(page, [
            "input[placeholder*='רחוב']", "#street", "input[name='street']",
        ], street)

        _try_fill(page, [
            "input[placeholder*='מספר']", "#houseNum", "input[name='houseNum']",
        ], number)

        _try_click(page, [
            "button:has-text('חיפוש')", "input[type='submit']",
            "button[type='submit']",
        ])
        time.sleep(3)

        _try_click(page, [
            ".result-item:first-child", "tr:nth-child(2) td",
            ".building-result:first-child",
        ])
        time.sleep(2)

        return _download_pdf(page)
    finally:
        browser.close()
        pw.stop()


# ─────────────────────────────────────────────────────────────────────────────
# File picker
# ─────────────────────────────────────────────────────────────────────────────

def _pick_best_file(files: list, prefer_oldest: bool) -> Optional[dict]:
    if not files:
        return None
    pdf_files = [
        f for f in files
        if (f.get("fileName") or f.get("name") or "").lower().endswith(".pdf")
        or (f.get("fileType") or f.get("type") or "").lower() == "pdf"
    ] or files
    priority = [
        f for f in pdf_files
        if any(kw in (f.get("fileName") or f.get("name") or "").lower()
               for kw in PRIORITY_KEYWORDS)
    ]
    candidates = priority or pdf_files

    def get_date(f):
        for k in ("date", "uploadDate", "createdDate", "fileDate", "Date"):
            if f.get(k):
                return str(f[k])
        return ""

    return sorted(candidates, key=get_date, reverse=not prefer_oldest)[0]
