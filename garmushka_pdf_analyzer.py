"""
garmushka_pdf_analyzer.py
=========================
מחלץ אוטומטית מ-PDF של תיק בניין סרוק:
  - תאריך ההגשה / היתר
  - תכנית קומה טיפוסית / קומה א / קומת קרקע
  - חתך (רוחבי / אורכי)

שימוש:
    from garmushka_pdf_analyzer import analyze_building_pdf

    result = analyze_building_pdf("/path/to/00015775.pdf")
    # result = {
    #   "date": "15.03.1987",
    #   "floors": "2",           # אם נמצא
    #   "typical_floor_image": "/tmp/garm_typical_floor.png",
    #   "section_image":        "/tmp/garm_section.png",
    #   "ground_floor_image":   "/tmp/garm_ground_floor.png",
    # }

דרישות (כבר קיימות ב-Render / venv):
    Pillow, pypdf   (ו-pdftoppm מ-poppler-utils ברמת OS)
"""

import os
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Optional
from PIL import Image

# ── כותרות לחיפוש (עברית, RTL) ──────────────────────────────────────────────
FLOOR_KEYWORDS = [
    "קומה טיפוסית",
    "קומה א",
    "קומה ב",
    "קומת קרקע",
    "קרקע",
    "typical",
]
SECTION_KEYWORDS = [
    "חתך",
    "section",
]
DATE_PATTERNS = [
    r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\b",   # DD.MM.YYYY or DD/MM/YY
    r"\b(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})\b",      # YYYY-MM-DD
]

RENDER_DPI = 200        # רזולוציה לסריקה ראשונית
CROP_DPI   = 300        # רזולוציה לחיתוך סופי (איכות גבוהה יותר לתמונות)
STRIP_H    = 0.12       # חלק תחתון של כל פס אנכי שבו מחפשים כותרת (12%)
MIN_SECTION_W = 0.06    # רוחב מינימלי של פס (אחוז מרוחב הדף) — מסנן שוליים


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def analyze_building_pdf(pdf_path: str, output_dir: Optional[str] = None) -> dict:
    """
    מנתח PDF של תיק בניין ומחזיר dict עם:
      date, typical_floor_image, section_image, ground_floor_image, floors
    כל image הוא נתיב לקובץ PNG (או None אם לא נמצא).
    """
    pdf_path = str(pdf_path)
    if not os.path.exists(pdf_path):
        return {"error": f"קובץ לא נמצא: {pdf_path}"}

    out_dir = output_dir or tempfile.mkdtemp(prefix="garm_")
    os.makedirs(out_dir, exist_ok=True)

    result = {
        "date": None,
        "floors": None,
        "typical_floor_image": None,
        "section_image": None,
        "ground_floor_image": None,
        "pages_analyzed": 0,
    }

    page_count = _get_page_count(pdf_path)
    print(f"[Analyzer] {page_count} עמודים ב-{os.path.basename(pdf_path)}")

    for page_num in range(1, page_count + 1):
        page_img = _rasterize_page(pdf_path, page_num, RENDER_DPI, out_dir)
        if not page_img:
            continue
        result["pages_analyzed"] += 1

        img = Image.open(page_img)
        w, h = img.size

        # ── חיפוש תאריך בעמוד הראשון ──────────────────────────────────────
        if page_num == 1 and not result["date"]:
            result["date"] = _extract_date_from_image(img, pdf_path, page_num)

        # ── זיהוי עמוד פנורמי (רוחב >> גובה) — תיק בניין קלאסי ──────────
        is_panoramic = (w / h) > 2.5

        if is_panoramic:
            strips = _split_into_vertical_strips(img, MIN_SECTION_W)
            for strip_img, strip_x, strip_w in strips:
                title = _read_caption_area(strip_img)
                print(f"  [Analyzer] עמוד {page_num} | x={strip_x} | כותרת: {repr(title)}")

                if not result["typical_floor_image"] and _matches(title, FLOOR_KEYWORDS, exclude=["קרקע","ground"]):
                    path = _save_strip(strip_img, out_dir, f"typical_floor_p{page_num}_x{strip_x}")
                    result["typical_floor_image"] = path
                    print(f"  [Analyzer] ✓ קומה טיפוסית → {path}")

                if not result["ground_floor_image"] and _matches(title, ["קרקע", "קומת קרקע", "ground"]):
                    path = _save_strip(strip_img, out_dir, f"ground_floor_p{page_num}_x{strip_x}")
                    result["ground_floor_image"] = path
                    print(f"  [Analyzer] ✓ קומת קרקע → {path}")

                if not result["section_image"] and _matches(title, SECTION_KEYWORDS):
                    # שמור את הפס כולו (שני חתכים ביחד) אם יש שניים
                    path = _save_strip(strip_img, out_dir, f"section_p{page_num}_x{strip_x}")
                    result["section_image"] = path
                    print(f"  [Analyzer] ✓ חתך → {path}")

        else:
            # עמוד רגיל — ניסיון לקרוא כותרת בתחתית הדף כולו
            title = _read_caption_area(img)
            print(f"  [Analyzer] עמוד {page_num} (רגיל) | כותרת: {repr(title)}")
            if not result["typical_floor_image"] and _matches(title, FLOOR_KEYWORDS, exclude=["קרקע"]):
                path = _save_strip(img, out_dir, f"typical_floor_p{page_num}")
                result["typical_floor_image"] = path
            if not result["section_image"] and _matches(title, SECTION_KEYWORDS):
                path = _save_strip(img, out_dir, f"section_p{page_num}")
                result["section_image"] = path

        # עצור אם מצאנו הכל
        if result["typical_floor_image"] and result["section_image"]:
            break

    return result


# ─────────────────────────────────────────────────────────────────────────────
# PDF utilities
# ─────────────────────────────────────────────────────────────────────────────

def _get_page_count(pdf_path: str) -> int:
    try:
        from pypdf import PdfReader
        return len(PdfReader(pdf_path).pages)
    except Exception:
        return 3


def _rasterize_page(pdf_path: str, page_num: int, dpi: int, out_dir: str) -> Optional[str]:
    """ממיר עמוד PDF לתמונה JPEG ומחזיר את הנתיב."""
    prefix = os.path.join(out_dir, f"page_{page_num:03d}")
    try:
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", str(dpi),
             "-f", str(page_num), "-l", str(page_num),
             pdf_path, prefix],
            check=True, capture_output=True,
        )
        # pdftoppm adds zero-padded suffix
        candidates = sorted(Path(out_dir).glob(f"page_{page_num:03d}-*.jpg"))
        return str(candidates[0]) if candidates else None
    except Exception as e:
        print(f"[Analyzer] שגיאה בהמרת עמוד {page_num}: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Image analysis
# ─────────────────────────────────────────────────────────────────────────────

def _split_into_vertical_strips(img: Image.Image, min_width_ratio: float) -> list:
    """
    מחלק תמונה פנורמית לפסים אנכיים לפי עמודות אנכיות כהות (קווי הפרדה בין תרשימים).
    מחזיר רשימה של (crop_image, x_start, width).
    """
    import numpy as np

    w, h = img.size
    gray = img.convert("L")
    arr = np.array(gray)

    # ממוצע לכל עמודה — עמודות הפרדה הן כהות (ערך נמוך)
    col_mean = arr.mean(axis=0)

    # נרמול
    threshold = col_mean.min() + (col_mean.max() - col_mean.min()) * 0.35

    # זיהוי אזורי הפרדה (עמודות רציפות מתחת לסף)
    in_sep = col_mean < threshold
    min_px = int(w * 0.005)    # רוחב מינימלי של קו הפרדה

    boundaries = [0]
    i = 0
    while i < w:
        if in_sep[i]:
            j = i
            while j < w and in_sep[j]:
                j += 1
            if j - i >= min_px:
                mid = (i + j) // 2
                boundaries.append(mid)
            i = j
        else:
            i += 1
    boundaries.append(w)

    strips = []
    min_w_px = int(w * min_width_ratio)
    for idx in range(len(boundaries) - 1):
        x0, x1 = boundaries[idx], boundaries[idx + 1]
        if x1 - x0 >= min_w_px:
            crop = img.crop((x0, 0, x1, h))
            strips.append((crop, x0, x1 - x0))

    # fallback: חלוקה אחידה ל-6 אם לא נמצאו גבולות
    if len(strips) < 2:
        n = 6
        sw = w // n
        strips = [(img.crop((i*sw, 0, (i+1)*sw, h)), i*sw, sw) for i in range(n)]

    return strips


def _read_caption_area(img: Image.Image) -> str:
    """
    קורא את אזור הכותרת — החלק התחתון של התמונה.
    מכיוון שאין tesseract עברי, משתמש ב-Claude Vision דרך API.
    """
    w, h = img.size
    # חלק תחתון (12%) + 5% גובה מינימלי
    caption_h = max(int(h * STRIP_H), 40)
    caption_crop = img.crop((0, h - caption_h, w, h))

    # שמור זמנית וקרא כ-base64
    import io, base64
    buf = io.BytesIO()
    caption_crop.save(buf, format="JPEG", quality=85)
    b64 = base64.b64encode(buf.getvalue()).decode()

    try:
        import requests, os
        api_key = os.environ.get("MONDAY_API_KEY", "")  # לא נכון — נשתמש ב-ANTHROPIC_API_KEY
        anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not anthropic_key:
            return _fallback_ocr_heuristic(caption_crop)

        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": anthropic_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-haiku-4-5-20251001",
                "max_tokens": 100,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                        {"type": "text", "text": "What Hebrew text appears in this image? Return ONLY the text, nothing else. Focus on title/caption text at the bottom."}
                    ]
                }]
            },
            timeout=10,
        )
        if resp.status_code == 200:
            return resp.json()["content"][0]["text"].strip()
    except Exception:
        pass

    return _fallback_ocr_heuristic(caption_crop)


def _fallback_ocr_heuristic(img: Image.Image) -> str:
    """
    Fallback כאשר אין API זמין — מנסה pytesseract עם English OCR.
    לא יידע לקרוא עברית אבל יידע לזהות מספרים ומילים לועזיות.
    """
    try:
        import pytesseract
        text = pytesseract.image_to_string(img, lang="eng")
        return text.strip()
    except Exception:
        return ""


def _extract_date_from_image(img: Image.Image, pdf_path: str, page_num: int) -> Optional[str]:
    """מחלץ תאריך מהעמוד הראשון של הקובץ."""
    # נסיון ראשון: OCR על אזור הכותרת (פינה עליונה ימנית — כותרת הדף)
    w, h = img.size
    # פינה עליונה ימנית (RTL — שם בדרך כלל כותרת + תאריך)
    header_crop = img.crop((w // 2, 0, w, h // 6))
    text = _read_caption_area(header_crop)

    for pat in DATE_PATTERNS:
        m = re.search(pat, text)
        if m:
            return _normalize_date(m)

    # נסיון שני: ממטא-דאטה של ה-PDF
    try:
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        meta = reader.metadata
        for attr in ["/CreationDate", "/ModDate"]:
            raw = getattr(meta, attr.lstrip("/").lower(), None) or meta.get(attr, "")
            if raw:
                m = re.search(r"(\d{4})(\d{2})(\d{2})", str(raw))
                if m:
                    return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    except Exception:
        pass

    return None


def _normalize_date(match) -> str:
    groups = match.groups()
    if len(groups[0]) == 4:          # YYYY-MM-DD
        return f"{groups[2]}.{groups[1]}.{groups[0]}"
    else:                             # DD.MM.YYYY or DD.MM.YY
        year = groups[2]
        if len(year) == 2:
            year = "19" + year if int(year) > 30 else "20" + year
        return f"{groups[0].zfill(2)}.{groups[1].zfill(2)}.{year}"


def _matches(text: str, keywords: list, exclude: list = None) -> bool:
    """בודק אם הטקסט מכיל אחת מהמילות המפתח (ולא מילות החרגה)."""
    text_lower = text.lower()
    if exclude:
        for ex in exclude:
            if ex.lower() in text_lower:
                return False
    return any(kw.lower() in text_lower for kw in keywords)


def _save_strip(img: Image.Image, out_dir: str, name: str) -> str:
    """שומר תמונת פס כ-PNG ומחזיר את הנתיב."""
    # חיתוך שוליים לבנים
    img_clean = _trim_whitespace(img)
    path = os.path.join(out_dir, f"{name}.png")
    img_clean.save(path, format="PNG")
    return path


def _trim_whitespace(img: Image.Image, threshold: int = 240) -> Image.Image:
    """חיתוך שוליים לבנים/בהירים מסביב לתמונה."""
    import numpy as np
    arr = np.array(img.convert("L"))
    mask = arr < threshold
    rows = np.any(mask, axis=1)
    cols = np.any(mask, axis=0)
    if not rows.any():
        return img
    rmin, rmax = np.where(rows)[0][[0, -1]]
    cmin, cmax = np.where(cols)[0][[0, -1]]
    padding = 20
    rmin = max(0, rmin - padding)
    rmax = min(arr.shape[0], rmax + padding)
    cmin = max(0, cmin - padding)
    cmax = min(arr.shape[1], cmax + padding)
    return img.crop((cmin, rmin, cmax, rmax))
