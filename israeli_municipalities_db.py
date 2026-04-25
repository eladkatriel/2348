"""
israeli_municipalities_db.py
=============================
מאגר מרכזי של כל הרשויות המקומיות בישראל עם פרטי ארכיון תיקי הבניין.

מערכות:
  COMPLOT  — complot.co.il (מערכת קומפלוט) — הנפוצה ביותר
  BARTECH  — bartech-net.co.il (ערד ואחרות)
  TAVAPP   — תל אביב (מערכת עצמאית — archive-binyan.tel-aviv.gov.il)
  NONE     — אין מערכת דיגיטלית (ארכיון פיזי בלבד)
  UNKNOWN  — לא נבדק
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Municipality:
    name: str                        # שם הרשות
    aliases: list                    # שמות נוספים / כינויים
    system: str                      # COMPLOT / BARTECH / TAVAPP / NONE / UNKNOWN
    archive_url: Optional[str]       # URL לארכיון תיקי בניין
    search_url: Optional[str]        # URL ישיר לחיפוש
    api_base: Optional[str]          # base URL ל-API (קומפלוט)
    notes: str = ""                  # הערות


# ─────────────────────────────────────────────────────────────────────────────
# המאגר המלא
# ─────────────────────────────────────────────────────────────────────────────

MUNICIPALITIES: list[Municipality] = [

    # ══════════════════════════════════════════════════════
    # מערכת קומפלוט — complot.co.il
    # כולן עובדות עם אותו API:
    #   GET {api_base}/api/Buildings/GetBuildingsByAddress?street=X&houseNum=Y
    # ══════════════════════════════════════════════════════

    Municipality(
        name="בית שמש",
        aliases=["beit shemesh", "bet shemesh"],
        system="COMPLOT",
        archive_url="https://bsbuilding.org.il/newengine/Pages/buildings2.aspx",
        search_url="https://bsbuilding.org.il/newengine/Pages/buildings2.aspx",
        api_base="https://bsbuilding.org.il",
    ),
    Municipality(
        name="רמת גן",
        aliases=["ramat gan"],
        system="COMPLOT",
        archive_url="https://handasa.ramat-gan.muni.il/newengine/Pages/buildings2.aspx",
        search_url="https://handasa.ramat-gan.muni.il/newengine/Pages/buildings2.aspx",
        api_base="https://handasa.ramat-gan.muni.il",
    ),
    Municipality(
        name="ראשון לציון",
        aliases=["rishon lezion", "rishon le-zion"],
        system="COMPLOT",
        archive_url="https://rishonlezion.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://rishonlezion.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://rishonlezion.complot.co.il",
    ),
    Municipality(
        name="רחובות",
        aliases=["rehovot"],
        system="COMPLOT",
        archive_url="https://rechovot.complot.co.il/buildings/",
        search_url="https://rechovot.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://rechovot.complot.co.il",
    ),
    Municipality(
        name="דימונה",
        aliases=["dimona"],
        system="COMPLOT",
        archive_url="https://dimona.complot.co.il/",
        search_url="https://dimona.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://dimona.complot.co.il",
    ),
    Municipality(
        name="בת ים",
        aliases=["bat yam"],
        system="COMPLOT",
        archive_url="https://batyam.complot.co.il/tikbinyan/",
        search_url="https://batyam.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://batyam.complot.co.il",
    ),
    Municipality(
        name="מודיעין",
        aliases=["modiin", "מודיעין מכבים רעות"],
        system="COMPLOT",
        archive_url="https://modiin.complot.co.il/buildings/",
        search_url="https://modiin.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://modiin.complot.co.il",
        notes="ארכיב פיזי נפרד — יש לבקש ממנהל הארכיון",
    ),
    Municipality(
        name="אור יהודה",
        aliases=["or yehuda"],
        system="COMPLOT",
        archive_url="https://oryehuda.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://oryehuda.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://oryehuda.complot.co.il",
    ),
    Municipality(
        name="גבעתיים",
        aliases=["givatayim"],
        system="COMPLOT",
        archive_url="https://givatayim.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://givatayim.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://givatayim.complot.co.il",
    ),
    Municipality(
        name="נס ציונה",
        aliases=["nes ziona"],
        system="COMPLOT",
        archive_url="https://nesziona.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://nesziona.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://nesziona.complot.co.il",
    ),
    Municipality(
        name="לוד",
        aliases=["lod"],
        system="COMPLOT",
        archive_url="https://lod.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://lod.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://lod.complot.co.il",
    ),
    Municipality(
        name="רמלה",
        aliases=["ramla"],
        system="COMPLOT",
        archive_url="https://ramla.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://ramla.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://ramla.complot.co.il",
    ),
    Municipality(
        name="נתניה",
        aliases=["netanya"],
        system="COMPLOT",
        archive_url="https://netanya.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://netanya.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://netanya.complot.co.il",
    ),
    Municipality(
        name="אשדוד",
        aliases=["ashdod"],
        system="COMPLOT",
        archive_url="https://ashdod.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://ashdod.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://ashdod.complot.co.il",
    ),
    Municipality(
        name="אשקלון",
        aliases=["ashkelon"],
        system="COMPLOT",
        archive_url="https://ashkelon.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://ashkelon.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://ashkelon.complot.co.il",
    ),
    Municipality(
        name="באר שבע",
        aliases=["beer sheva", "be'er sheva"],
        system="COMPLOT",
        archive_url="https://beersheva.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://beersheva.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://beersheva.complot.co.il",
    ),
    Municipality(
        name="הרצלייה",
        aliases=["herzliya"],
        system="COMPLOT",
        archive_url="https://herzliya.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://herzliya.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://herzliya.complot.co.il",
    ),
    Municipality(
        name="כפר סבא",
        aliases=["kfar saba"],
        system="COMPLOT",
        archive_url="https://kfarsaba.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://kfarsaba.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://kfarsaba.complot.co.il",
    ),
    Municipality(
        name="הוד השרון",
        aliases=["hod hasharon"],
        system="COMPLOT",
        archive_url="https://hodhasharon.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://hodhasharon.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://hodhasharon.complot.co.il",
    ),
    Municipality(
        name="רעננה",
        aliases=["raanana"],
        system="COMPLOT",
        archive_url="https://raanana.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://raanana.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://raanana.complot.co.il",
    ),
    Municipality(
        name="חולון",
        aliases=["holon"],
        system="COMPLOT",
        archive_url="https://holon.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://holon.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://holon.complot.co.il",
    ),
    Municipality(
        name="פתח תקווה",
        aliases=["petah tikva", "petach tikva"],
        system="COMPLOT",
        archive_url="https://pt.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://pt.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://pt.complot.co.il",
        notes="גם: https://www.petah-tikva.muni.il/engineering/planning-and-building/building2",
    ),
    Municipality(
        name="גבעת שמואל",
        aliases=["givat shmuel"],
        system="COMPLOT",
        archive_url="https://givatshmu.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://givatshmu.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://givatshmu.complot.co.il",
    ),
    Municipality(
        name="בני ברק",
        aliases=["bnei brak"],
        system="COMPLOT",
        archive_url="https://bneibrak.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://bneibrak.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://bneibrak.complot.co.il",
    ),
    Municipality(
        name="יבנה",
        aliases=["yavne"],
        system="COMPLOT",
        archive_url="https://yavne.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://yavne.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://yavne.complot.co.il",
    ),
    Municipality(
        name="קריית גת",
        aliases=["kiryat gat"],
        system="COMPLOT",
        archive_url="https://kgat.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://kgat.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://kgat.complot.co.il",
    ),
    Municipality(
        name="אופקים",
        aliases=["ofakim"],
        system="COMPLOT",
        archive_url="https://ofakim.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://ofakim.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://ofakim.complot.co.il",
    ),
    Municipality(
        name="נתיבות",
        aliases=["netivot"],
        system="COMPLOT",
        archive_url="https://netivot.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://netivot.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://netivot.complot.co.il",
    ),
    Municipality(
        name="שדרות",
        aliases=["sderot"],
        system="COMPLOT",
        archive_url="https://sderot.complot.co.il/newengine/Pages/buildings2.aspx",
        search_url="https://sderot.complot.co.il/newengine/Pages/buildings2.aspx",
        api_base="https://sderot.complot.co.il",
    ),
    Municipality(
        name="ערד",
        aliases=["arad"],
        system="BARTECH",
        archive_url="https://arad.bartech-net.co.il/SearchPermitApplication",
        search_url="https://arad.bartech-net.co.il/SearchPermitApplication",
        api_base="https://arad.bartech-net.co.il",
    ),

    # ══════════════════════════════════════════════════════
    # תל אביב — מערכת עצמאית
    # ══════════════════════════════════════════════════════

    Municipality(
        name="תל אביב",
        aliases=["תל אביב-יפו", "tel aviv", "tel-aviv"],
        system="TAVAPP",
        archive_url="https://archive-binyan.tel-aviv.gov.il/",
        search_url="https://handasa.tel-aviv.gov.il/Pages/SearchResultsAnonPageNew.aspx",
        api_base="https://archive-binyan.tel-aviv.gov.il",
        notes="חיפוש לפי רחוב + מספר. מחזיר רשימת תיקים עם PDFs",
    ),

    # ══════════════════════════════════════════════════════
    # ירושלים — מערכת עצמאית
    # ══════════════════════════════════════════════════════

    Municipality(
        name="ירושלים",
        aliases=["jerusalem"],
        system="UNKNOWN",
        archive_url="https://gis.jerusalem.muni.il/",
        search_url="https://gis.jerusalem.muni.il/",
        api_base=None,
        notes="GIS בלבד — אין ארכיון תיקי בניין מקוון ציבורי",
    ),

    # ══════════════════════════════════════════════════════
    # חיפה — מערכת עצמאית
    # ══════════════════════════════════════════════════════

    Municipality(
        name="חיפה",
        aliases=["haifa"],
        system="UNKNOWN",
        archive_url="https://www.haifa.muni.il/",
        search_url=None,
        api_base=None,
        notes="אין ארכיון דיגיטלי ציבורי לתיקי בניין",
    ),

    # ══════════════════════════════════════════════════════
    # ללא מערכת דיגיטלית
    # ══════════════════════════════════════════════════════

    Municipality(
        name="אשתאול",
        aliases=["eshta'ol"],
        system="NONE",
        archive_url=None,
        search_url=None,
        api_base=None,
        notes="מועצה אזורית מטה יהודה — ארכיון פיזי בלבד",
    ),
    Municipality(
        name="מצפה רמון",
        aliases=["mitzpe ramon"],
        system="NONE",
        archive_url=None,
        search_url=None,
        api_base=None,
        notes="ארכיון פיזי — יש לפנות ישירות לעירייה",
    ),
    Municipality(
        name="ערד",  # כבר הוגדר למעלה תחת BARTECH
        aliases=[],
        system="BARTECH",
        archive_url="https://arad.bartech-net.co.il/SearchPermitApplication",
        search_url="https://arad.bartech-net.co.il/SearchPermitApplication",
        api_base="https://arad.bartech-net.co.il",
    ),
]

# ── הסרת כפילויות ────────────────────────────────────────────────────────────
_seen = set()
MUNICIPALITIES_UNIQUE = []
for m in MUNICIPALITIES:
    if m.name not in _seen:
        _seen.add(m.name)
        MUNICIPALITIES_UNIQUE.append(m)
MUNICIPALITIES = MUNICIPALITIES_UNIQUE


# ─────────────────────────────────────────────────────────────────────────────
# Lookup API
# ─────────────────────────────────────────────────────────────────────────────

def get_municipality(city_name: str) -> Optional[Municipality]:
    """
    מחזיר Municipality לפי שם עיר (עברית או אנגלית, כולל כינויים).
    מחזיר None אם לא נמצא.
    """
    city_lower = city_name.strip().lower()
    for m in MUNICIPALITIES:
        if m.name.lower() == city_lower:
            return m
        if any(alias.lower() == city_lower for alias in m.aliases):
            return m
    # fuzzy — חיפוש חלקי
    for m in MUNICIPALITIES:
        if city_lower in m.name.lower() or m.name.lower() in city_lower:
            return m
        if any(city_lower in alias.lower() for alias in m.aliases):
            return m
    return None


def get_scraper_info(city_name: str) -> dict:
    """
    מחזיר dict עם כל המידע הדרוש ל-scraper.
    """
    m = get_municipality(city_name)
    if not m:
        return {"found": False, "city": city_name, "system": "UNKNOWN"}
    return {
        "found": True,
        "city": m.name,
        "system": m.system,
        "archive_url": m.archive_url,
        "search_url": m.search_url,
        "api_base": m.api_base,
        "notes": m.notes,
        "can_scrape": m.system in ("COMPLOT", "BARTECH", "TAVAPP"),
    }


def list_scrapable() -> list:
    """מחזיר רשימת כל הרשויות שניתן לסרוק אוטומטית."""
    return [m for m in MUNICIPALITIES if m.system in ("COMPLOT", "BARTECH", "TAVAPP")]


def list_not_available() -> list:
    """מחזיר רשימת רשויות ללא מערכת דיגיטלית."""
    return [m for m in MUNICIPALITIES if m.system == "NONE"]
