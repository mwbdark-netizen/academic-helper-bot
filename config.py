# -*- coding: utf-8 -*-
"""الإعدادات والثوابت وقائمة المواد — يُستورد من كل الملفات الأخرى."""

import os
import logging
from datetime import datetime, timezone, timedelta

try:
    from telegram import LinkPreviewOptions
    NO_PREVIEW = LinkPreviewOptions(is_disabled=True)
except Exception:
    NO_PREVIEW = None

# ══════════════════════════════ الإعدادات ══════════════════════════════

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("telegram.ext.Application").setLevel(logging.WARNING)
logging.getLogger("fontTools").setLevel(logging.ERROR)
logger = logging.getLogger("academic_bot")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()

_admin_raw = os.getenv("ADMIN_ID", "0").strip()
ADMIN_ID = int(_admin_raw) if _admin_raw.lstrip("-").isdigit() else 0

PORT = int(os.getenv("PORT", "10000"))

DATA_DIR = (os.getenv("DATA_DIR", ".").rstrip("/") or ".")
STATS_FILE = os.path.join(DATA_DIR, "bot_stats.json")
PERSIST_FILE = os.path.join(DATA_DIR, "bot_persistence.pkl")

# قناة خاصة لحفظ نسخة الإحصائيات (تبدأ بـ -100)
BACKUP_CHAT_ID = os.getenv("BACKUP_CHAT_ID", "").strip()
BACKUP_HOURS = float(os.getenv("BACKUP_HOURS", "6") or 6)
BAGHDAD_TZ = timezone(timedelta(hours=3))
KEEP_DAYS = 60

# بيانات الغلاف الاختياري
UNIVERSITY = os.getenv("UNIVERSITY", "جامعة القادسية")
COLLEGE = os.getenv("COLLEGE", "كلية علوم الحاسوب وتكنولوجيا المعلومات")
DEPARTMENT = os.getenv("DEPARTMENT", "قسم علوم الحاسوب")
STAGE = os.getenv("STAGE", "المرحلة الثانية")

# السنة الدراسية: تُحسب تلقائياً حسب الشهر، أو تُثبَّت عبر المتغيّر ACADEMIC_YEAR
ACADEMIC_YEAR_ENV = os.getenv("ACADEMIC_YEAR", "").strip()
ACADEMIC_START_MONTH = int(os.getenv("ACADEMIC_START_MONTH", "8") or 8)


def academic_year() -> str:
    """بعد شهر آب تبدأ سنة دراسية جديدة: 2026 - 2027 مثلاً."""
    if ACADEMIC_YEAR_ENV:
        return ACADEMIC_YEAR_ENV
    now = datetime.now(BAGHDAD_TZ)
    y = now.year if now.month >= ACADEMIC_START_MONTH else now.year - 1
    return f"{y} - {y + 1}"


# الخطوط
ARABIC_FONT = os.getenv("ARABIC_FONT", "Arial")
LATIN_FONT = os.getenv("LATIN_FONT", "Times New Roman")
MONO_FONT = os.getenv("MONO_FONT", "Courier New")   # خط كتل الكود في Word
FONT_DIR = os.getenv("FONT_DIR", "fonts")
AR_TTF = os.path.join(FONT_DIR, os.getenv("AR_TTF", "Amiri-Regular.ttf"))
AR_TTF_BOLD = os.path.join(FONT_DIR, os.getenv("AR_TTF_BOLD", "Amiri-Bold.ttf"))
# خط ثابت العرض للكود في PDF؛ اختياري، وإن غاب يُستخدم الخط العادي
MONO_TTF = os.path.join(FONT_DIR, os.getenv("MONO_TTF", "DejaVuSansMono.ttf"))

MAX_INPUT_CHARS = 8000
TELEGRAM_CHUNK = 3800

# حدود الصور
MAX_IMAGES = 5
IMG_MAX_SIDE = 1600
IMG_QUALITY = 85
IMG_MAX_BYTES = 4 * 1024 * 1024
ALBUM_WAIT = 2.5

# يوتيوب
YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY", "").strip()
YT_RESULTS = 3
YT_FETCH = 8          # نطلب نتائج أكثر ثم نفلترها؛ الكلفة 100 وحدة سواء 3 أو 8
YT_MAX_QUERIES = int(os.getenv("YT_MAX_QUERIES", "2") or 2)   # كل عبارة = 100 وحدة
YT_TIMEOUT = 12.0
YT_API = "https://www.googleapis.com/youtube/v3/search"
YT_SEARCH_URL = "https://www.youtube.com/results?search_query="

# لغة الشرح: عربية افتراضياً مع الاحتفاظ بالمصطلحات الإنكليزية
EXPLAIN_IN_ARABIC = os.getenv("EXPLAIN_IN_ARABIC", "1").strip() != "0"

BUSY_USERS: set[int] = set()

# ══════════════════════════════ المواد ══════════════════════════════
# lang: لغة المادة -> لغة الناتج      is_code: تظهر في خدمة تدقيق الأكواد

SUBJECTS = {
    # ── الكورس الأول ──
    "t1_oop":  {"name": "البرمجة الكيانية (OOP)",
                "term": 1, "lang": "en", "is_code": True,
                "en": "Object Oriented Programming"},
    "t1_mp":   {"name": "المعالجات الدقيقة (Microprocessors)",
                "term": 1, "lang": "en", "is_code": True,
                "en": "Microprocessors"},
    "t1_num":  {"name": "الطرق العددية (Numerical Methods)",
                "term": 1, "lang": "en", "is_code": True,
                "en": "Numerical Methods"},
    "t1_algo": {"name": "تحليل وتصميم الخوارزميات (Algorithms)",
                "term": 1, "lang": "en", "is_code": True,
                "en": "Analysis and Design of Algorithms"},
    "t1_emg":  {"name": "الإدارة الإلكترونية",
                "term": 1, "lang": "ar", "is_code": False,
                "en": "Electronic Management"},
    "t1_ar":   {"name": "اللغة العربية",
                "term": 1, "lang": "ar", "is_code": False,
                "en": "Arabic Language"},
    # ── الكورس الثاني ──
    "t2_ds":   {"name": "هياكل البيانات (Data Structures)",
                "term": 2, "lang": "en", "is_code": True,
                "en": "Data Structures"},
    "t2_java": {"name": "جافا (Java)",
                "term": 2, "lang": "en", "is_code": True,
                "en": "Java Programming"},
    "t2_toc":  {"name": "النظرية الاحتسابية (Theory of Computation)",
                "term": 2, "lang": "en", "is_code": False,
                "en": "Theory of Computation"},
    "t2_stat": {"name": "الإحصاء والاحتمالية (Probability & Statistics)",
                "term": 2, "lang": "en", "is_code": False,
                "en": "Probability and Statistics"},
    "t2_arch": {"name": "معمارية الحاسوب (Computer Architecture)",
                "term": 2, "lang": "en", "is_code": True,
                "en": "Computer Architecture"},
    "t2_en":   {"name": "اللغة الإنكليزية",
                "term": 2, "lang": "en", "is_code": False,
                "en": "English Language"},
}

SERVICE_LABELS = {
    "rep": "تقرير أكاديمي",
    "exp": "شرح تعليمي",
    "math": "حل مسألة",
    "code": "تدقيق كود",
    "quiz": "بنك أسئلة",
    "exp_img": "شرح من صورة",
    "math_img": "مسألة من صورة",
    "code_img": "كود من صورة",
}
