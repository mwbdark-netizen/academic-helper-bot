# -*- coding: utf-8 -*-
"""
المنصة الأكاديمية الذكية - بوت تلغرام
المرحلة الثانية - علوم الحاسوب

الخدمات: تقرير (Word / PDF) - شرح موضوع أو صفحة ملزمة - حل مسائل
         تدقيق أكواد - بنك أسئلة - تحميل ملازم المحاضرات
لغة الناتج تتبع لغة دراسة المادة، ما عدا الشرح فيكون بالعربية افتراضياً
مع الحفاظ على المصطلحات الإنكليزية. الصور مدعومة في المسائل والأكواد والشرح.

إحصائيات دائمة: عدد المستخدمين ومرات استخدام كل خدمة، مع نسخ احتياطي
تلقائي إلى قناة خاصة واستعادة تلقائية بعد كل إعادة نشر.

هذا الملف هو نقطة التشغيل: الواجهة والأوامر والأزرار ومعالجة الرسائل.
"""

import io
import os
import html
import json
import time
import asyncio
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.constants import ChatAction
from telegram.error import Forbidden, RetryAfter, Conflict, InvalidToken
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    PicklePersistence,
    filters,
)

from config import (
    logger, TELEGRAM_BOT_TOKEN, GROQ_API_KEY, ADMIN_ID, PORT,
    DATA_DIR, STATS_FILE, PERSIST_FILE, BACKUP_CHAT_ID, BACKUP_HOURS,
    MAX_INPUT_CHARS, MAX_IMAGES, ALBUM_WAIT,
    YOUTUBE_API_KEY, YT_RESULTS, YT_FETCH, YT_MAX_QUERIES,
    EXPLAIN_IN_ARABIC, BUSY_USERS, SUBJECTS, SERVICE_LABELS,
    ACADEMIC_YEAR_ENV, academic_year, AR_TTF, AR_TTF_BOLD, MONO_TTF,
    GEMINI_API_KEY, PROVIDER_LABELS, available_providers,
)
from text_utils import (
    clean_code_output, clean_report_text, clean_text_strictly,
    strip_latin_parentheticals, headings_for_telegram, extract_topic,
    split_for_telegram, safe_filename,
)
from ai_client import query_ai_text, query_ai_vision, VisionUnavailable, fetch_images, PIL_OK
from youtube_api import build_youtube_block, send_plain, _YT_CACHE
from prompts import SYS_BUILDERS, explain_lang, OCR_EXPLAIN, OCR_PREFIX, OCR_CODE_EXTRA
from documents import create_academic_docx, create_academic_pdf, _pdf_ready, FPDF_OK, BIDI_OK
from stats import (
    record_activity, record_file, remove_users, get_stats, build_stats_text,
    do_backup, restore_stats, backup_job, mark_clean,
    _stats_lock, _read_stats_file, _write_stats_file, _merge_stats,
)


# ===================== إعداد المرونة عند التعارض =====================

POLL_RETRY_CONFLICT = 25      # ثانية انتظار قبل إعادة المحاولة بعد Conflict
POLL_RETRY_OTHER = 15         # ثانية انتظار بعد أي انقطاع آخر
POLL_MAX_FAILURES = 5         # بعد هذا العدد نُخرج العملية ليعيدها مشرف العمليات

DROP_PENDING_UPDATES = os.getenv("DROP_PENDING_UPDATES", "0").strip() == "1"


# ===================== الملازم والمحاضرات =====================

LECTURE_SUBJECTS = {
    "oop":       "البرمجة الكيانية نظري",
    "micro":     "المعالجات الدقيقة نظري",
    "micro_lab": "المعالجات الدقيقة عملي",
    "algo":      "الخوارزميات",
    "emgmt":     "الإدارة الإلكترونية",
    "arab":      "اللغة العربية",
    "oop_prac":  "البرمجة الكيانية عملي",
}

LECTURE_COUNT = 7             # الحد الأدنى لعدد أزرار المحاضرات المعروضة لكل مادة
DELETED_MARKER = "__DELETED__"

# الملازم الأساسية الثابتة في الكود، وتُدمج معها تلقائياً أي ملازم يضيفها
# المشرف أو يحذفها مباشرة عبر البوت.
LECTURES: dict[str, dict[int, str]] = {
    "oop": {
        1: "BQACAgIAAxkBAAIIZWrKkMUw6UBpTHUOMX0dfBfow3WNAALXoAACe0xYSnUgFN1EgYB-PQQ",
    },
    "micro": {
        1: "BQACAgIAAxkBAAIGXGq4GTM_ehC2koTfuyxShWEi6PBbAAIpowACxuDBSWkJNjeZxCzpPQQ",
        2: "BQACAgIAAxkBAAIHvmrH2j7r1zbrowvihi0HvNFio6ZPAAJrqgACDmtASsU-27ak9tXNPQQ",
        3: "BQACAgIAAxkBAAIIg2rKlwtZa2W4e09H-YOv61t1otFkAAOhAAJ7TFhKcirMAXnZDMY9BA",
        4: "BQACAgIAAxkBAAIILmrKbAZxeAVUinOYeEB-iP2awRrsAAKtnwACe0xYStvC_yZ0cYeRPQQ",
        5: "BQACAgIAAxkBAAIIMGrKbA0ngPs4UkPsXOd7gvA4jaaqAAKunwACe0xYSjZR4fAYrfQWPQQ",
        6: "BQACAgIAAxkBAAIIMmrKbBP6IbssDufAYc1ArUK5lZ5OAAKwnwACe0xYSmxV_cJOp4-CPQQ",
        7: "BQACAgIAAxkBAAIINGrKbBmFuYDKpQYoUJxFgkhf7A4WAAKxnwACe0xYStXSxM_-obAfPQQ",
    },
    "micro_lab": {
        1: "BQACAgIAAxkBAAIHx2rH3UNOgbeCw7cN5MOXPtYQjxU3AAKrqgACDmtAStJ7uBPpN-PGPQQ",
    },
    "algo": {
        1: "BQACAgIAAxkBAAIHQ2rEDbkeZAtD4FJuI_MWbnt13RyjAAJhsAACmnQgSk9PRF1j-iHFPQQ",
        2: "BQACAgIAAxkBAAII72rOEZ3aPZGsmIl-XQABolxRaJJWJQACz6gAAucJcUqHDmajvwIBaT0E",
    },
    "emgmt": {
        1: "BQACAgIAAxkBAAIHrWrH1Xlkd9MNVNmKaPVjXxlbHfLvAAKMsAACzhg5Sgg-btX63PpsPQQ",
        2: "BQACAgIAAxkBAAIHr2rH1amXWSshfsPVyWwM8PwzV9H1AAKNsAACzhg5SnINS0mStTH3PQQ",
        3: "BQACAgIAAxkBAAIIzGrNWVkoZE53GoFXqmyvnaFhHylYAAJfrQAChb1oSk2DjeI3lqn-PQQ",
    },
    "arab": {
        1: "BQACAgIAAxkBAAIJQ2rQlF3KJsrRs_X2XWsUnJFuDL7UAAIZrAACC3uBSm_bkhAgPEaxPQQ",
    },
    "oop_prac": {
        1: "BQACAgIAAxkBAAIJRWrQlI2_rV4YQSOVmNEWhBjm0DQuAAIcrAACC3uBSvgQ5pETdmNRPQQ",
    },
}


def get_subject_lectures(key: str) -> dict[int, str]:
    """تدمج الملازم الثابتة مع المضافة ديناميكياً، وتستبعد ما حذفه المشرف."""
    base = {int(k): v.strip() for k, v in LECTURES.get(key, {}).items() if v}
    try:
        raw = _read_stats_file()
        dyn = raw.get("custom_lectures", {}).get(key, {})
        for k, v in dyn.items():
            num = int(k)
            val = (v or "").strip()
            if val == DELETED_MARKER:
                base.pop(num, None)
            elif val:
                base[num] = val
    except Exception:
        pass
    return base


def subject_max_slot(key: str, for_admin: bool = False) -> int:
    """يحدد عدد أزرار المحاضرات: للطلاب 7 أو أعلى محاضرة مرفوعة، وللمشرف يفتح خانتين إضافيتين دائماً."""
    have = get_subject_lectures(key)
    highest = max(have.keys(), default=0)
    base_max = max(LECTURE_COUNT, highest)
    return base_max + 2 if for_admin else base_max


def lectures_total() -> int:
    """عدد الملازم المرفوعة فعلاً في كل المواد، لعرضه في لوحة التحكم وسجل الإقلاع."""
    return sum(len(get_subject_lectures(k)) for k in LECTURE_SUBJECTS)


def lectures_total_slots() -> int:
    """مجموع الخانات المتاحة حالياً عبر جميع المواد."""
    return sum(subject_max_slot(k, for_admin=False) for k in LECTURE_SUBJECTS)


# ===================== الواجهة =====================

STALE_SUBJECT_MSG = (
    "هذه المادة لم تعد متاحة في البوت.\n"
    "اختر الخدمة والمادة من جديد:"
)

DISABLED_SERVICES: dict[str, set[str]] = {
    "rep": {"t1_ar"},
}


def service_allowed(service: str, code: str | None) -> bool:
    """هل الخدمة متاحة لهذه المادة؟"""
    return code not in DISABLED_SERVICES.get(service, set())


def get_subject(state: dict | None) -> dict | None:
    """يعيد المادة من الحالة، أو None إن كانت الحالة ناقصة أو المادة محذوفة
    أو الخدمة معطّلة لها."""
    if not state or "service" not in state:
        return None
    code = state.get("code")
    if not service_allowed(state["service"], code):
        return None
    return SUBJECTS.get(code)


def main_menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🎓 شرح موضوع أو صفحة ملزمة", callback_data="svc:exp")],
        [InlineKeyboardButton("📝 إنشاء تقرير (Word / PDF)", callback_data="svc:rep")],
        [InlineKeyboardButton("🧮 حل مسألة (نص أو صورة)", callback_data="svc:math")],
        [InlineKeyboardButton("💻 تدقيق كود (نص أو صورة)", callback_data="svc:code")],
        [InlineKeyboardButton("📚 بنك أسئلة", callback_data="svc:quiz")],
        [InlineKeyboardButton("📂 الملازم (محاضرات المواد)", callback_data="lectures")],
        [InlineKeyboardButton("⚙️ الإعدادات", callback_data="settings")],
    ])


def subjects_kb(service: str) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    for term in (1, 2):
        items = [(c, s) for c, s in SUBJECTS.items()
                 if s["term"] == term
                 and (s["is_code"] if service == "code" else True)
                 and service_allowed(service, c)]
        if not items:
            continue
        rows.append([InlineKeyboardButton(
            f"— الكورس {'الأول' if term == 1 else 'الثاني'} —", callback_data="noop")])
        for code, sub in items:
            rows.append([InlineKeyboardButton(sub["name"], callback_data=f"sub:{code}")])
    rows.append([InlineKeyboardButton("🔙 القائمة الرئيسية", callback_data="menu")])
    return InlineKeyboardMarkup(rows)


def has_subjects(service: str) -> bool:
    """هل بقيت مادة واحدة على الأقل لهذه الخدمة بعد الحذف من config؟"""
    return any(
        (s["is_code"] if service == "code" else True) and service_allowed(service, c)
        for c, s in SUBJECTS.items()
    )


def settings_kb(context: ContextTypes.DEFAULT_TYPE) -> InlineKeyboardMarkup:
    cover = context.user_data.get("cover", False)
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(
            f"صفحة الغلاف: {'مفعّلة ✅' if cover else 'معطّلة ❌'}",
            callback_data="cover:toggle")],
        [InlineKeyboardButton("🔙 القائمة الرئيسية", callback_data="menu")],
    ])


def download_kb(context: ContextTypes.DEFAULT_TYPE) -> InlineKeyboardMarkup:
    cover = (context.user_data or {}).get("cover", False)
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📄 تحميل Word", callback_data="dl:word"),
         InlineKeyboardButton("📕 تحميل PDF", callback_data="dl:pdf")],
        [InlineKeyboardButton(
            f"{'إزالة' if cover else 'إضافة'} صفحة الغلاف", callback_data="cover:toggle")],
        [InlineKeyboardButton("🔙 القائمة الرئيسية", callback_data="menu")],
    ])


def lectures_kb() -> InlineKeyboardMarkup:
    """قائمة مواد الملازم، ومع كل مادة عدد ما توفّر منها."""
    rows: list[list[InlineKeyboardButton]] = []
    for key, name in LECTURE_SUBJECTS.items():
        ready = len(get_subject_lectures(key))
        tag = f" ({ready})" if ready else " (لا شيء بعد)"
        rows.append([InlineKeyboardButton(name + tag, callback_data=f"lec_sub:{key}")])
    rows.append([InlineKeyboardButton("🔙 القائمة الرئيسية", callback_data="menu")])
    return InlineKeyboardMarkup(rows)


def lecture_numbers_kb(key: str) -> InlineKeyboardMarkup:
    """أزرار أرقام المحاضرات: أربعة في السطر، وتتمدد تلقائياً إن زادت عن 7."""
    have = get_subject_lectures(key)
    max_slot = subject_max_slot(key, for_admin=False)
    btns = [
        InlineKeyboardButton(f"{'✅' if have.get(n) else '🚫'} {n}",
                             callback_data=f"lec_get:{key}:{n}")
        for n in range(1, max_slot + 1)
    ]
    rows = [btns[i:i + 4] for i in range(0, len(btns), 4)]
    rows.append([InlineKeyboardButton("🔙 المواد", callback_data="lectures")])
    return InlineKeyboardMarkup(rows)


WELCOME = (
    "🎓 *المنصة الأكاديمية الذكية*\n"
    "المرحلة الثانية — علوم الحاسوب\n\n"
    "اختر الخدمة، ثم المادة، ثم أرسل طلبك نصاً أو صورة.\n"
    "ولتحميل ملازم المحاضرات اختر «الملازم»."
)

SERVICE_NOTES = {
    "exp": "تقبل هذه الخدمة اسم الموضوع أو صورة صفحة من الملزمة، "
           "وتقترح لك مقاطع يوتيوب للموضوع.",
    "rep": "تقبل هذه الخدمة عنواناً نصياً، وتُخرج التقرير بصيغة Word أو PDF.",
    "math": "تقبل هذه الخدمة نص المسألة أو صورتها، وتعرض الحل خطوة بخطوة.",
    "code": "تقبل هذه الخدمة الكود نصاً أو صورة، وتعرض الأخطاء والكود المصحّح.\n"
            "(في هذه الخدمة تظهر المواد البرمجية فقط)",
    "quiz": "تقبل هذه الخدمة اسم الموضوع أو الفصل، وتُخرج أسئلة مع مفتاح الإجابات.",
}

DISCLAIMER = (
    "\n\n⚠️ هذا ناتج آلي مساعد للمراجعة ولا يُغني عن المحاضرة الأصلية. "
    "راجع الصفحة، وتحقّق بنفسك من أي أرقام أو حسابات."
)

SERVICE_TEMPERATURE = {
    "rep": 0.6,
    "exp": 0.6,
    "quiz": 0.4,
    "math": 0.2,
    "code": 0.2,
}


# ===================== الأوامر =====================

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("state", None)
    u = update.effective_user
    await record_activity(u.id, None, u.full_name or "", u.username or "")
    await update.message.reply_text(WELCOME, reply_markup=main_menu_kb(),
                                    parse_mode="Markdown")


async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("state", None)
    context.user_data.pop("await_forward", None)
    context.user_data.pop("pending_lecture_fid", None)
    BUSY_USERS.discard(update.effective_user.id)
    await update.message.reply_text("تم إلغاء الطلب الحالي.", reply_markup=main_menu_kb())


async def settings_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    name = context.user_data.get("student", "غير محدد")
    await update.message.reply_text(
        f"الإعدادات\nاسم الطالب على الغلاف: {name}\n"
        f"السنة الدراسية على الغلاف: {academic_year()}\n"
        "لتغيير الاسم: /name الاسم الثلاثي",
        reply_markup=settings_kb(context),
    )


async def name_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    arg = " ".join(context.args).strip()
    if not arg:
        await update.message.reply_text("اكتب الاسم بعد الأمر، مثال:\n/name أحمد علي حسين")
        return
    context.user_data["student"] = arg[:80]
    await update.message.reply_text(f"تم الحفظ: {arg[:80]}")


async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    st = await get_stats()
    ok, why = _pdf_ready()
    bold_ok = os.path.exists(AR_TTF_BOLD)
    mono_ok = os.path.exists(MONO_TTF)
    yt = "مفتاح API مضبوط ✅" if YOUTUBE_API_KEY else "روابط بحث فقط (بلا مفتاح)"
    provs = available_providers()
    prov_txt = " ← ".join(PROVIDER_LABELS.get(p, p) for p in provs) if provs else "لا يوجد ❌"
    lec_slots = lectures_total_slots()
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑 إدارة وحذف الملازم المرفوعة", callback_data="adm_lec_del_menu")]
    ])
    await update.message.reply_text(
        "🛠 لوحة تحكم المنصة الأكاديمية\n\n"
        f"الطلبة المسجلين: {len(st['users'])}\n"
        f"مجموع العمليات: {st['requests_count']}\n"
        f"طلبات قيد التنفيذ الآن: {len(BUSY_USERS)}\n\n"
        f"المواد المتاحة: {len(SUBJECTS)}\n"
        f"الملازم المرفوعة: {lectures_total()} من {lec_slots}\n"
        f"مسار البيانات: {STATS_FILE}\n"
        f"قناة النسخ: {BACKUP_CHAT_ID or 'غير مضبوطة'}\n"
        f"النسخ التلقائي: كل {BACKUP_HOURS} ساعة\n"
        f"السنة الدراسية: {academic_year()}"
        f"{' (مثبّتة يدوياً)' if ACADEMIC_YEAR_ENV else ' (محسوبة تلقائياً)'}\n"
        f"مزوّدو الذكاء: {prov_txt}\n"
        f"Groq: {'مفتاح مضبوط' if GROQ_API_KEY else 'مفقود'} | "
        f"Gemini: {'مفتاح مضبوط' if GEMINI_API_KEY else 'مفقود'}\n"
        f"يوتيوب: {yt}\n"
        f"ذاكرة بحث يوتيوب: {len(_YT_CACHE)} عبارة\n"
        f"نتائج لكل بحث: {YT_FETCH} تُفلتر وتُرتّب إلى {YT_RESULTS}\n"
        f"عبارات لكل طلب: {YT_MAX_QUERIES} (كل عبارة = 100 وحدة)\n"
        f"لغة الشرح: {'العربية دائماً' if EXPLAIN_IN_ARABIC else 'حسب لغة المادة'}\n"
        f"الرسائل المتراكمة: {'تُحذف عند الإقلاع' if DROP_PENDING_UPDATES else 'تُعالج بعد العودة'}\n"
        f"PDF: {'جاهز ✅' if ok else 'غير جاهز ❌ — ' + why}\n"
        f"الخط العادي: {'موجود' if os.path.exists(AR_TTF) else 'مفقود'}\n"
        f"الخط العريض: {'موجود' if bold_ok else 'بديل (يستخدم العادي)'}\n"
        f"خط الكود: {'موجود' if mono_ok else 'بديل (يستخدم العادي)'}\n"
        f"محرّك التشكيل: {'uharfbuzz' if FPDF_OK else '-'} / bidi={BIDI_OK}\n"
        f"ضغط الصور: {'مفعّل' if PIL_OK else 'معطّل'}\n"
        f"حد الصور للطلب: {MAX_IMAGES}\n\n"
        "لإضافة ملزمة: أرسلها هنا كملف واختر المادة ورقم المحاضرة من الأزرار.\n"
        "الأوامر: /stats /users /whois /backup /restore /getid /broadcast",
        reply_markup=kb,
    )


async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    st = await get_stats()
    for part in split_for_telegram(build_stats_text(st)):
        await update.message.reply_text(part)


async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    st = await get_stats()
    rows = ["id\tname\tusername\trequests\tfirst_seen\tlast_seen"]
    for uid, r in sorted(st["users"].items(), key=lambda x: -x[1].get("requests", 0)):
        rows.append("\t".join([uid, r.get("name", ""), r.get("username", ""),
                               str(r.get("requests", 0)),
                               r.get("first_seen", ""), r.get("last_seen", "")]))
    buf = io.BytesIO("\n".join(rows).encode("utf-8"))
    await update.message.reply_document(document=buf, filename="users.tsv",
                                        caption=f"عدد الطلبة: {len(st['users'])}")


async def whois_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """يسحب صورة الحساب ومعلوماته الكاملة من تلغرام عبر الـ ID حتى لو كان بدون يوزر."""
    if update.effective_user.id != ADMIN_ID:
        return
    if not context.args:
        await update.message.reply_text("اكتب الـ ID بعد الأمر، مثال:\n/whois 123456789")
        return

    uid_str = context.args[0].strip()
    try:
        uid = int(uid_str)
    except ValueError:
        await update.message.reply_text("الـ ID يجب أن يكون أرقاماً فقط.")
        return

    st = await get_stats()
    rec = st["users"].get(uid_str, {})

    try:
        chat = await context.bot.get_chat(uid)
        full_name = html.escape(chat.full_name or rec.get("name") or "بدون اسم")
        uname = f"@{chat.username}" if chat.username else "لا يوجد (مخفي)"
        bio = html.escape(chat.bio or "لا توجد نبذة (Bio)")
    except Exception as e:
        full_name = html.escape(rec.get("name") or "غير معروف")
        uname = f"@{rec['username']}" if rec.get("username") else "لا يوجد"
        bio = f"تعذّر الجلب المباشر ({e})"

    info_text = (
        f"👤 <b>بطاقة المستخدم من تلغرام:</b>\n\n"
        f"• <b>الاسم الحالي:</b> <a href='tg://user?id={uid}'>{full_name}</a>\n"
        f"• <b>الاسم المسجل بالبوت:</b> {html.escape(rec.get('name', '-'))}\n"
        f"• <b>المعرّف (ID):</b> <code>{uid}</code>\n"
        f"• <b>اليوزر:</b> {uname}\n"
        f"• <b>النبذة (Bio):</b> {bio}\n"
        f"• <b>عدد الطلبات:</b> {rec.get('requests', 0)}\n"
        f"• <b>أول دخول:</b> {rec.get('first_seen', 'غير مسجل')}\n"
        f"• <b>آخر ظهور:</b> {rec.get('last_seen', 'غير مسجل')}\n\n"
        f"🔗 <a href='tg://user?id={uid}'>اضغط هنا لمحاولة فتح البروفايل</a>"
    )

    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 فتح البروفايل في تلغرام", url=f"tg://user?id={uid}")]
    ])

    try:
        photos = await context.bot.get_user_profile_photos(uid, limit=1)
        if photos and photos.total_count > 0:
            photo_fid = photos.photos[0][-1].file_id
            await update.message.reply_photo(
                photo=photo_fid,
                caption=info_text,
                parse_mode="HTML",
                reply_markup=kb,
            )
            return
    except Exception:
        pass

    await update.message.reply_text(info_text, parse_mode="HTML", reply_markup=kb)


async def backup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    if await do_backup(context.bot, "يدوي"):
        mark_clean()
        await update.message.reply_text("تم إرسال النسخة إلى القناة وتثبيتها ✅")
        return
    st = await get_stats()
    blob = json.dumps(st, ensure_ascii=False, indent=1).encode("utf-8")
    await update.message.reply_document(
        document=io.BytesIO(blob), filename="bot_stats.json",
        caption="قناة النسخ غير مضبوطة أو فشل الإرسال، احفظ هذا الملف يدوياً.")


async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    data = await restore_stats(context.bot)
    if data:
        await update.message.reply_text(
            f"تمت الاستعادة ✅\nالطلبة: {len(data['users'])}\n"
            f"العمليات: {data['requests_count']}\n"
            f"الملازم المتاحة: {lectures_total()}")
    else:
        await update.message.reply_text(
            "لم أجد نسخة مثبّتة في القناة. يمكنك إرسال ملف bot_stats.json هنا مباشرة.")


async def getid_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    context.user_data["await_forward"] = True
    await update.message.reply_text(
        "أعد توجيه أي رسالة من القناة الآن، وسأخبرك بمعرّفها.")


async def _broadcast_worker(bot, ids: list[int], text: str, admin_id: int) -> None:
    sent, failed, blocked = 0, 0, []
    for uid in ids:
        try:
            await bot.send_message(uid, text)
            sent += 1
        except Forbidden:
            blocked.append(uid)
            failed += 1
        except RetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
            try:
                await bot.send_message(uid, text)
                sent += 1
            except Exception:
                failed += 1
        except Exception:
            failed += 1
        await asyncio.sleep(0.03)
    await remove_users(blocked)
    try:
        await bot.send_message(
            admin_id,
            f"انتهى الإرسال الجماعي.\nنجح: {sent}\nفشل: {failed}\n"
            f"حُذف من القائمة (حاظرون): {len(blocked)}",
        )
    except Exception:
        pass


async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != ADMIN_ID:
        return
    text = " ".join(context.args).strip()
    if not text:
        await update.message.reply_text("اكتب النص بعد الأمر:\n/broadcast مرحباً بالطلبة")
        return
    st = await get_stats()
    ids = [int(u) for u in st["users"].keys() if str(u) != str(ADMIN_ID)]
    await update.message.reply_text(f"بدأ الإرسال إلى {len(ids)} مستخدم…")
    asyncio.create_task(_broadcast_worker(context.bot, ids, text, ADMIN_ID))


# ===================== أزرار الملازم =====================

async def lectures_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    context.user_data.pop("state", None)
    await q.edit_message_text(
        "📂 الملازم والمحاضرات\n\n"
        "اختر المادة، ثم رقم المحاضرة.\n"
        "الرقم بين قوسين هو عدد المحاضرات المتوفرة حالياً، "
        "والبقية تُضاف عند وصولها من الدكاترة.",
        reply_markup=lectures_kb(),
    )


async def lecture_subject(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    key = q.data.split(":", 1)[1]
    name = LECTURE_SUBJECTS.get(key)
    if name is None:
        await q.edit_message_text("هذه المادة لم تعد متاحة.", reply_markup=lectures_kb())
        return
    ready = len(get_subject_lectures(key))
    max_slot = subject_max_slot(key, for_admin=False)
    await q.edit_message_text(
        f"📖 {name}\n"
        f"المتوفر: {ready} من {max_slot}\n\n"
        "اختر رقم المحاضرة:\n"
        "✅ متوفرة   🚫 غير متوفرة بعد",
        reply_markup=lecture_numbers_kb(key),
    )


async def lecture_send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    try:
        _, key, num_raw = q.data.split(":")
        num = int(num_raw)
    except ValueError:
        await q.answer()
        return

    name = LECTURE_SUBJECTS.get(key, "")
    fid = get_subject_lectures(key).get(num, "").strip()
    if not fid:
        await q.answer(
            "هذه المحاضرة غير متوفرة بعد، ستُضاف حين تصل من الدكتور.",
            show_alert=True)
        return

    await q.answer("جارٍ الإرسال…")
    u = q.from_user
    chat_id = q.message.chat_id
    try:
        await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
        await context.bot.send_document(
            chat_id=chat_id,
            document=fid,
            caption=f"📚 {name} — المحاضرة {num}",
        )
        await record_activity(u.id, None, u.full_name or "", u.username or "")
    except Exception:
        logger.exception("failed to send lecture %s/%s", key, num)
        await context.bot.send_message(
            chat_id,
            "تعذّر إرسال الملزمة. إن تكرر ذلك فقد يكون معرّف الملف قديماً، "
            "أبلغ المشرف بذلك.")


async def admin_lecture_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """معالج خاص بالمشرف لإضافة أو استبدال أو حذف الملازم مباشرة."""
    q = update.callback_query
    if q.from_user.id != ADMIN_ID:
        await q.answer("غير مصرح لك.", show_alert=True)
        return

    data = q.data or ""

    if data == "adm_lec_cancel":
        context.user_data.pop("pending_lecture_fid", None)
        await q.answer("تم الإغلاق")
        await q.edit_message_text("تم إغلاق نافذة إدارة الملازم.")
        return

    # ─── مسار حذف الملازم ───
    if data == "adm_lec_del_menu":
        await q.answer()
        rows = []
        for key, name in LECTURE_SUBJECTS.items():
            cnt = len(get_subject_lectures(key))
            if cnt > 0:
                rows.append([InlineKeyboardButton(
                    f"{name} ({cnt} متوفرة)", callback_data=f"adm_lec_delsub:{key}")])
        if not rows:
            await q.edit_message_text("لا توجد أي ملازم مرفوعة حالياً لحذفها.")
            return
        rows.append([InlineKeyboardButton("❌ إغلاق", callback_data="adm_lec_cancel")])
        await q.edit_message_text(
            "🗑 <b>حذف ملزمة</b>\nاختر المادة التي تريد حذف ملزمة منها:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(rows),
        )
        return

    if data.startswith("adm_lec_delsub:"):
        await q.answer()
        key = data.split(":", 1)[1]
        name = LECTURE_SUBJECTS.get(key, key)
        have = get_subject_lectures(key)
        if not have:
            await q.edit_message_text("لا توجد ملازم مرفوعة في هذه المادة.")
            return
        btns = [
            InlineKeyboardButton(f"🗑 محاضرة {n}", callback_data=f"adm_lec_deldo:{key}:{n}")
            for n in sorted(have.keys())
        ]
        rows = [btns[i:i + 3] for i in range(0, len(btns), 3)]
        rows.append([InlineKeyboardButton("🔙 رجوع للمواد", callback_data="adm_lec_del_menu")])
        await q.edit_message_text(
            f"📖 المادة: <b>{name}</b>\nاضغط على رقم المحاضرة لحذفها فوراً:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(rows),
        )
        return

    if data.startswith("adm_lec_deldo:"):
        _, key, num_str = data.split(":")
        num = int(num_str)
        name = LECTURE_SUBJECTS.get(key, key)

        async with _stats_lock:
            st_data = await asyncio.to_thread(_read_stats_file)
            custom = st_data.setdefault("custom_lectures", {})
            subj_dict = custom.setdefault(key, {})
            subj_dict[str(num)] = DELETED_MARKER
            await asyncio.to_thread(_write_stats_file, st_data)

        await q.answer("تم حذف الملزمة 🗑")
        backed_up = await do_backup(context.bot, f"حذف ملزمة {name} #{num}")
        if backed_up:
            mark_clean()
            backup_msg = "وتم تحديث النسخة الاحتياطية في القناة ✅"
        else:
            backup_msg = "⚠️ تذكّر عمل /backup لتحديث النسخة الاحتياطية."

        await q.edit_message_text(
            f"✅ تم حذف <b>المحاضرة {num}</b> من مادة <b>{name}</b> بنجاح.\n{backup_msg}",
            parse_mode="HTML",
        )
        return

    # ─── مسار إضافة / استبدال ملزمة (يفتح أرقاماً جديدة تلقائياً) ───
    fid = context.user_data.get("pending_lecture_fid")
    if not fid:
        await q.answer("انتهت صلاحية الملف، أرسله مرة أخرى.", show_alert=True)
        return

    if data.startswith("adm_lec_sub:"):
        await q.answer()
        key = data.split(":", 1)[1]
        name = LECTURE_SUBJECTS.get(key, key)
        have = get_subject_lectures(key)
        max_slot = subject_max_slot(key, for_admin=True)

        btns = [
            InlineKeyboardButton(
                f"{'🔄' if have.get(n) else '➕'} محاضرة {n}",
                callback_data=f"adm_lec_save:{key}:{n}"
            )
            for n in range(1, max_slot + 1)
        ]
        rows = [btns[i:i + 4] for i in range(0, len(btns), 4)]
        rows.append([InlineKeyboardButton("❌ إلغاء", callback_data="adm_lec_cancel")])

        await q.edit_message_text(
            f"📖 المادة: <b>{name}</b>\n"
            "اختر رقم المحاضرة لحفظ الملف فيها:\n"
            "(➕ = إضافة جديدة | 🔄 = استبدال ملزمة موجودة)\n"
            "💡 <i>تُفتح أرقام محاضرات إضافية تلقائياً كلما زاد عدد الملازم.</i>",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(rows),
        )
        return

    if data.startswith("adm_lec_save:"):
        _, key, num_str = data.split(":")
        num = int(num_str)
        name = LECTURE_SUBJECTS.get(key, key)

        async with _stats_lock:
            st_data = await asyncio.to_thread(_read_stats_file)
            custom = st_data.setdefault("custom_lectures", {})
            subj_dict = custom.setdefault(key, {})
            subj_dict[str(num)] = fid
            await asyncio.to_thread(_write_stats_file, st_data)

        context.user_data.pop("pending_lecture_fid", None)
        await q.answer("تم الحفظ بنجاح ✅")

        backed_up = await do_backup(context.bot, f"إضافة ملزمة {name} #{num}")
        if backed_up:
            mark_clean()
            backup_msg = "وتم حفظ نسخة احتياطية في القناة تلقائياً ✅"
        else:
            backup_msg = "⚠️ تذكّر عمل /backup لحفظها في القناة."

        await q.edit_message_text(
            f"✅ <b>تم تفعيل الملزمة للطلاب فوراً!</b>\n\n"
            f"المادة: {name}\n"
            f"المحاضرة: {num}\n"
            f"{backup_msg}",
            parse_mode="HTML",
        )


# ===================== الأزرار =====================

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    data = q.data or ""
    await q.answer()

    if data == "noop":
        return

    if data == "menu":
        context.user_data.pop("state", None)
        await q.edit_message_text(WELCOME, reply_markup=main_menu_kb(),
                                  parse_mode="Markdown")
        return

    if data == "settings":
        await q.edit_message_text(
            "الإعدادات\n\nصفحة الغلاف تُضاف للتقرير أو الشرح فقط عند تفعيلها، "
            "والوضع الافتراضي هو النص وحده.\n"
            f"السنة الدراسية الحالية على الغلاف: {academic_year()}\n"
            "لتعيين اسمك على الغلاف: /name الاسم الثلاثي",
            reply_markup=settings_kb(context),
        )
        return

    if data == "cover:toggle":
        context.user_data["cover"] = not context.user_data.get("cover", False)
        on = context.user_data["cover"]
        if context.user_data.get("last_report"):
            await q.edit_message_text(
                f"صفحة الغلاف أصبحت {'مفعّلة ✅' if on else 'معطّلة ❌'}.\n"
                "اختر صيغة التحميل:",
                reply_markup=download_kb(context),
            )
        else:
            await q.edit_message_text(
                f"صفحة الغلاف أصبحت {'مفعّلة ✅' if on else 'معطّلة ❌'}.",
                reply_markup=settings_kb(context),
            )
        return

    if data.startswith("svc:"):
        svc = data.split(":", 1)[1]
        if svc not in SYS_BUILDERS:
            return
        if not has_subjects(svc):
            await q.edit_message_text(
                "لا توجد مواد متاحة لهذه الخدمة حالياً.",
                reply_markup=main_menu_kb())
            return
        context.user_data["state"] = {"service": svc}
        note = SERVICE_NOTES.get(svc, "")
        await q.edit_message_text(
            f"الخدمة: {SERVICE_LABELS.get(svc, svc)}\n"
            f"{note}\n\nاختر المادة:\n\nللإلغاء في أي وقت: /cancel",
            reply_markup=subjects_kb(svc),
        )
        return

    if data.startswith("sub:"):
        code = data.split(":", 1)[1]
        sub = SUBJECTS.get(code)
        state = context.user_data.get("state") or {}
        if sub is None:
            context.user_data.pop("state", None)
            await q.edit_message_text(STALE_SUBJECT_MSG, reply_markup=main_menu_kb())
            return
        if "service" not in state:
            await q.edit_message_text("انتهت صلاحية الاختيار. ابدأ من جديد.",
                                      reply_markup=main_menu_kb())
            return
        if not service_allowed(state["service"], code):
            context.user_data.pop("state", None)
            await q.edit_message_text(
                "هذه الخدمة غير متاحة لهذه المادة.\n"
                "اختر الخدمة والمادة من جديد:",
                reply_markup=main_menu_kb())
            return
        state.update({"code": code, "subject": sub["name"], "lang": sub["lang"]})
        context.user_data["state"] = state
        svc = state["service"]
        prompts = {
            "rep": "أرسل عنوان التقرير المطلوب.",
            "exp": ("أرسل اسم الموضوع الذي تريد شرحه،\n"
                    f"أو صوّر صفحة من الملزمة وأرسل الصورة (حتى {MAX_IMAGES} صور كألبوم).\n"
                    "يمكنك إضافة تعليق مع الصورة مثل: «اشرح لي الجزء الثاني فقط»."),
            "math": f"أرسل نص المسألة، أو صوّرها وأرسل الصورة (حتى {MAX_IMAGES} صور كألبوم).",
            "code": ("أرسل الكود نصاً (الأفضل والأدق)، أو صورة له.\n"
                     "ملاحظة: تصوير الكود من محرر بواجهة عربية قد يقلب مواضع الأقواس."),
            "quiz": "أرسل الموضوع أو الفصل المطلوب إنشاء أسئلة عنه.",
        }
        await q.edit_message_text(
            f"المادة: {sub['name']}\n\n"
            f"{prompts[svc]}\n\nللإلغاء: /cancel"
        )
        return

    if data.startswith("dl:"):
        fmt = data.split(":", 1)[1]
        rep = context.user_data.get("last_report")
        if not rep:
            await q.edit_message_text("لا يوجد ملف محفوظ. أنشئ تقريراً أو شرحاً جديداً.",
                                      reply_markup=main_menu_kb())
            return
        cover = context.user_data.get("cover", False)
        student = context.user_data.get("student", "")
        base = safe_filename(rep["title"])
        chat_id = q.message.chat_id
        try:
            if fmt == "word":
                await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
                blob = await asyncio.to_thread(
                    create_academic_docx, rep["title"], rep["body"], rep["subject"],
                    rep["lang"], cover, student)
                await context.bot.send_document(
                    chat_id, document=io.BytesIO(blob), filename=f"{base}.docx",
                    caption=f"📄 {rep['title']}")
                await record_file("word")
            else:
                ok, why = _pdf_ready(rep["lang"])
                if not ok:
                    await context.bot.send_message(chat_id, f"تعذّر إنشاء PDF: {why}")
                    return
                await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
                blob = await asyncio.to_thread(
                    create_academic_pdf, rep["title"], rep["body"], rep["subject"],
                    rep["lang"], cover, student)
                await context.bot.send_document(
                    chat_id, document=io.BytesIO(blob), filename=f"{base}.pdf",
                    caption=f"📕 {rep['title']}")
                await record_file("pdf")
        except Exception as e:
            logger.exception("download failed")
            if ADMIN_ID and q.from_user.id == ADMIN_ID:
                await context.bot.send_message(
                    chat_id, f"تفاصيل الخطأ:\n{type(e).__name__}: {e}"[:3500])
            else:
                await context.bot.send_message(
                    chat_id, "حدث خطأ أثناء توليد الملف. جرّب الصيغة الأخرى.")
        return


# ===================== معالجة النص =====================

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    msg = update.message
    text = (msg.text or "").strip()
    if not text:
        return

    if ADMIN_ID and user.id != ADMIN_ID:
        try:
            safe_name = html.escape(user.full_name or "طالب")
            uname_str = f"@{user.username}" if user.username else f"ID: <code>{user.id}</code>"
            await context.bot.send_message(
                chat_id=ADMIN_ID,
                text=(f"📩 طلب جديد من: <a href='tg://user?id={user.id}'>{safe_name}</a> ({uname_str})\n\n"
                      f"النص:\n{html.escape(text[:1000])}"),
                parse_mode="HTML",
            )
        except Exception:
            pass

    state = context.user_data.get("state")

    if not state or "service" not in state or "code" not in state:
        await msg.reply_text("اختر الخدمة والمادة أولاً:", reply_markup=main_menu_kb())
        return

    sub = get_subject(state)
    if sub is None:
        context.user_data.pop("state", None)
        await msg.reply_text(STALE_SUBJECT_MSG, reply_markup=main_menu_kb())
        return

    if user.id in BUSY_USERS:
        await msg.reply_text("طلبك السابق قيد المعالجة، انتظر انتهاءه أو أرسل /cancel")
        return

    if len(text) > MAX_INPUT_CHARS:
        await msg.reply_text(f"النص طويل جداً ({len(text)} حرف). الحد {MAX_INPUT_CHARS} حرف.")
        return

    svc = state["service"]
    lang = explain_lang(sub) if svc == "exp" else sub["lang"]

    BUSY_USERS.add(user.id)
    notice = None
    try:
        notice = await msg.reply_text("⏳ جارٍ المعالجة…")
        await context.bot.send_chat_action(msg.chat_id, ChatAction.TYPING)
        system = SYS_BUILDERS[svc](sub)
        merge = svc in ("math", "quiz")

        if svc == "rep":
            user_msg = (f"Report title: {text}" if sub["lang"] == "en"
                        else f"عنوان التقرير: {text}")
        elif svc == "exp":
            user_msg = (f"الموضوع المطلوب شرحه: {text}\n"
                        f"المادة: {sub['name']} — {sub.get('en', '')}")
        else:
            user_msg = text

        raw = await query_ai_text(system, user_msg, merge_system=merge,
                                  temperature=SERVICE_TEMPERATURE.get(svc, 0.4))

        if svc == "code":
            body = clean_code_output(raw)
        elif svc in ("rep", "exp"):
            body = clean_report_text(raw, lang)
            if lang == "ar" and (svc == "rep" or sub["lang"] == "ar"):
                body = strip_latin_parentheticals(body)
        else:
            body = clean_text_strictly(raw, lang)
            if lang == "ar":
                body = strip_latin_parentheticals(body)

        if not body:
            raise RuntimeError("ناتج فارغ بعد التنظيف")

        if notice is not None:
            try:
                await notice.delete()
            except Exception:
                pass

        if svc == "rep":
            context.user_data["last_report"] = {
                "title": text[:120], "body": body,
                "subject": sub["name"], "lang": lang,
            }
            cover = context.user_data.get("cover", False)
            await msg.reply_text(
                f"✅ تم إنشاء التقرير: {text[:120]}\n"
                f"المادة: {sub['name']}\n"
                f"عدد الكلمات ≈ {len(body.split())}\n"
                f"صفحة الغلاف: {'مفعّلة' if cover else 'معطّلة'}\n\n"
                "اختر صيغة التحميل:" + DISCLAIMER,
                reply_markup=download_kb(context),
            )

        elif svc == "exp":
            topic = text[:120]
            context.user_data["last_report"] = {
                "title": f"شرح: {topic}"[:120], "body": body,
                "subject": sub["name"], "lang": lang,
            }
            for part in split_for_telegram(headings_for_telegram(body)):
                await msg.reply_text(part)
            try:
                yt = await build_youtube_block(topic, sub)
                if yt:
                    await send_plain(context.bot, msg.chat_id, yt, parse_mode="HTML")
            except Exception as e:
                logger.warning("youtube block failed: %s", e)
            await msg.reply_text(
                "يمكنك تحميل الشرح كملف للمراجعة، أو إرسال موضوع آخر لشرحه:"
                + DISCLAIMER,
                reply_markup=download_kb(context),
            )

        else:
            for part in split_for_telegram(body):
                await msg.reply_text(part)
            await msg.reply_text("للطلب التالي اختر من القائمة:" + DISCLAIMER,
                                 reply_markup=main_menu_kb())

        await record_activity(user.id, svc, user.full_name or "",
                              user.username or "", state["code"])

    except Exception:
        logger.exception("handle_text failed")
        fail_txt = "تعذّر إكمال الطلب حالياً. أعد المحاولة بعد قليل."
        try:
            if notice is not None:
                await notice.edit_text(fail_txt)
            else:
                await msg.reply_text(fail_txt)
        except Exception:
            try:
                await msg.reply_text(fail_txt)
            except Exception:
                pass
    finally:
        BUSY_USERS.discard(user.id)
        if svc not in ("rep", "exp"):
            context.user_data.pop("state", None)


# ===================== معالجة الصور =====================

async def run_image_request(context: ContextTypes.DEFAULT_TYPE, chat_id: int, user_id: int,
                            state: dict, file_ids: list[str], caption: str,
                            name: str = "", username: str = "") -> None:
    sub = get_subject(state)
    if sub is None:
        await context.bot.send_message(chat_id, STALE_SUBJECT_MSG,
                                       reply_markup=main_menu_kb())
        return
    svc = state["service"]
    lang = explain_lang(sub) if svc == "exp" else sub["lang"]

    if user_id in BUSY_USERS:
        await context.bot.send_message(chat_id, "طلبك السابق قيد المعالجة، انتظر انتهاءه.")
        return

    BUSY_USERS.add(user_id)
    wait_note = ("🔍 جارٍ قراءة الصفحة وتحضير الشرح…" if svc == "exp"
                 else "🔍 جارٍ قراءة الصورة ومعالجتها…")
    notice = None
    try:
        notice = await context.bot.send_message(chat_id, wait_note)
        await context.bot.send_chat_action(chat_id, ChatAction.TYPING)
        images = await fetch_images(context.bot, file_ids)
        if not images:
            raise RuntimeError("لا توجد صور صالحة")

        parts = [SYS_BUILDERS[svc](sub), ""]
        parts.append(OCR_EXPLAIN if svc == "exp" else OCR_PREFIX)
        if svc == "code":
            parts.append(OCR_CODE_EXTRA)
        if caption:
            parts.append(("تعليق الطالب: " if lang == "ar" else "Student note: ") + caption)
        prompt = "\n\n".join(parts)

        raw = await query_ai_vision(prompt, images)

        if svc == "code":
            body = clean_code_output(raw)
        elif svc == "exp":
            body = clean_report_text(raw, lang)
            if lang == "ar" and sub["lang"] == "ar":
                body = strip_latin_parentheticals(body)
        else:
            body = clean_text_strictly(raw, lang)
            if lang == "ar":
                body = strip_latin_parentheticals(body)

        if not body:
            raise RuntimeError("ناتج فارغ")

        if notice is not None:
            try:
                await notice.delete()
            except Exception:
                pass

        if svc == "exp":
            topic = extract_topic(body, caption or sub["name"])
            kb = main_menu_kb()
            try:
                if context.user_data is not None:
                    context.user_data["last_report"] = {
                        "title": f"شرح: {topic}"[:120], "body": body,
                        "subject": sub["name"], "lang": lang,
                    }
                    kb = download_kb(context)
            except Exception as e:
                logger.warning("cannot store explanation: %s", e)

            for part in split_for_telegram(headings_for_telegram(body)):
                await context.bot.send_message(chat_id, part)
            try:
                yt = await build_youtube_block(topic, sub)
                if yt:
                    await send_plain(context.bot, chat_id, yt, parse_mode="HTML")
            except Exception as e:
                logger.warning("youtube block failed: %s", e)
            await context.bot.send_message(
                chat_id,
                "يمكنك تحميل الشرح كملف، أو إرسال صفحة أخرى لشرحها:" + DISCLAIMER,
                reply_markup=kb,
            )
        else:
            for part in split_for_telegram(body):
                await context.bot.send_message(chat_id, part)
            await context.bot.send_message(chat_id,
                                           "للطلب التالي اختر من القائمة:" + DISCLAIMER,
                                           reply_markup=main_menu_kb())

        await record_activity(user_id, f"{svc}_img", name, username, state["code"])

    except VisionUnavailable:
        logger.exception("vision chain unavailable")
        txt = ("خدمة قراءة الصور غير متاحة الآن. "
               "أرسل السؤال أو الكود مكتوباً نصاً وسيعمل الطلب عادياً.")
        try:
            if notice is not None:
                await notice.edit_text(txt)
            else:
                await context.bot.send_message(chat_id, txt)
        except Exception:
            try:
                await context.bot.send_message(chat_id, txt)
            except Exception:
                pass
    except Exception:
        logger.exception("image request failed")
        txt = "تعذّرت قراءة الصورة. جرّب صورة أوضح أو أرسل النص مكتوباً."
        try:
            if notice is not None:
                await notice.edit_text(txt)
            else:
                await context.bot.send_message(chat_id, txt)
        except Exception:
            try:
                await context.bot.send_message(chat_id, txt)
            except Exception:
                pass
    finally:
        BUSY_USERS.discard(user_id)
        if svc not in ("rep", "exp") and context.user_data is not None:
            context.user_data.pop("state", None)


async def _album_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    job = context.job
    key = job.data["key"]
    state = job.data["state"]
    buf = (context.chat_data or {}).pop(key, None)
    if not buf:
        return
    await run_image_request(context, job.chat_id, job.user_id, state,
                            buf["file_ids"], buf.get("caption", ""),
                            job.data.get("name", ""), job.data.get("username", ""))


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.message
    user = update.effective_user
    state = context.user_data.get("state")

    if not state or "service" not in state or "code" not in state:
        await msg.reply_text("اختر الخدمة والمادة أولاً ثم أرسل الصورة:",
                             reply_markup=main_menu_kb())
        return

    if get_subject(state) is None:
        context.user_data.pop("state", None)
        await msg.reply_text(STALE_SUBJECT_MSG, reply_markup=main_menu_kb())
        return

    svc = state["service"]
    if svc == "rep":
        await msg.reply_text("خدمة التقرير تحتاج عنواناً نصياً. أرسل عنوان التقرير.")
        return
    if svc == "quiz":
        await msg.reply_text("خدمة بنك الأسئلة تحتاج اسم الموضوع نصاً.")
        return

    fid = msg.photo[-1].file_id
    caption = (msg.caption or "").strip()

    if msg.media_group_id:
        key = f"album:{msg.media_group_id}"
        buf = context.chat_data.setdefault(key, {"file_ids": [], "caption": ""})
        if len(buf["file_ids"]) < MAX_IMAGES:
            buf["file_ids"].append(fid)
        if caption and not buf["caption"]:
            buf["caption"] = caption
        if context.job_queue:
            for j in context.job_queue.get_jobs_by_name(key):
                j.schedule_removal()
            context.job_queue.run_once(
                _album_job, ALBUM_WAIT, name=key,
                chat_id=msg.chat_id, user_id=user.id,
                data={"key": key, "state": state,
                      "name": user.full_name or "", "username": user.username or ""},
            )
        return

    await run_image_request(context, msg.chat_id, user.id, state, [fid], caption,
                            user.full_name or "", user.username or "")


async def handle_forwarded(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """يعمل بعد /getid فقط؛ وإلا يمرّر الرسالة إلى معالجها الطبيعي."""
    msg = update.message
    if not context.user_data.pop("await_forward", False):
        if msg.photo:
            await handle_photo(update, context)
        elif msg.document:
            await handle_document(update, context)
        elif msg.text:
            await handle_text(update, context)
        return
    origin = getattr(msg, "forward_origin", None)
    src = getattr(origin, "chat", None) or getattr(origin, "sender_chat", None)
    lines = [f"معرّف هذه الدردشة: {msg.chat_id}"]
    if src is not None:
        lines += [f"نوع المصدر: {src.type}",
                  f"العنوان: {src.title or '-'}",
                  f"معرّف المصدر: {src.id}", "",
                  "ضع معرّف المصدر في المتغيّر BACKUP_CHAT_ID"]
    else:
        lines.append("تعذّر قراءة معرّف المصدر. أعد التوجيه من القناة نفسها.")
    await msg.reply_text("\n".join(lines))


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.message
    doc = msg.document
    is_admin = update.effective_user.id == ADMIN_ID

    # استعادة الإحصائيات من ملف JSON مرسل يدوياً
    if is_admin and doc and (doc.file_name or "").lower().endswith(".json"):
        try:
            f = await doc.get_file()
            raw = bytes(await f.download_as_bytearray())
            incoming = json.loads(raw.decode("utf-8"))
            async with _stats_lock:
                data = await asyncio.to_thread(_read_stats_file)
                data = _merge_stats(data, incoming)
                await asyncio.to_thread(_write_stats_file, data)
            await msg.reply_text(
                f"تم دمج النسخة ✅\nالطلبة: {len(data['users'])}\n"
                f"العمليات: {data['requests_count']}\n"
                f"الملازم المتاحة: {lectures_total()}")
        except Exception as e:
            await msg.reply_text(f"فشل قراءة الملف: {e}")
        return

    # أي ملف آخر من المشرف: نحفظ معرّفه مؤقتاً ونعرض أزرار إضافته للملازم مباشرة
    if is_admin and doc:
        size_mb = (doc.file_size or 0) / 1048576
        fname = html.escape(doc.file_name or "-")
        context.user_data["pending_lecture_fid"] = doc.file_id.strip()

        rows = [
            [InlineKeyboardButton(name, callback_data=f"adm_lec_sub:{key}")]
            for key, name in LECTURE_SUBJECTS.items()
        ]
        rows.append([InlineKeyboardButton("❌ إلغاء", callback_data="adm_lec_cancel")])

        await msg.reply_text(
            "🆔 <b>تم استلام الملف بنجاح</b>\n"
            f"الاسم: {fname} ({size_mb:.2f} MB)\n"
            f"<code>{html.escape(doc.file_id)}</code>\n\n"
            "📌 <b>اختر المادة لإضافة هذه الملزمة إليها مباشرة:</b>",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(rows),
        )
        return

    await msg.reply_text("الملفات غير مدعومة. أرسل النص مكتوباً أو صورة للصفحة أو المسألة.")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    if isinstance(context.error, Conflict):
        logger.warning("conflict: another instance is polling the same token")
        return
    logger.error("unhandled error", exc_info=context.error)


# ===================== خادم الفحص =====================

class SimpleHealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        body = b'{"status":"ok","service":"academic_bot"}'
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, *args):
        return


def run_http_server() -> None:
    try:
        ThreadingHTTPServer(("0.0.0.0", PORT), SimpleHealthHandler).serve_forever()
    except Exception as e:
        logger.error("health server stopped: %s", e)


# ===================== الإقلاع والإيقاف =====================

async def post_init(app) -> None:
    data = await get_stats()
    if not data["users"] and not data.get("custom_lectures"):
        restored = await restore_stats(app.bot)
        if restored:
            data = restored
    if app.job_queue and BACKUP_CHAT_ID:
        app.job_queue.run_repeating(backup_job, interval=int(BACKUP_HOURS * 3600),
                                    first=180, name="stats_backup")
        logger.info("auto backup every %s hours", BACKUP_HOURS)
    if ADMIN_ID:
        provs = available_providers()
        prov_txt = " ← ".join(PROVIDER_LABELS.get(p, p) for p in provs) if provs else "لا يوجد"
        try:
            await app.bot.send_message(
                ADMIN_ID,
                f"البوت يعمل الآن ✅\nالطلبة: {len(data['users'])}\n"
                f"العمليات: {data['requests_count']}\n"
                f"المواد المتاحة: {len(SUBJECTS)}\n"
                f"الملازم المرفوعة: {lectures_total()}\n"
                f"السنة الدراسية: {academic_year()}\n"
                f"مزوّدو الذكاء: {prov_txt}\n"
                f"يوتيوب: {'مفتاح مضبوط' if YOUTUBE_API_KEY else 'روابط بحث فقط'}")
        except Exception:
            pass


async def post_stop(app) -> None:
    await do_backup(app.bot, "إيقاف")


# ===================== التشغيل =====================

def build_app() -> Application:
    """يبني تطبيقاً جديداً بكل معالجاته."""
    persistence = PicklePersistence(filepath=PERSIST_FILE)
    app = (
        Application.builder()
        .token(TELEGRAM_BOT_TOKEN)
        .persistence(persistence)
        .post_init(post_init)
        .post_stop(post_stop)
        .concurrent_updates(True)
        .build()
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("cancel", cancel_command))
    app.add_handler(CommandHandler("settings", settings_command))
    app.add_handler(CommandHandler("name", name_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("users", users_command))
    app.add_handler(CommandHandler("whois", whois_command))
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("restore", restore_command))
    app.add_handler(CommandHandler("getid", getid_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))

    # أزرار الملازم ومعالج المشرف قبل المعالج العام
    app.add_handler(CallbackQueryHandler(lectures_menu, pattern=r"^lectures$"))
    app.add_handler(CallbackQueryHandler(lecture_subject, pattern=r"^lec_sub:"))
    app.add_handler(CallbackQueryHandler(lecture_send, pattern=r"^lec_get:"))
    app.add_handler(CallbackQueryHandler(admin_lecture_callback, pattern=r"^adm_lec_"))
    app.add_handler(CallbackQueryHandler(handle_callback))

    if ADMIN_ID:
        app.add_handler(MessageHandler(
            filters.FORWARDED & filters.User(ADMIN_ID), handle_forwarded))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    return app


def main() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise SystemExit("TELEGRAM_BOT_TOKEN غير موجود في متغيرات البيئة.")
    providers = available_providers()
    if not providers:
        raise SystemExit(
            "لا يوجد أي مفتاح مزوّد. اضبط GROQ_API_KEY أو GEMINI_API_KEY "
            "(أو كليهما) في متغيرات البيئة.")
    if not SUBJECTS:
        raise SystemExit("قاموس SUBJECTS فارغ في config.py.")

    os.makedirs(DATA_DIR, exist_ok=True)

    threading.Thread(target=run_http_server, daemon=True).start()
    logger.info("health server on port %s", PORT)
    logger.info("ai providers: %s | max images per request: %s",
                ", ".join(providers), MAX_IMAGES)
    logger.info("subjects loaded: %s | %s",
                len(SUBJECTS), ", ".join(f"{c}={s['lang']}" for c, s in SUBJECTS.items()))
    logger.info("lectures uploaded: %s/%s across %s subjects",
                lectures_total(), lectures_total_slots(),
                len(LECTURE_SUBJECTS))
    logger.info("PDF: fpdf2=%s font=%s bold=%s mono=%s bidi_fallback=%s",
                FPDF_OK, os.path.exists(AR_TTF), os.path.exists(AR_TTF_BOLD),
                os.path.exists(MONO_TTF), BIDI_OK)
    logger.info("stats file: %s | backup chat: %s", STATS_FILE, BACKUP_CHAT_ID or "none")
    logger.info("academic year: %s | youtube key: %s | queries per request: %s",
                academic_year(), bool(YOUTUBE_API_KEY), YT_MAX_QUERIES)
    logger.info("drop pending updates: %s", DROP_PENDING_UPDATES)
    logger.info("disabled services: %s", {k: sorted(v) for k, v in DISABLED_SERVICES.items()})

    failures = 0
    while True:
        app = build_app()
        try:
            logger.info("bot started")
            app.run_polling(
                drop_pending_updates=DROP_PENDING_UPDATES,
                allowed_updates=Update.ALL_TYPES,
                close_loop=False,
            )
            logger.info("polling stopped normally, exiting")
            return
        except InvalidToken:
            raise SystemExit("TELEGRAM_BOT_TOKEN غير صالح، صحّحه في متغيرات البيئة.")
        except Conflict:
            failures += 1
            logger.warning(
                "conflict from telegram (attempt %s/%s): another instance is running, "
                "retrying in %ss", failures, POLL_MAX_FAILURES, POLL_RETRY_CONFLICT)
            time.sleep(POLL_RETRY_CONFLICT)
        except Exception:
            failures += 1
            logger.exception(
                "polling crashed (attempt %s/%s), retrying in %ss",
                failures, POLL_MAX_FAILURES, POLL_RETRY_OTHER)
            time.sleep(POLL_RETRY_OTHER)

        if failures >= POLL_MAX_FAILURES:
            raise SystemExit(
                f"فشل الاستقصاء {failures} مرات متتالية، الخروج ليعيد المشرف التشغيل.")


if __name__ == "__main__":
    main()
