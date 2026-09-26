# -*- coding: utf-8 -*-
"""
المنصة الأكاديمية الذكية - بوت تلغرام
المرحلة الثانية - علوم الحاسوب

الخدمات: تقرير (Word / PDF) - شرح موضوع أو صفحة ملزمة - حل مسائل
         تدقيق أكواد - بنك أسئلة
لغة الناتج تتبع لغة دراسة المادة، ما عدا الشرح فيكون بالعربية افتراضياً
مع الحفاظ على المصطلحات الإنكليزية. الصور مدعومة في المسائل والأكواد والشرح.

إحصائيات دائمة: عدد المستخدمين ومرات استخدام كل خدمة، مع نسخ احتياطي
تلقائي إلى قناة خاصة واستعادة تلقائية بعد كل إعادة نشر.

هذا الملف هو نقطة التشغيل: الواجهة والأوامر والأزرار ومعالجة الرسائل.
"""

import io
import os
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

# تلغرام لا يسمح إلا بنسخة واحدة تستدعي getUpdates لكل توكن. أثناء إعادة
# النشر على Render تتعايش النسخة القديمة والجديدة لحظات، فيقطع تلغرام
# إحداهما بخطأ Conflict. بلا امتصاص لهذا الخطأ تموت العملية ولا يعيدها
# أحد، فتبقى الخدمة ترد 503 حتى تدخل يدوياً.
POLL_RETRY_CONFLICT = 25      # ثانية انتظار قبل إعادة المحاولة بعد Conflict
POLL_RETRY_OTHER = 15         # ثانية انتظار بعد أي انقطاع آخر
POLL_MAX_FAILURES = 5         # بعد هذا العدد نُخرج العملية ليعيدها مشرف العمليات

# حذف الرسائل المتراكمة أثناء التوقف. الافتراضي الآن هو عدم الحذف حتى لا
# تُفقد طلبات الطلبة خلال انقطاع طويل. لتعطيل ذلك اضبط DROP_PENDING_UPDATES=1
DROP_PENDING_UPDATES = os.getenv("DROP_PENDING_UPDATES", "0").strip() == "1"


# ===================== الواجهة =====================

# رسالة المادة المحذوفة: user_data محفوظ بـ PicklePersistence، فتبقى عند
# بعض الطلبة حالة تشير إلى مادة أُزيلت من SUBJECTS، وكذلك تبقى قوائم
# قديمة معروضة في محادثاتهم. بلا هذا الحارس يرفع SUBJECTS[code] خطأ
# KeyError خارج كتلة try فلا يصل الطالب أي ردّ.
STALE_SUBJECT_MSG = (
    "هذه المادة لم تعد متاحة في البوت.\n"
    "اختر الخدمة والمادة من جديد:"
)


def get_subject(state: dict | None) -> dict | None:
    """يعيد المادة من الحالة، أو None إن كانت الحالة ناقصة أو المادة محذوفة."""
    if not state or "service" not in state:
        return None
    return SUBJECTS.get(state.get("code"))


def main_menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🎓 شرح موضوع أو صفحة ملزمة", callback_data="svc:exp")],
        [InlineKeyboardButton("📝 إنشاء تقرير (Word / PDF)", callback_data="svc:rep")],
        [InlineKeyboardButton("🧮 حل مسألة (نص أو صورة)", callback_data="svc:math")],
        [InlineKeyboardButton("💻 تدقيق كود (نص أو صورة)", callback_data="svc:code")],
        [InlineKeyboardButton("📚 بنك أسئلة", callback_data="svc:quiz")],
        [InlineKeyboardButton("⚙️ الإعدادات", callback_data="settings")],
    ])


def subjects_kb(service: str) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    for term in (1, 2):
        items = [(c, s) for c, s in SUBJECTS.items()
                 if s["term"] == term and (s["is_code"] if service == "code" else True)]
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
    return any(s["is_code"] if service == "code" else True for s in SUBJECTS.values())


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


# القائمة الرئيسية موجزة: شرح كل خدمة انتقل إلى داخلها، وسطر الأوامر حُذف
WELCOME = (
    "🎓 *المنصة الأكاديمية الذكية*\n"
    "المرحلة الثانية — علوم الحاسوب\n\n"
    "اختر الخدمة، ثم المادة، ثم أرسل طلبك نصاً أو صورة."
)

# وصف كل خدمة يُعرض بعد الضغط على زرها فقط
SERVICE_NOTES = {
    "exp": "تقبل هذه الخدمة اسم الموضوع أو صورة صفحة من الملزمة، "
           "وتقترح لك مقاطع يوتيوب للموضوع.",
    "rep": "تقبل هذه الخدمة عنواناً نصياً، وتُخرج التقرير بصيغة Word أو PDF.",
    "math": "تقبل هذه الخدمة نص المسألة أو صورتها، وتعرض الحل خطوة بخطوة.",
    "code": "تقبل هذه الخدمة الكود نصاً أو صورة، وتعرض الأخطاء والكود المصحّح.\n"
            "(في هذه الخدمة تظهر المواد البرمجية فقط)",
    "quiz": "تقبل هذه الخدمة اسم الموضوع أو الفصل، وتُخرج أسئلة مع مفتاح الإجابات.",
}

# تنبيه يُذيَّل به كل ناتج: النموذج قد يخطئ ولا يعرف محتوى المحاضرة الأصلية
DISCLAIMER = (
    "\n\n⚠️ هذا ناتج آلي مساعد للمراجعة ولا يُغني عن المحاضرة الأصلية. "
    "راجع الصفحة، وتحقّق بنفسك من أي أرقام أو حسابات."
)

# حرارة التوليد لكل خدمة. التقرير والشرح يحتاجان حرارة أعلى لأن 0.3 كانت
# تدفع النموذج إلى الانغلاق على عبارة وتكرارها حتى نفاد الرموز في القوائم
# العربية المتشابهة. المسائل والكود تبقى منخفضة، فالدقة فيها أهم من التنوّع.
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
    await update.message.reply_text(
        "🛠 لوحة تحكم المنصة الأكاديمية\n\n"
        f"الطلبة المسجلين: {len(st['users'])}\n"
        f"مجموع العمليات: {st['requests_count']}\n"
        f"طلبات قيد التنفيذ الآن: {len(BUSY_USERS)}\n\n"
        f"المواد المتاحة: {len(SUBJECTS)}\n"
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
        "الأوامر: /stats /users /backup /restore /getid /broadcast"
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
            f"العمليات: {data['requests_count']}")
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
            # ضغط على زر من قائمة قديمة عالقة في محادثته لمادة أُزيلت
            context.user_data.pop("state", None)
            await q.edit_message_text(STALE_SUBJECT_MSG, reply_markup=main_menu_kb())
            return
        if "service" not in state:
            await q.edit_message_text("انتهت صلاحية الاختيار. ابدأ من جديد.",
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

    state = context.user_data.get("state")
    if not state or "service" not in state or "code" not in state:
        await msg.reply_text("اختر الخدمة والمادة أولاً:", reply_markup=main_menu_kb())
        return

    # المادة قد تكون أُزيلت من config بعد أن اختارها الطالب (الحالة محفوظة)
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

    # رسالة الانتظار تُنشأ داخل try: لو فشل إرسالها (حظر أو قطع شبكة) كان
    # المستخدم يبقى في BUSY_USERS إلى الأبد فلا يُقبل منه أي طلب لاحق.
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
            # الطالب كتب الموضوع بنفسه، فلا يُستخرج من المتن
            # (كان يُعاد «نظرة عامة» فتفسد عبارة بحث يوتيوب واسم الملف)
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
    # الحالة قد تكون قديمة أو المادة محذوفة، خصوصاً في مسار الألبوم المؤجَّل
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
        # نماذج الرؤية على Groq من فئة المعاينة وقد تُسحب بإشعار قصير
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
    if (update.effective_user.id == ADMIN_ID and doc
            and (doc.file_name or "").lower().endswith(".json")):
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
                f"العمليات: {data['requests_count']}")
        except Exception as e:
            await msg.reply_text(f"فشل قراءة الملف: {e}")
        return
    await msg.reply_text("الملفات غير مدعومة. أرسل النص مكتوباً أو صورة للصفحة أو المسألة.")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Conflict ليس عطلاً في الكود بل نسخة أخرى تسحب التحديثات، فيُسجَّل
    # سطراً واحداً واضحاً بلا تتبّع طويل يُغرق سجل Render.
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
    if not data["users"]:
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
                f"السنة الدراسية: {academic_year()}\n"
                f"مزوّدو الذكاء: {prov_txt}\n"
                f"يوتيوب: {'مفتاح مضبوط' if YOUTUBE_API_KEY else 'روابط بحث فقط'}")
        except Exception:
            pass


async def post_stop(app) -> None:
    await do_backup(app.bot, "إيقاف")


# ===================== التشغيل =====================

def build_app() -> Application:
    """يبني تطبيقاً جديداً بكل معالجاته.

    مستقل عن main لأن إعادة تشغيل الاستقصاء بعد Conflict تحتاج كائن
    Application جديداً: الكائن الذي مرّ عليه shutdown لا يُعاد تشغيله
    بأمان في python-telegram-bot.
    """
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
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("restore", restore_command))
    app.add_handler(CommandHandler("getid", getid_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))
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

    # الخادم يُشغَّل مرة واحدة فقط وخارج حلقة إعادة المحاولة: لو أُعيد
    # تشغيله لكان bind على نفس المنفذ يفشل بـ Address already in use،
    # وتوقّف Render عن رؤية المنفذ فيُعلن فشل النشر.
    threading.Thread(target=run_http_server, daemon=True).start()
    logger.info("health server on port %s", PORT)
    logger.info("ai providers: %s | max images per request: %s",
                ", ".join(providers), MAX_IMAGES)
    logger.info("subjects loaded: %s | %s",
                len(SUBJECTS), ", ".join(f"{c}={s['lang']}" for c, s in SUBJECTS.items()))
    logger.info("PDF: fpdf2=%s font=%s bold=%s mono=%s bidi_fallback=%s",
                FPDF_OK, os.path.exists(AR_TTF), os.path.exists(AR_TTF_BOLD),
                os.path.exists(MONO_TTF), BIDI_OK)
    logger.info("stats file: %s | backup chat: %s", STATS_FILE, BACKUP_CHAT_ID or "none")
    logger.info("academic year: %s | youtube key: %s | queries per request: %s",
                academic_year(), bool(YOUTUBE_API_KEY), YT_MAX_QUERIES)
    logger.info("drop pending updates: %s", DROP_PENDING_UPDATES)

    failures = 0
    while True:
        app = build_app()
        try:
            logger.info("bot started")
            # close_loop=False ضروري: بدونه يُغلق حلقة asyncio عند الخروج
            # فتفشل أي محاولة تالية بـ Event loop is closed.
            app.run_polling(
                drop_pending_updates=DROP_PENDING_UPDATES,
                allowed_updates=Update.ALL_TYPES,
                close_loop=False,
            )
            logger.info("polling stopped normally, exiting")
            return
        except InvalidToken:
            # خطأ إعداد لا يُصلحه الانتظار، فالخروج أنفع من حلقة عبثية
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
            # الخروج بحالة غير صفرية متعمَّد: حلقة bash في Start Command
            # ستعيد العملية من الصفر، وهذا أنظف من تكديس حالة داخلية معطوبة.
            raise SystemExit(
                f"فشل الاستقصاء {failures} مرات متتالية، الخروج ليعيد المشرف التشغيل.")


if __name__ == "__main__":
    main()
