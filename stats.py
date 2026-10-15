# -*- coding: utf-8 -*-
"""الإحصائيات الدائمة والنسخ الاحتياطي إلى قناة خاصة والاستعادة منها."""

import io
import os
import json
import asyncio
from datetime import datetime

from telegram.ext import ContextTypes

from config import (
    logger, STATS_FILE, BAGHDAD_TZ, KEEP_DAYS, BACKUP_CHAT_ID,
    SUBJECTS, SERVICE_LABELS, PROVIDER_LABELS,
)

# ══════════════════════════════ الإحصائيات ══════════════════════════════

_stats_lock = asyncio.Lock()
_stats_dirty = False


def mark_clean() -> None:
    """تصفير مؤشّر التغيير من وحدة أخرى بعد نسخة احتياطية يدوية ناجحة."""
    global _stats_dirty
    _stats_dirty = False


def _today_str() -> str:
    return datetime.now(BAGHDAD_TZ).strftime("%Y-%m-%d")


def _now_str() -> str:
    return datetime.now(BAGHDAD_TZ).isoformat(timespec="seconds")


def _empty_stats() -> dict:
    return {
        "schema": 3,
        "users": {},            # "معرّف" -> {first_seen, last_seen, name, username, requests}
        "requests_count": 0,
        "services_usage": {},
        "subjects_usage": {},
        "providers_usage": {},  # "groq"/"gemini" -> عدد الطلبات التي خدمها
        "files": {"word": 0, "pdf": 0},
        "daily": {},            # "YYYY-MM-DD" -> {requests, users[]}
        "saved_at": None,
    }


def _migrate(data: dict) -> dict:
    """يقبل الصيغة القديمة (قائمة معرّفات) والجديدة ويعيد بنية موحّدة."""
    base = _empty_stats()
    if not isinstance(data, dict):
        return base

    users = data.get("users")
    if isinstance(users, list):
        for uid in users:
            base["users"][str(uid)] = {
                "first_seen": "", "last_seen": "", "name": "",
                "username": "", "requests": 0,
            }
    elif isinstance(users, dict):
        for uid, rec in users.items():
            if isinstance(rec, dict):
                base["users"][str(uid)] = {
                    "first_seen": rec.get("first_seen", "") or "",
                    "last_seen": rec.get("last_seen", "") or "",
                    "name": rec.get("name", "") or "",
                    "username": rec.get("username", "") or "",
                    "requests": int(rec.get("requests", 0) or 0),
                }

    try:
        base["requests_count"] = int(data.get("requests_count", 0) or 0)
    except Exception:
        pass

    for key in ("services_usage", "subjects_usage", "providers_usage", "files"):
        src = data.get(key)
        if isinstance(src, dict):
            for k, v in src.items():
                try:
                    base[key][str(k)] = int(v or 0)
                except Exception:
                    continue

    daily = data.get("daily")
    if isinstance(daily, dict):
        for day, info in daily.items():
            if isinstance(info, dict):
                base["daily"][str(day)] = {
                    "requests": int(info.get("requests", 0) or 0),
                    "users": [str(u) for u in info.get("users", []) or []],
                }

    base["saved_at"] = data.get("saved_at")
    return base


def _prune_daily(data: dict) -> None:
    if len(data["daily"]) > KEEP_DAYS:
        for day in sorted(data["daily"])[:-KEEP_DAYS]:
            data["daily"].pop(day, None)


def _read_stats_file() -> dict:
    try:
        with open(STATS_FILE, "r", encoding="utf-8") as f:
            return _migrate(json.load(f))
    except FileNotFoundError:
        return _empty_stats()
    except Exception as e:
        logger.warning("stats read failed: %s", e)
        return _empty_stats()


def _write_stats_file(data: dict) -> None:
    try:
        os.makedirs(os.path.dirname(STATS_FILE) or ".", exist_ok=True)
        data["saved_at"] = _now_str()
        tmp = STATS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, STATS_FILE)
    except Exception as e:
        logger.warning("stats write failed: %s", e)


def _merge_stats(base: dict, incoming: dict) -> dict:
    """دمج آمن: يأخذ الأكبر لكل عدّاد واتحاد المستخدمين، فلا تتضخّم الأرقام بالتكرار."""
    inc = _migrate(incoming)

    for uid, rec in inc["users"].items():
        cur = base["users"].get(uid)
        if not cur:
            base["users"][uid] = rec
            continue
        cur["requests"] = max(cur.get("requests", 0), rec["requests"])
        if rec["first_seen"] and (not cur.get("first_seen")
                                  or rec["first_seen"] < cur["first_seen"]):
            cur["first_seen"] = rec["first_seen"]
        if rec["last_seen"] > cur.get("last_seen", ""):
            cur["last_seen"] = rec["last_seen"]
            cur["name"] = rec["name"] or cur.get("name", "")
            cur["username"] = rec["username"] or cur.get("username", "")

    for key in ("services_usage", "subjects_usage", "providers_usage", "files"):
        for k, v in inc[key].items():
            base[key][k] = max(base[key].get(k, 0), v)

    for day, info in inc["daily"].items():
        cur = base["daily"].setdefault(day, {"requests": 0, "users": []})
        cur["requests"] = max(cur["requests"], info["requests"])
        cur["users"] = sorted(set(cur["users"]) | set(info["users"]))

    base["requests_count"] = max(base["requests_count"], inc["requests_count"])
    _prune_daily(base)
    return base


def _detect_provider() -> str:
    """اسم المزوّد الذي خدم آخر طلب؛ استيراد كسول تفادياً لأي تعارض استيراد."""
    try:
        from ai_client import get_last_provider
        return get_last_provider() or ""
    except Exception:
        return ""


async def record_activity(user_id: int, service: str | None = None,
                          name: str = "", username: str = "",
                          subject_code: str | None = None,
                          provider: str | None = None) -> None:
    global _stats_dirty
    uid = str(user_id)
    async with _stats_lock:
        data = await asyncio.to_thread(_read_stats_file)
        rec = data["users"].get(uid)
        if rec is None:
            rec = {"first_seen": _now_str(), "last_seen": "", "name": "",
                   "username": "", "requests": 0}
            data["users"][uid] = rec
        rec["last_seen"] = _now_str()
        if name:
            rec["name"] = str(name)[:64]
        if username:
            rec["username"] = str(username)[:32]

        day = data["daily"].setdefault(_today_str(), {"requests": 0, "users": []})
        if uid not in day["users"]:
            day["users"].append(uid)

        if service:
            rec["requests"] = rec.get("requests", 0) + 1
            data["requests_count"] += 1
            data["services_usage"][service] = data["services_usage"].get(service, 0) + 1
            day["requests"] += 1
            if subject_code:
                data["subjects_usage"][subject_code] = \
                    data["subjects_usage"].get(subject_code, 0) + 1
            prov = provider or _detect_provider()
            if prov:
                data["providers_usage"][prov] = \
                    data["providers_usage"].get(prov, 0) + 1

        _prune_daily(data)
        await asyncio.to_thread(_write_stats_file, data)
        _stats_dirty = True


async def record_file(kind: str) -> None:
    """kind = 'word' أو 'pdf'."""
    global _stats_dirty
    async with _stats_lock:
        data = await asyncio.to_thread(_read_stats_file)
        data["files"][kind] = data["files"].get(kind, 0) + 1
        await asyncio.to_thread(_write_stats_file, data)
        _stats_dirty = True


async def remove_users(ids: list[int]) -> None:
    if not ids:
        return
    async with _stats_lock:
        data = await asyncio.to_thread(_read_stats_file)
        for uid in ids:
            data["users"].pop(str(uid), None)
        await asyncio.to_thread(_write_stats_file, data)


async def get_stats() -> dict:
    async with _stats_lock:
        return await asyncio.to_thread(_read_stats_file)


def build_stats_text(data: dict) -> str:
    users = data["users"]
    daily = data["daily"]
    today = _today_str()
    today_d = daily.get(today, {"requests": 0, "users": []})

    last7 = sorted(daily)[-7:]
    active7 = set()
    for d in last7:
        active7.update(daily[d].get("users", []))
    new_today = sum(1 for r in users.values()
                    if str(r.get("first_seen", ""))[:10] == today)

    lines = [
        "📊 إحصائيات المنصة الأكاديمية", "",
        f"الطلبة الكلي: {len(users)}",
        f"جديد اليوم: {new_today}",
        f"نشط اليوم: {len(today_d.get('users', []))}",
        f"نشط آخر سبعة أيام: {len(active7)}",
        f"مجموع العمليات: {data['requests_count']}",
        f"عمليات اليوم: {today_d.get('requests', 0)}", "",
        "الخدمات:",
    ]
    svc = sorted(data["services_usage"].items(), key=lambda x: -x[1])
    lines += [f"• {SERVICE_LABELS.get(k, k)}: {v}" for k, v in svc] or ["• لا يوجد"]

    subs = sorted(data["subjects_usage"].items(), key=lambda x: -x[1])[:8]
    if subs:
        lines += ["", "أكثر المواد طلباً:"]
        for code, count in subs:
            lines.append(f"• {SUBJECTS.get(code, {}).get('name', code)}: {count}")

    provs = sorted(data.get("providers_usage", {}).items(), key=lambda x: -x[1])
    if provs:
        total_p = sum(v for _, v in provs) or 1
        lines += ["", "مزوّدو الذكاء الاصطناعي:"]
        for code, count in provs:
            pct = round(count * 100 / total_p)
            lines.append(f"• {PROVIDER_LABELS.get(code, code)}: {count} ({pct}%)")

    f = data["files"]
    lines += ["", f"الملفات المحمّلة: Word {f.get('word', 0)} — PDF {f.get('pdf', 0)}"]

    if last7:
        lines += ["", "آخر الأيام:"]
        for d in reversed(last7):
            info = daily[d]
            lines.append(f"• {d}: {info.get('requests', 0)} عملية / "
                         f"{len(info.get('users', []))} طالب")

    lines += ["", f"آخر حفظ: {data.get('saved_at') or 'غير معروف'}"]
    return "\n".join(lines)


# ───── النسخ الاحتياطي إلى القناة والاستعادة منها ─────

async def do_backup(bot, reason: str = "دوري") -> bool:
    if not BACKUP_CHAT_ID:
        return False
    try:
        data = await get_stats()
        blob = json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")
        caption = (f"نسخة احتياطية ({reason})\n{_now_str()}\n"
                   f"الطلبة: {len(data['users'])} | العمليات: {data['requests_count']}")
        msg = await bot.send_document(
            chat_id=BACKUP_CHAT_ID, document=io.BytesIO(blob),
            filename="bot_stats.json", caption=caption, disable_notification=True)
        try:
            await bot.unpin_all_chat_messages(chat_id=BACKUP_CHAT_ID)
        except Exception:
            pass
        try:
            await bot.pin_chat_message(chat_id=BACKUP_CHAT_ID,
                                       message_id=msg.message_id,
                                       disable_notification=True)
        except Exception as e:
            logger.warning("pin failed: %s", e)
        logger.info("backup sent (%s)", reason)
        return True
    except Exception as e:
        logger.warning("backup failed: %s", e)
        return False


async def restore_stats(bot) -> dict | None:
    """يقرأ الرسالة المثبّتة في قناة النسخ ويدمجها مع الملف المحلي."""
    if not BACKUP_CHAT_ID:
        return None
    try:
        chat = await bot.get_chat(BACKUP_CHAT_ID)
        pinned = getattr(chat, "pinned_message", None)
        if pinned is None or pinned.document is None:
            logger.info("no pinned backup found")
            return None
        f = await bot.get_file(pinned.document.file_id)
        raw = bytes(await f.download_as_bytearray())
        incoming = json.loads(raw.decode("utf-8"))
        async with _stats_lock:
            data = await asyncio.to_thread(_read_stats_file)
            data = _merge_stats(data, incoming)
            await asyncio.to_thread(_write_stats_file, data)
        logger.info("stats restored: %s users", len(data["users"]))
        return data
    except Exception as e:
        logger.warning("restore failed: %s", e)
        return None


async def backup_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    global _stats_dirty
    if not _stats_dirty:
        return
    if await do_backup(context.bot, "دوري"):
        _stats_dirty = False
