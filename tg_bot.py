#!/usr/bin/env python3
"""
Telegram-бот для интерактивного полевого обхода магазинов и сбора аналитики оттока.

Функции:
  ▶️ Начать / Продолжить обход — карточка текущей точки с кнопками WhatsApp, 2ГИС, GMaps
  📋 Весь маршрут — полный список всех точек со статусами и ETA
  📊 Скачать Excel-отчет — готовый xlsx с цветовой разметкой и листом «Аналитика оттока»
  🔄 Сбросить прогресс — очистка статусов обхода
  📤 Отправить xlsx — пересчёт маршрута
  📍 Отправить геолокацию — пересчёт оставшихся точек от текущих координат
  📸 Поддержка текста и фото при записи заметок из магазина
"""
import asyncio
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)
from telegram.constants import ParseMode

# ── Настройки ─────────────────────────────────────────────
TOKEN = os.environ.get("BOT_TOKEN", "9093591208:AAHLuair6rfphdfsQKwM8RttPpF6qkAQW3s")
WORK_DIR = Path(__file__).resolve().parent
SRC_FILE = WORK_DIR / "Маршрут_Астана.xlsx"
OUT_FILE = WORK_DIR / "Маршрут_Астана_готовый.xlsx"
SCRIPT = WORK_DIR / "route_planner.py"
STATE_FILE = WORK_DIR / "visit_state.json"
REPORT_FILE = WORK_DIR / "Маршрут_Астана_отчет.xlsx"
PHOTOS_DIR = WORK_DIR / "photos"
PYTHON = sys.executable

# Среднее время пребывания в одной точке (парковка + разговор), минут
AVG_SERVICE_MINS = 12

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("route-bot")

WAITING_NOTE = 0

MAIN_KB = ReplyKeyboardMarkup(
    [
        ["▶️ Начать / Продолжить обход", "📋 Весь маршрут"],
        ["📊 Скачать Excel-отчет", "🔄 Сбросить прогресс"],
    ],
    resize_keyboard=True,
)

CITY_KB = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("📍 Астана", callback_data="city:Астана"),
            InlineKeyboardButton("📍 Алматы (скоро)", callback_data="city:Алматы"),
        ],
        [
            InlineKeyboardButton("📍 Шымкент (скоро)", callback_data="city:Шымкент"),
            InlineKeyboardButton("📍 Костанай (скоро)", callback_data="city:Костанай"),
        ],
    ]
)

REASON_TAGS = {
    "price": "💸 Цена/Тариф",
    "tech": "⚙️ Тех. проблемы/Баги",
    "competitor": "🔄 Ушли к конкуренту",
    "loyal": "🤝 Лояльны / Работают",
    "freetext": "✍️ Сразу текст",
}

STATUS_EMOJI = {
    "pending": "⏳",
    "visited": "✅",
    "skipped": "⏭",
    "closed": "🔒",
}


# ══════════════════════════════════════════════════════════
# ── Утилиты и работа с файлами ───────────────────────────
# ══════════════════════════════════════════════════════════

def _esc(text) -> str:
    """Безопасное экранирование спецсимволов для Telegram Markdown v1."""
    if text is None:
        return ""
    text = str(text)
    for ch in ("_", "*", "`", "["):
        text = text.replace(ch, "\\" + ch)
    return text


def _get_out_file() -> Path:
    """Находит актуальный файл готового маршрута, учитывая возможные дубликаты при блокировке."""
    if OUT_FILE.exists():
        alts = sorted(WORK_DIR.glob("Маршрут_Астана_готовый_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if alts and alts[0].stat().st_mtime > OUT_FILE.stat().st_mtime:
            return alts[0]
        return OUT_FILE
    alts = sorted(WORK_DIR.glob("Маршрут_Астана_готовый_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if alts:
        return alts[0]
    return OUT_FILE


# ══════════════════════════════════════════════════════════
# ── Работа с состоянием обхода (visit_state.json) ────────
# ══════════════════════════════════════════════════════════

def _point_uid(p: dict) -> str:
    """Уникальный ключ магазина (Имя + Адрес), чтобы заметки не терялись при пересчёте маршрута."""
    return f"{str(p.get('name', '')).strip()}|{str(p.get('addr', '')).strip()}"


def _load_state() -> dict:
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            log.error(f"Ошибка чтения state: {e}")
    return {}


def _save_state(state: dict):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.error(f"Ошибка записи state: {e}")


def _init_state_from_route():
    """
    Создаёт или обновляет visit_state.json из Маршрут_Астана_готовый.xlsx.
    Сопоставляет точки по уникальному ключу (name|addr), чтобы не терять заметки.
    """
    points = _read_route_points()
    if not points:
        return {}

    old_state = _load_state()
    old_by_uid = {_point_uid(p): p for p in old_state.get("points", [])}

    new_points = []
    for p in points:
        uid = _point_uid(p)
        old = old_by_uid.get(uid, {})
        new_points.append(
            {
                "order": p["order"],
                "name": p["name"],
                "addr": p["addr"],
                "owner": p["owner"],
                "contact": p["contact"],
                "phone": p["phone"],
                "mins": p["mins"],
                "total": p["total"],
                "src": p["src"],
                "gmaps": p["gmaps"],
                "dgis": p["dgis"],
                "lat": p.get("lat"),
                "lon": p.get("lon"),
                "status": old.get("status", "pending"),
                "visit_time": old.get("visit_time", ""),
                "reason_tag": old.get("reason_tag", ""),
                "notes": old.get("notes", ""),
                "photo": old.get("photo", ""),
            }
        )

    state = {
        "city": old_state.get("city", "Астана"),
        "points": new_points,
    }
    _save_state(state)
    return state


def _get_point_by_order(state: dict, order: int | None) -> dict | None:
    if order is None:
        return None
    for p in state.get("points", []):
        if p.get("order") == order:
            return p
    return None


def _next_to_visit(state: dict) -> tuple[dict | None, bool]:
    """
    Возвращает (point, is_from_skipped).
    Сначала ищет 'pending', а если они закончились — возвращает первую 'skipped'.
    """
    pts = state.get("points", [])
    for p in pts:
        if p.get("status") == "pending":
            return p, False
    for p in pts:
        if p.get("status") == "skipped":
            return p, True
    return None, False


def _progress_summary(state: dict) -> str:
    points = state.get("points", [])
    total = len(points)
    visited = sum(1 for p in points if p.get("status") == "visited")
    skipped = sum(1 for p in points if p.get("status") == "skipped")
    closed = sum(1 for p in points if p.get("status") == "closed")
    pending = sum(1 for p in points if p.get("status") == "pending")

    rem_pts = [p for p in points if p.get("status") in ("pending", "skipped")]
    rem_drive = sum(p.get("mins", 0) for p in rem_pts)
    rem_total_mins = rem_drive + len(rem_pts) * AVG_SERVICE_MINS
    eta_str = ""
    if rem_pts:
        eta_dt = datetime.now() + timedelta(minutes=rem_total_mins)
        eta_str = (
            f"\n🏁 *Прогноз завершения ({len(rem_pts)} точ.):* "
            f"≈ {eta_dt:%H:%M} (~{rem_total_mins} мин с визитами)"
        )

    return (
        f"📊 *Прогресс:* {visited + closed} из {total} закрыто\n"
        f"   ✅ Посещено: {visited} | ⏭ Отложено: {skipped} | "
        f"🔒 Закрыто: {closed} | ⏳ Ожидает: {pending}"
        f"{eta_str}"
    )


# ══════════════════════════════════════════════════════════
# ── Чтение xlsx ──────────────────────────────────────────
# ══════════════════════════════════════════════════════════

def _read_route_points():
    cur_file = _get_out_file()
    if not cur_file.exists():
        return []

    try:
        wb = load_workbook(cur_file)
    except Exception as e:
        log.error(f"Не удалось загрузить workbook {cur_file}: {e}")
        return []

    if "Маршрут" not in wb.sheetnames:
        return []

    ws = wb["Маршрут"]
    points = []
    for row in ws.iter_rows(min_row=2, max_col=11):
        val = row[0].value
        if val is None or not isinstance(val, int):
            continue
        gmaps = row[9].hyperlink.target if row[9].hyperlink else None
        dgis = row[10].hyperlink.target if row[10].hyperlink else None

        lat, lon = None, None
        if gmaps and "query=" in gmaps:
            try:
                coords_str = gmaps.split("query=")[1].split("&")[0]
                lat, lon = [float(x) for x in coords_str.split(",")]
            except (ValueError, IndexError):
                pass

        points.append(
            dict(
                order=val,
                name=str(row[1].value or ""),
                addr=str(row[2].value or ""),
                owner=str(row[3].value or ""),
                contact=str(row[4].value or ""),
                phone=str(row[5].value or ""),
                mins=int(row[6].value or 0),
                total=int(row[7].value or 0),
                src=str(row[8].value or ""),
                gmaps=gmaps or "",
                dgis=dgis or "",
                lat=lat,
                lon=lon,
            )
        )
    return points


# ══════════════════════════════════════════════════════════
# ── Форматирование карточки ──────────────────────────────
# ══════════════════════════════════════════════════════════

def _whatsapp_url(phone_raw: str) -> str | None:
    digits = re.sub(r"\D", "", str(phone_raw))
    if not digits:
        return None
    if digits.startswith("8") and len(digits) == 11:
        digits = "7" + digits[1:]
    if len(digits) >= 10:
        return f"https://wa.me/{digits}"
    return None


def _card_text(p: dict, state: dict) -> str:
    status_em = STATUS_EMOJI.get(p.get("status", "pending"), "")
    lines = [
        f"*{p['order']}.* {_esc(p.get('name', ''))} {status_em}",
        f"📍 {_esc(p.get('addr', ''))}",
    ]

    next_p = _get_point_by_order(state, p["order"] + 1)
    if next_p and next_p.get("mins", 0) == 0:
        lines.append(f"🏢 *В этом же здании/рядом:* {_esc(next_p.get('name', ''))}")

    if p.get("owner"):
        lines.append(f"👤 Владелец: {_esc(p['owner'])}")
    if p.get("contact"):
        lines.append(f"🧑‍💼 Контакт: {_esc(p['contact'])}")
    if p.get("phone"):
        wa = _whatsapp_url(p["phone"])
        if wa:
            lines.append(f"📞 {_esc(p['phone'])} · [💬 WhatsApp]({wa})")
        else:
            lines.append(f"📞 {_esc(p['phone'])}")

    if p["order"] == 1:
        lines.append("🏁 СТАРТ")
    elif p.get("mins", 0) == 0:
        lines.append(f"⏱ 0 мин (в том же месте, всего в пути {p.get('total', 0)} мин)")
    else:
        lines.append(f"⏱ +{p.get('mins', 0)} мин езды (всего в пути {p.get('total', 0)} мин)")

    links = []
    if p.get("gmaps"):
        links.append(f"[Google Maps]({p['gmaps']})")
    if p.get("dgis"):
        links.append(f"[2ГИС]({p['dgis']})")
    if links:
        lines.append("🔗 " + " · ".join(links))

    if p.get("reason_tag"):
        lines.append(f"\n🏷 Категория: *{_esc(p['reason_tag'])}*")
    if p.get("notes"):
        # Без обрамления в курсив, чтобы не ломался Markdown при переносах строк
        lines.append(f"💬 Заметка: {_esc(p['notes'])}")
    if p.get("photo"):
        lines.append(f"📸 Прикреплено фото: {_esc(p['photo'])}")
    if p.get("visit_time"):
        lines.append(f"🕐 Время визита: {_esc(p['visit_time'])}")

    return "\n".join(lines)


def _card_keyboard(p: dict, total: int) -> InlineKeyboardMarkup:
    order = p["order"]
    buttons = []

    if p.get("status") == "visited":
        buttons.append(
            [InlineKeyboardButton("✍️ Дополнить заметку / фото", callback_data=f"visit:{order}")]
        )
    else:
        buttons.append(
            [InlineKeyboardButton("✅ Записать итог / Заметку", callback_data=f"visit:{order}")]
        )

    buttons.append(
        [
            InlineKeyboardButton("⏭ Отложить на потом", callback_data=f"skip:{order}"),
            InlineKeyboardButton("🔒 Магазин закрыт", callback_data=f"closed:{order}"),
        ]
    )

    nav = []
    if order > 1:
        nav.append(InlineKeyboardButton("⬅️ Пред.", callback_data=f"nav:{order - 1}"))
    if order < total:
        nav.append(InlineKeyboardButton("След. ➡️", callback_data=f"nav:{order + 1}"))
    if nav:
        buttons.append(nav)

    return InlineKeyboardMarkup(buttons)


# ══════════════════════════════════════════════════════════
# ── Генерация Excel-отчёта ───────────────────────────────
# ══════════════════════════════════════════════════════════

def export_report_xlsx() -> Path:
    """
    Пересобирает лист «Маршрут» строго по актуальному порядку state['points']
    и добавляет лист «Аналитика оттока».
    """
    state = _load_state()
    all_points = state.get("points", [])

    cur_file = _get_out_file()
    if cur_file.exists():
        try:
            wb = load_workbook(cur_file)
        except Exception:
            wb = Workbook()
    else:
        wb = Workbook()

    if "Маршрут" in wb.sheetnames:
        del wb["Маршрут"]
    ws = wb.create_sheet("Маршрут", 0)

    # Удаляем пустой лист по умолчанию, если он был создан
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    F = "Arial"
    hf = PatternFill("solid", fgColor="1F3A5F")

    headers = [
        "Порядок",
        "Точка",
        "Адрес",
        "Владелец",
        "Контакт",
        "Телефон",
        "Минут от предыдущей",
        "Всего минут",
        "Откуда координаты",
        "Google Maps",
        "2ГИС",
        "Статус",
        "Время визита",
        "Категория / Причина",
        "Комментарий / Что сказали в магазине",
    ]
    ws.append(headers)

    fill_visited = PatternFill("solid", fgColor="C6EFCE")
    fill_skipped = PatternFill("solid", fgColor="FFEB9C")
    fill_closed = PatternFill("solid", fgColor="FFC7CE")
    fill_pending = PatternFill("solid", fgColor="D9E1F2")

    STATUS_RU = {
        "visited": "Посетил",
        "skipped": "Отложено",
        "closed": "Закрыто",
        "pending": "Ожидает",
    }
    STATUS_FILL = {
        "visited": fill_visited,
        "skipped": fill_skipped,
        "closed": fill_closed,
        "pending": fill_pending,
    }

    for p in all_points:
        status = p.get("status", "pending")
        note_full = p.get("notes", "")
        if p.get("photo"):
            note_full = (note_full + f" [Фото: {p['photo']}]").strip()

        ws.append(
            [
                p["order"],
                p.get("name", ""),
                p.get("addr", ""),
                p.get("owner", ""),
                p.get("contact", ""),
                p.get("phone", ""),
                p.get("mins", 0),
                p.get("total", 0),
                p.get("src", ""),
                "Открыть" if p.get("gmaps") else "",
                "Открыть" if p.get("dgis") else "",
                STATUS_RU.get(status, status),
                p.get("visit_time", ""),
                p.get("reason_tag", ""),
                note_full,
            ]
        )
        r = ws.max_row
        if p.get("gmaps"):
            ws.cell(r, 10).hyperlink = p["gmaps"]
            ws.cell(r, 10).font = Font(name=F, color="0563C1", underline="single")
        if p.get("dgis"):
            ws.cell(r, 11).hyperlink = p["dgis"]
            ws.cell(r, 11).font = Font(name=F, color="0563C1", underline="single")

        cell_status = ws.cell(r, 12)
        cell_status.fill = STATUS_FILL.get(status, fill_pending)
        cell_status.alignment = Alignment(horizontal="center", vertical="center")

    for c in ws[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = hf
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in ws.iter_rows(min_row=2, max_col=15):
        for idx_c, c in enumerate(row, 1):
            if idx_c not in (10, 11):
                c.font = Font(name=F, size=11)
            if idx_c in (1, 7, 8, 12, 13):
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                c.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [8, 26, 38, 24, 22, 18, 12, 11, 14, 13, 10, 14, 18, 24, 52]
    for col_letter, w in zip("ABCDEFGHIJKLMNO", widths):
        ws.column_dimensions[col_letter].width = w
    ws.freeze_panes = "C2"

    # ── Лист «Аналитика оттока» ──
    if "Аналитика оттока" in wb.sheetnames:
        del wb["Аналитика оттока"]
    an = wb.create_sheet("Аналитика оттока")

    total = len(all_points)
    visited = [p for p in all_points if p.get("status") == "visited"]
    skipped = [p for p in all_points if p.get("status") == "skipped"]
    closed = [p for p in all_points if p.get("status") == "closed"]
    pending = [p for p in all_points if p.get("status") == "pending"]

    an.append(["Показатель", "Количество"])
    an.append(["Всего точек в маршруте", total])
    an.append(["Посещено (успешный визит)", len(visited)])
    an.append(["Отложено", len(skipped)])
    an.append(["Магазин закрыт", len(closed)])
    an.append(["Ожидает посещения", len(pending)])
    an.append([])
    an.append(["Разбивка по причинам оттока / категориям", "Количество"])

    reason_counts = {}
    for p in all_points:
        tag = p.get("reason_tag", "")
        if tag:
            reason_counts[tag] = reason_counts.get(tag, 0) + 1
    if reason_counts:
        for reason, count in sorted(reason_counts.items(), key=lambda x: -x[1]):
            an.append([reason, count])
    else:
        an.append(["(пока нет отмеченных категорий)", 0])

    for c in an[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = hf
        c.alignment = Alignment(horizontal="center", vertical="center")
    for c in an[8]:
        if c.value:
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = hf
            c.alignment = Alignment(horizontal="center", vertical="center")

    an.column_dimensions["A"].width = 40
    an.column_dimensions["B"].width = 16

    try:
        wb.save(str(REPORT_FILE))
    except PermissionError:
        alt = WORK_DIR / f"Маршрут_Астана_отчет_{datetime.now():%H%M%S}.xlsx"
        wb.save(str(alt))
        return alt
    return REPORT_FILE


# ══════════════════════════════════════════════════════════
# ── Команды и хэндлеры бота ──────────────────────────────
# ══════════════════════════════════════════════════════════

async def cmd_start(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "👋 *Привет! Это бот-ассистент для умного объезда торговых точек и сбора обратной связи.*\n\n"
        "Я строю самый быстрый маршрут по дорогам, веду по точкам и сохраняю ваши заметки "
        "прямо в готовый Excel-отчет.\n\n"
        "📌 *Мини-инструкция, как работать:*\n"
        "1️⃣ *Выберите город* на кнопках ниже.\n"
        "2️⃣ *Отправьте .xlsx файл* со списком магазинов (первая строка — ваш СТАРТ).\n"
        "3️⃣ *Нажмите «▶️ Начать / Продолжить обход»* — откроется карточка с кнопками *WhatsApp*, *2ГИС* и *Google Maps*.\n"
        "4️⃣ *Фиксируйте итоги* — отмечайте статус визита и пишите текстом (или присылайте фото), почему клиент перешел на другую систему.\n"
        "5️⃣ *Геолокация* — если отклонились от маршрута, отправьте свою 📍 геопозицию, и я пересчитаю оставшиеся точки.\n"
        "6️⃣ *Скачайте отчет* — в любой момент заберите готовый Excel со всеми комментариями и аналитикой.\n\n"
        "🏙 *Выберите ваш город:* 👇"
    )
    await update.message.reply_text(
        welcome_text,
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=MAIN_KB,
    )
    await update.message.reply_text(
        "Доступные города:",
        reply_markup=CITY_KB,
    )


async def cb_city_select(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, city = query.data.split(":", 1)

    if city != "Астана":
        await query.message.reply_text(
            f"🚧 Город *{_esc(city)}* скоро будет подключен! Пока активен модуль по *г. Астана*.",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    state = _load_state()
    state["city"] = city
    _save_state(state)
    await query.message.reply_text(
        f"✅ Выбран город: *{_esc(city)}*.\n"
        "Отправьте файл `.xlsx` для расчёта или нажмите *▶️ Начать / Продолжить обход*.",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=MAIN_KB,
    )


# ── Показать весь маршрут ────────────────────────────────

async def cmd_show_all(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    state = _load_state()
    points = state.get("points", [])
    if not points:
        await update.message.reply_text(
            "📤 Сначала отправьте xlsx-файл с адресами.\n"
            "Я рассчитаю маршрут и покажу его здесь.",
            reply_markup=MAIN_KB,
        )
        return

    msgs, buf = [], ["🗺 *Маршрут по Астане*\n"]
    for p in points:
        em = STATUS_EMOJI.get(p.get("status", "pending"), "")
        links = []
        if p.get("gmaps"):
            links.append(f'[GMaps]({p["gmaps"]})')
        if p.get("dgis"):
            links.append(f'[2ГИС]({p["dgis"]})')
        lnk = " · ".join(links)

        if p["order"] == 1:
            time_s = "🏁 СТАРТ"
        else:
            time_s = f"⏱ +{p.get('mins', 0)} мин  (всего {p.get('total', 0)})"

        note_preview = f"\n   💬 {_esc(p['notes'])}" if p.get("notes") else ""

        block = (
            f"*{p['order']}.* {_esc(p.get('name', ''))} {em}\n"
            f"   📍 {_esc(p.get('addr', ''))}\n"
            f"   {time_s}  {lnk}{note_preview}\n"
        )

        joined = "\n".join(buf) + "\n" + block
        if len(joined) > 3800:
            msgs.append("\n".join(buf))
            buf = ["🗺 *Продолжение маршрута*\n"]
        buf.append(block)

    total_time = points[-1].get("total", 0) if points else 0
    buf.append(f"\n⏱ *Чистое время за рулём: ≈ {total_time} мин*")
    buf.append("\n" + _progress_summary(state))
    msgs.append("\n".join(buf))

    for msg in msgs:
        await update.message.reply_text(
            msg,
            parse_mode=ParseMode.MARKDOWN,
            disable_web_page_preview=True,
        )


# ── Начать/Продолжить обход ──────────────────────────────

async def _send_point_card(target, state: dict, order: int | None = None, edit: bool = False):
    total = len(state.get("points", []))
    is_from_skipped = False

    if order is not None:
        p = _get_point_by_order(state, order)
    else:
        p, is_from_skipped = _next_to_visit(state)

    msg_obj = getattr(target, "message", None) or target

    if p is None:
        text = "🎉 *Все точки маршрута пройдены!*\n\n" + _progress_summary(state)
        await msg_obj.reply_text(text, parse_mode=ParseMode.MARKDOWN, reply_markup=MAIN_KB)
        return None

    prefix = "🔄 *Возвращаемся к отложенной точке:*\n\n" if is_from_skipped else ""
    text = prefix + _card_text(p, state)
    kb = _card_keyboard(p, total)

    if edit and hasattr(target, "edit_message_text"):
        try:
            await target.edit_message_text(
                text,
                parse_mode=ParseMode.MARKDOWN,
                disable_web_page_preview=True,
                reply_markup=kb,
            )
            return p
        except Exception:
            pass

    await msg_obj.reply_text(
        text,
        parse_mode=ParseMode.MARKDOWN,
        disable_web_page_preview=True,
        reply_markup=kb,
    )
    return p


async def cmd_start_walk(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    state = _load_state()
    if not state.get("points"):
        cur_file = _get_out_file()
        if cur_file.exists():
            state = _init_state_from_route()
        else:
            await update.message.reply_text(
                "📤 Сначала отправьте xlsx-файл с адресами.",
                reply_markup=MAIN_KB,
            )
            return

    result = await _send_point_card(update.message, state)
    if result is None:
        await _auto_send_report(update.message)


async def _auto_send_report(message_or_target):
    try:
        report_path = export_report_xlsx()
        msg_obj = getattr(message_or_target, "message", None) or message_or_target
        with open(report_path, "rb") as f:
            await msg_obj.reply_document(
                document=InputFile(f, filename="Маршрут_Астана_отчет.xlsx"),
                caption="📊 Итоговый отчёт обхода со всеми заметками и аналитикой оттока",
            )
    except Exception as e:
        log.error(f"Ошибка автоотчёта: {e}")


# ── Callback-кнопки навигации, skip, closed ──────────────

async def cb_navigation(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, order_str = query.data.split(":", 1)
    state = _load_state()
    await _send_point_card(query, state, int(order_str), edit=True)


async def cb_skip(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer("⏭ Точка отложена на конец маршрута")
    _, order_str = query.data.split(":", 1)
    order = int(order_str)

    state = _load_state()
    p = _get_point_by_order(state, order)
    if p and p.get("status") in ("pending", "skipped"):
        p["status"] = "skipped"
        _save_state(state)

    pts = state.get("points", [])
    next_pending = next((pt for pt in pts if pt.get("status") == "pending" and pt.get("order", 0) > order), None)
    if not next_pending:
        next_pending = next((pt for pt in pts if pt.get("status") == "pending"), None)

    if next_pending:
        await _send_point_card(query, state, next_pending["order"])
    else:
        await _send_point_card(query, state)


async def cb_closed(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer("🔒 Отмечено: Магазин закрыт")
    _, order_str = query.data.split(":", 1)
    order = int(order_str)

    state = _load_state()
    p = _get_point_by_order(state, order)
    if p:
        p["status"] = "closed"
        p["visit_time"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        _save_state(state)

    await query.message.reply_text(_progress_summary(state), parse_mode=ParseMode.MARKDOWN)

    next_p, _ = _next_to_visit(state)
    if next_p is None:
        await query.message.reply_text(
            "🎉 *Все точки маршрута пройдены!*",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=MAIN_KB,
        )
        await _auto_send_report(query.message)
    else:
        await _send_point_card(query, state)


# ── Диалог визита (ConversationHandler) ──────────────────

async def cb_visit_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, order_str = query.data.split(":", 1)
    order = int(order_str)

    ctx.user_data["visit_order"] = order
    state = _load_state()
    p = _get_point_by_order(state, order)
    is_append = p and p.get("status") == "visited"
    p_name = _esc(p.get("name", "")) if p else f"Точка {order}"

    prompt = (
        f"✍️ Дополнение по точке *{order}. {p_name}* — выберите категорию (или сразу отправьте текст/фото):"
        if is_append
        else f"🏷 Точка *{order}. {p_name}* — выберите причину/категорию (или сразу напишите текст в чат):"
    )

    kb = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("💸 Цена/Тариф", callback_data="tag:price")],
            [InlineKeyboardButton("⚙️ Тех. проблемы/Баги", callback_data="tag:tech")],
            [InlineKeyboardButton("🔄 Ушли к конкуренту", callback_data="tag:competitor")],
            [InlineKeyboardButton("🤝 Лояльны / Работают", callback_data="tag:loyal")],
            [InlineKeyboardButton("✍️ Сразу текст / фото", callback_data="tag:freetext")],
        ]
    )
    await query.message.reply_text(prompt, parse_mode=ParseMode.MARKDOWN, reply_markup=kb)
    return WAITING_NOTE


async def cb_tag_selected(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, tag_key = query.data.split(":", 1)

    tag_label = REASON_TAGS.get(tag_key, tag_key)
    ctx.user_data["visit_tag"] = tag_label

    if tag_key == "freetext":
        await query.message.reply_text(
            "✍️ Напишите, что сказали в магазине (можно также отправить 📸 фото с подписью):"
        )
    else:
        await query.message.reply_text(
            f"Выбрана категория: *{_esc(tag_label)}*\n\n"
            "Напишите детали разговора в магазине (можно прикрепить 📸 фото) "
            "или отправьте *-* если без комментариев:",
            parse_mode=ParseMode.MARKDOWN,
        )
    return WAITING_NOTE


async def _finalize_visit_record(update: Update, ctx: ContextTypes.DEFAULT_TYPE, note_text: str, photo_filename: str = ""):
    order = ctx.user_data.get("visit_order")
    tag = ctx.user_data.get("visit_tag", "")

    if not order:
        await update.message.reply_text("⚠️ Не указан номер точки для записи.", reply_markup=MAIN_KB)
        return ConversationHandler.END

    if note_text == "-":
        note_text = ""

    state = _load_state()
    p = _get_point_by_order(state, order)

    if p:
        is_append = p.get("status") == "visited"
        if is_append:
            if note_text:
                old_notes = p.get("notes", "")
                p["notes"] = f"{old_notes}; {note_text}" if old_notes else note_text
            if tag and tag != "✍️ Сразу текст":
                p["reason_tag"] = tag
            if photo_filename:
                p["photo"] = photo_filename
        else:
            p["status"] = "visited"
            p["visit_time"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            if tag and tag != "✍️ Сразу текст":
                p["reason_tag"] = tag
            p["notes"] = note_text
            if photo_filename:
                p["photo"] = photo_filename

        _save_state(state)

    await update.message.reply_text(
        f"✅ Заметка по точке *{order}* сохранена!\n\n" + _progress_summary(state),
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=MAIN_KB,
    )

    next_p, _ = _next_to_visit(state)
    if next_p is None:
        await update.message.reply_text(
            "🎉 *Все точки маршрута пройдены!*",
            parse_mode=ParseMode.MARKDOWN,
        )
        await _auto_send_report(update.message)
    else:
        await _send_point_card(update.message, state)

    ctx.user_data.pop("visit_order", None)
    ctx.user_data.pop("visit_tag", None)
    return ConversationHandler.END


async def handle_note_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    note_text = update.message.text.strip()
    return await _finalize_visit_record(update, ctx, note_text)


async def handle_note_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    PHOTOS_DIR.mkdir(exist_ok=True)
    order = ctx.user_data.get("visit_order", 0)
    photo = update.message.photo[-1]
    tg_file = await ctx.bot.get_file(photo.file_id)
    fname = f"point_{order}_{datetime.now():%Y%m%d_%H%M%S}.jpg"
    await tg_file.download_to_drive(str(PHOTOS_DIR / fname))

    caption = (update.message.caption or "").strip()
    return await _finalize_visit_record(update, ctx, caption, photo_filename=fname)


async def cancel_visit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.pop("visit_order", None)
    ctx.user_data.pop("visit_tag", None)
    await update.message.reply_text("❌ Запись заметки отменена.", reply_markup=MAIN_KB)
    return ConversationHandler.END


async def cancel_and_start_walk(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await cancel_visit(update, ctx)
    return await cmd_start_walk(update, ctx)


async def cancel_and_show_all(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await cancel_visit(update, ctx)
    return await cmd_show_all(update, ctx)


async def cancel_and_download_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await cancel_visit(update, ctx)
    return await cmd_download_report(update, ctx)


async def cancel_and_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await cancel_visit(update, ctx)
    return await cmd_reset(update, ctx)


# ── Скачать Excel-отчет ──────────────────────────────────

async def cmd_download_report(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    state = _load_state()
    if not state.get("points"):
        await update.message.reply_text(
            "📤 Сначала загрузите xlsx и рассчитайте маршрут.",
            reply_markup=MAIN_KB,
        )
        return

    wait_msg = await update.message.reply_text("⏳ Формирую актуальный Excel-отчёт…")
    try:
        report_path = export_report_xlsx()
        with open(report_path, "rb") as f:
            await update.message.reply_document(
                document=InputFile(f, filename="Маршрут_Астана_отчет.xlsx"),
                caption="📊 Итоговый Excel-отчёт обхода (с аналитикой оттока)",
            )
        await wait_msg.delete()
    except Exception as e:
        await wait_msg.edit_text(f"❌ Ошибка формирования отчёта: {e}")


# ── Сбросить прогресс ────────────────────────────────────

async def cmd_reset(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    state = _load_state()
    if not state.get("points"):
        await update.message.reply_text("Маршрут ещё не загружен.", reply_markup=MAIN_KB)
        return

    for p in state["points"]:
        p["status"] = "pending"
        p["visit_time"] = ""
        p["reason_tag"] = ""
        p["notes"] = ""
        p["photo"] = ""
    _save_state(state)

    await update.message.reply_text(
        "🔄 Прогресс сброшен. Все точки снова в статусе «⏳ Ожидает».",
        reply_markup=MAIN_KB,
    )


# ── Обработка xlsx-файла ─────────────────────────────────

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type not in (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
    ) and not (doc.file_name or "").lower().endswith(".xlsx"):
        await update.message.reply_text(
            "❌ Отправьте файл в формате *.xlsx*",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    status = await update.message.reply_text("⏳ Загружаю файл…", reply_markup=MAIN_KB)
    try:
        tg_file = await ctx.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(str(SRC_FILE))
    except Exception as e:
        await status.edit_text(
            f"❌ Не удалось сохранить файл: {e}\n"
            "Если файл открыт в Excel на компьютере — пожалуйста, закройте его."
        )
        return

    await status.edit_text("🔄 Считаю оптимальный маршрут по дорогам Астаны… ⏳")

    try:
        proc = await asyncio.create_subprocess_exec(
            PYTHON,
            str(SCRIPT),
            cwd=str(WORK_DIR),
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.STDOUT,
        )
        stdout, _ = await asyncio.wait_for(proc.communicate(), timeout=600)
        output = stdout.decode("utf-8", errors="replace")
    except asyncio.TimeoutError:
        await status.edit_text("❌ Таймаут: расчёт занял больше 10 минут.")
        return
    except Exception as e:
        await status.edit_text(f"❌ Ошибка запуска скрипта: {e}")
        return

    if proc.returncode != 0:
        snippet = output[:3000] if output else "неизвестная ошибка"
        await status.edit_text(
            f"❌ Ошибка при расчёте:\n```\n{snippet}\n```",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    state = _init_state_from_route()
    ctx.user_data["has_route"] = True
    await status.edit_text("✅ Маршрут успешно рассчитан!")

    cur_file = _get_out_file()
    if cur_file.exists():
        with open(cur_file, "rb") as f:
            await update.message.reply_document(
                document=InputFile(f, filename="Маршрут_Астана_готовый.xlsx"),
                caption="📎 Готовый маршрут со ссылками Google Maps и 2ГИС",
            )

    if output.strip():
        if len(output) > 3500:
            output = output[:3500] + "\n…"
        await update.message.reply_text(
            f"📝 *Лог геокодинга и расчёта:*\n```\n{output}\n```",
            parse_mode=ParseMode.MARKDOWN,
        )

    await update.message.reply_text(
        _progress_summary(state) + "\n\nНажмите *▶️ Начать / Продолжить обход*",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=MAIN_KB,
    )


# ── Обработка геолокации — пересчёт маршрута ─────────────

async def handle_location(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    loc = update.message.location
    lat, lon = loc.latitude, loc.longitude

    state = _load_state()
    points = state.get("points", [])
    if not points:
        await update.message.reply_text(
            "📤 Сначала загрузите xlsx и рассчитайте маршрут.",
            reply_markup=MAIN_KB,
        )
        return

    remaining = [
        p
        for p in points
        if p.get("status") in ("pending", "skipped") and p.get("lat") and p.get("lon")
    ]
    remaining_no_coords = [
        p
        for p in points
        if p.get("status") in ("pending", "skipped") and not (p.get("lat") and p.get("lon"))
    ]

    if not remaining and not remaining_no_coords:
        await update.message.reply_text(
            "✅ Все точки уже пройдены, пересчитывать нечего.",
            reply_markup=MAIN_KB,
        )
        return

    wait_msg = await update.message.reply_text(
        f"📍 Ваша геопозиция: {lat:.5f}, {lon:.5f}\n"
        f"🔄 Перестраиваю оптимальный порядок для {len(remaining)} оставшихся точек…"
    )

    ordered, minutes = [], []
    if remaining:
        try:
            sys.path.insert(0, str(WORK_DIR))
            from route_planner import reroute_remaining

            rp_points = []
            for p in remaining:
                rp_points.append(
                    {
                        "order": p["order"],
                        "name": p.get("name", ""),
                        "addr": p.get("addr", ""),
                        "owner": p.get("owner", ""),
                        "contact": p.get("contact", ""),
                        "phone": p.get("phone", ""),
                        "xy": (p["lat"], p["lon"]),
                        "src": p.get("src", ""),
                        "_status": p.get("status", "pending"),
                        "_visit_time": p.get("visit_time", ""),
                        "_reason_tag": p.get("reason_tag", ""),
                        "_notes": p.get("notes", ""),
                        "_photo": p.get("photo", ""),
                        "_gmaps": p.get("gmaps", ""),
                        "_dgis": p.get("dgis", ""),
                    }
                )

            ordered, minutes = reroute_remaining(lat, lon, rp_points)
        except Exception as e:
            await wait_msg.edit_text(f"❌ Ошибка пересчёта: {e}")
            return
    else:
        ordered, minutes = [], []

    visited_closed = [p for p in points if p.get("status") in ("visited", "closed")]
    new_points = list(visited_closed)

    cumulative = visited_closed[-1].get("total", 0) if visited_closed else 0
    for idx, (rp, m) in enumerate(zip(ordered, minutes)):
        cumulative += m
        new_order = len(visited_closed) + idx + 1
        new_points.append(
            {
                "order": new_order,
                "name": rp["name"],
                "addr": rp["addr"],
                "owner": rp["owner"],
                "contact": rp["contact"],
                "phone": rp["phone"],
                "mins": round(m),
                "total": round(cumulative),
                "src": rp["src"],
                "gmaps": rp.get("_gmaps", ""),
                "dgis": rp.get("_dgis", ""),
                "lat": rp["xy"][0],
                "lon": rp["xy"][1],
                "status": rp.get("_status", "pending"),
                "visit_time": rp.get("_visit_time", ""),
                "reason_tag": rp.get("_reason_tag", ""),
                "notes": rp.get("_notes", ""),
                "photo": rp.get("_photo", ""),
            }
        )

    # Добавляем оставшиеся точки без координат в конец маршрута, чтобы они не терялись
    for p in remaining_no_coords:
        new_points.append(p)

    for i, p in enumerate(new_points, 1):
        p["order"] = i

    state["points"] = new_points
    _save_state(state)

    await wait_msg.edit_text(
        f"✅ *Маршрут перестроен от вашей текущей позиции!*\n"
        f"Осталось точек: *{len(ordered) + len(remaining_no_coords)}*\n\n" + _progress_summary(state),
        parse_mode=ParseMode.MARKDOWN,
    )
    await _send_point_card(update.message, state)


# ══════════════════════════════════════════════════════════
# ── Запуск ───────────────────────────────────────────────
# ══════════════════════════════════════════════════════════

def main():
    if OUT_FILE.exists() and not STATE_FILE.exists():
        _init_state_from_route()

    app = Application.builder().token(TOKEN).build()

    visit_conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(cb_visit_start, pattern=r"^visit:\d+$"),
        ],
        states={
            WAITING_NOTE: [
                CallbackQueryHandler(cb_tag_selected, pattern=r"^tag:"),
                MessageHandler(filters.PHOTO, handle_note_photo),
                # Не перехватываем системные кнопки главного меню
                MessageHandler(
                    filters.TEXT & ~filters.COMMAND & ~filters.Regex(r"^(🔄|▶️|📋|📊)"),
                    handle_note_text,
                ),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel_visit),
            MessageHandler(filters.Regex(r"^▶️"), cancel_and_start_walk),
            MessageHandler(filters.Regex(r"^📋"), cancel_and_show_all),
            MessageHandler(filters.Regex(r"^📊"), cancel_and_download_report),
            MessageHandler(filters.Regex(r"^🔄"), cancel_and_reset),
        ],
        per_message=False,
    )

    app.add_handler(visit_conv)
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("route", cmd_show_all))

    app.add_handler(CallbackQueryHandler(cb_city_select, pattern=r"^city:"))
    app.add_handler(CallbackQueryHandler(cb_navigation, pattern=r"^nav:\d+$"))
    app.add_handler(CallbackQueryHandler(cb_skip, pattern=r"^skip:\d+$"))
    app.add_handler(CallbackQueryHandler(cb_closed, pattern=r"^closed:\d+$"))

    app.add_handler(MessageHandler(filters.Regex(r"^▶️"), cmd_start_walk))
    app.add_handler(MessageHandler(filters.Regex(r"^📋"), cmd_show_all))
    app.add_handler(MessageHandler(filters.Regex(r"^📊"), cmd_download_report))
    app.add_handler(MessageHandler(filters.Regex(r"^🔄"), cmd_reset))

    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))

    log.info("🤖 Бот запущен, ожидаю сообщения…")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
