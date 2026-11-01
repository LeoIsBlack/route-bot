"""Клавиатуры и формат callback_data.

callback_data для действий с точкой содержит её постоянный id, а не порядковый номер:
после пересчёта маршрута номера меняются, и старые кнопки попадали бы не в тот магазин.
"""
from __future__ import annotations

from telegram import InlineKeyboardButton as Btn
from telegram import InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup

from ..models import Point, Status

# ── Тексты кнопок главного меню (по ним же матчатся хэндлеры) ──
BTN_WALK = "▶️ Начать / Продолжить обход"
BTN_ROUTE = "📋 Весь маршрут"
BTN_REPORT = "📊 Скачать Excel-отчет"
BTN_RESET = "🔄 Сбросить прогресс"
BTN_LOCATION = "📍 Пересчитать от меня"
BTN_PIN = "📌 Отправить мою точку"
BTN_CANCEL = "❌ Отмена"

MENU_TEXTS = [BTN_WALK, BTN_ROUTE, BTN_REPORT, BTN_RESET, BTN_CANCEL]

MAIN_KB = ReplyKeyboardMarkup(
    [
        [BTN_WALK, BTN_ROUTE],
        [BTN_REPORT, BTN_RESET],
        [KeyboardButton(BTN_LOCATION, request_location=True)],
    ],
    resize_keyboard=True,
)

PIN_KB = ReplyKeyboardMarkup(
    [[KeyboardButton(BTN_PIN, request_location=True)], [BTN_CANCEL]],
    resize_keyboard=True,
    one_time_keyboard=True,
)

CITY_KB = InlineKeyboardMarkup(
    [
        [Btn("📍 Астана", callback_data="city:Астана"), Btn("📍 Алматы (скоро)", callback_data="city:Алматы")],
        [Btn("📍 Шымкент (скоро)", callback_data="city:Шымкент"), Btn("📍 Костанай (скоро)", callback_data="city:Костанай")],
    ]
)

# Категории итога визита: ключ в callback_data → подпись (она же пишется в отчёт)
REASON_TAGS = {
    "price": "💸 Цена/Тариф",
    "tech": "⚙️ Тех. проблемы/Баги",
    "competitor": "🔄 Ушли к конкуренту",
    "loyal": "🤝 Лояльны / Работают",
}
TAG_FREETEXT = "freetext"

TAGS_KB = InlineKeyboardMarkup(
    [[Btn(label, callback_data=f"tag:{key}")] for key, label in REASON_TAGS.items()]
    + [[Btn("✍️ Сразу текст / фото", callback_data=f"tag:{TAG_FREETEXT}")]]
)

RESET_CONFIRM_KB = InlineKeyboardMarkup(
    [[Btn("🗑 Да, удалить всё", callback_data="reset:yes"), Btn("Отмена", callback_data="reset:no")]]
)


def card_kb(p: Point, total: int) -> InlineKeyboardMarkup:
    visit_label = "✍️ Дополнить заметку / фото" if p.status is Status.VISITED else "✅ Записать итог / заметку"
    rows = [
        [Btn(visit_label, callback_data=f"visit:{p.id}")],
        [Btn("⏭ Отложить на потом", callback_data=f"skip:{p.id}"), Btn("🔒 Магазин закрыт", callback_data=f"closed:{p.id}")],
    ]
    if p.status is not Status.PENDING:
        rows.append([Btn("↩️ Вернуть в «Ожидает»", callback_data=f"undo:{p.id}")])
    nav = []
    if p.order > 1:
        nav.append(Btn("⬅️ Пред.", callback_data=f"nav:{p.order - 1}"))
    nav.append(Btn("📌 Я тут", callback_data=f"pin:{p.id}"))
    if p.order < total:
        nav.append(Btn("След. ➡️", callback_data=f"nav:{p.order + 1}"))
    rows.append(nav)
    return InlineKeyboardMarkup(rows)
