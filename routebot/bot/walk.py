"""Обход: старт, карточки, навигация, «Отложить», «Закрыт», «Вернуть», список маршрута."""
from __future__ import annotations

from telegram import Update
from telegram.ext import ContextTypes

from ..models import Status, now_str
from . import keyboards as kb
from . import views
from .common import callback_point, continue_walk, edit_card, load_state, reply, save_state


async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await reply(
        update.effective_message,
        "👋 <b>Привет! Это бот-ассистент для объезда торговых точек и сбора обратной связи.</b>\n\n"
        "Я строю самый быстрый маршрут по дорогам, веду по точкам и сохраняю ваши заметки "
        "в Excel-отчёт.\n\n" + views.HELP,
        reply_markup=kb.MAIN_KB,
    )
    await update.effective_message.reply_text("🏙 Выберите город:", reply_markup=kb.CITY_KB)


async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await reply(update.effective_message, views.HELP, reply_markup=kb.MAIN_KB)


async def cb_city(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    city = query.data.split(":", 1)[1]
    if city != "Астана":
        await reply(query.message, f"🚧 Город <b>{views.esc(city)}</b> скоро будет подключён. Пока работает только Астана.")
        return
    state = load_state(update, ctx)
    state.city = city
    save_state(update, ctx, state)
    await reply(
        query.message,
        f"✅ Город: <b>{views.esc(city)}</b>.\nОтправьте .xlsx для расчёта или нажмите «{kb.BTN_WALK}».",
        reply_markup=kb.MAIN_KB,
    )


async def cmd_walk(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await continue_walk(update.effective_message, ctx, load_state(update, ctx))


async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    state = load_state(update, ctx)
    if not state.points:
        await reply(update.effective_message, views.NO_ROUTE, reply_markup=kb.MAIN_KB)
        return
    text = views.progress(state)
    link = views.next_stops_link(state)
    await reply(update.effective_message, text + (f"\n{link}" if link else ""), reply_markup=kb.MAIN_KB)


async def cmd_route(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    state = load_state(update, ctx)
    if not state.points:
        await reply(update.effective_message, views.NO_ROUTE, reply_markup=kb.MAIN_KB)
        return
    for page in views.route_pages(state):
        await reply(update.effective_message, page)


async def cb_nav(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    state = load_state(update, ctx)
    p = state.by_order(int(query.data.split(":", 1)[1]))
    if p is None:
        await query.answer("Такой точки уже нет — маршрут обновился.", show_alert=True)
        return
    await query.answer()
    await edit_card(query, state, p)


async def cb_skip(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    state, p = await callback_point(update, ctx)
    if p is None:
        return
    query = update.callback_query
    if not p.status.is_active:
        await query.answer(f"Точка уже отмечена: {p.status.label}. Сначала «↩️ Вернуть».", show_alert=True)
        return
    p.status = Status.SKIPPED
    save_state(update, ctx, state)
    await query.answer("⏭ Точка отложена на конец маршрута")
    await continue_walk(query.message, ctx, state, after_order=p.order)


async def cb_closed(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    state, p = await callback_point(update, ctx)
    if p is None:
        return
    query = update.callback_query
    p.status, p.visit_time = Status.CLOSED, now_str()
    save_state(update, ctx, state)
    await query.answer("🔒 Отмечено: магазин закрыт")
    await reply(query.message, views.progress(state))
    await continue_walk(query.message, ctx, state, after_order=p.order)


async def cb_undo(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Ошибочно нажали «Закрыт» / «Отложить» / записали визит — возвращаем в ожидание. Заметки остаются."""
    state, p = await callback_point(update, ctx)
    if p is None:
        return
    p.status, p.visit_time = Status.PENDING, ""
    save_state(update, ctx, state)
    await update.callback_query.answer("↩️ Точка снова в статусе «Ожидает»")
    await edit_card(update.callback_query, state, p)


async def cb_stale_tag(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Кнопка категории вне диалога (например, диалог уже закрыт)."""
    await update.callback_query.answer(
        "Этот диалог уже закрыт. Нажмите «✅ Записать итог» на карточке ещё раз.", show_alert=True
    )
