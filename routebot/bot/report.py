"""Отчёт и сброс прогресса."""
from __future__ import annotations

from telegram import Update
from telegram.ext import ContextTypes

from ..models import Status
from . import keyboards as kb
from . import views
from .common import load_state, reply, save_state, send_report


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    state = load_state(update, ctx)
    if not state.points:
        await reply(msg, views.NO_ROUTE, reply_markup=kb.MAIN_KB)
        return
    wait = await msg.reply_text("⏳ Формирую актуальный Excel-отчёт…")
    await send_report(msg, ctx, state)
    await wait.delete()


async def cmd_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    state = load_state(update, ctx)
    if not state.points:
        await reply(msg, "Маршрут ещё не загружен.", reply_markup=kb.MAIN_KB)
        return
    touched = sum(1 for p in state.points if p.status is not Status.PENDING or p.notes or p.photos)
    if not touched:
        await reply(msg, "Сбрасывать нечего — ни одна точка ещё не отмечена.", reply_markup=kb.MAIN_KB)
        return
    await reply(
        msg,
        f"⚠️ Будут удалены статусы, заметки и фото по <b>{touched}</b> точкам. "
        "Если отчёт нужен — сначала скачайте его. Сбросить?",
        reply_markup=kb.RESET_CONFIRM_KB,
    )


async def cb_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    if query.data != "reset:yes":
        await query.edit_message_text("Сброс отменён, прогресс сохранён.")
        return
    state = load_state(update, ctx)
    state.reset_progress()
    save_state(update, ctx, state)
    await query.edit_message_text("🔄 Прогресс сброшен. Все точки снова в статусе «⏳ Ожидает».")
