"""Расчёт маршрута из xlsx, пересчёт от GPS и «📌 Я тут» (координаты точки с места).

Долгие операции (геокодирование — минуты, OSRM — секунды) идут фоновыми задачами
в отдельном потоке: бот продолжает отвечать и этому, и другим чатам.
"""
from __future__ import annotations

import asyncio
import logging

from telegram import InputFile, Location, Message, Update
from telegram.ext import ContextTypes

from ..excel import InputError, read_input, write_route
from ..models import Point, RouteState, Status, now_str
from ..osrm import OsrmError
from ..planner import PlanError
from . import keyboards as kb
from . import views
from .common import (
    ROUTE_FILENAME,
    ProgressMessage,
    callback_point,
    load_state,
    reply,
    runtime,
    save_state,
    send_card,
)

log = logging.getLogger(__name__)

PIN_ID = "pin_id"  # ключ user_data: точка, для которой ждём геопозицию
PLAN_TIMEOUT = 15 * 60


# ═══════════════ Загрузка xlsx → расчёт маршрута ═══════════════

async def on_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.message
    name = (msg.document.file_name or "").lower()
    if not name.endswith(".xlsx"):
        hint = (
            "Старый формат .xls не поддерживается — пересохраните файл как .xlsx."
            if name.endswith(".xls")
            else "Отправьте таблицу в формате .xlsx."
        )
        await reply(msg, f"❌ {hint}", reply_markup=kb.MAIN_KB)
        return

    rt, chat_id = runtime(ctx), update.effective_chat.id
    if chat_id in rt.busy:
        await reply(msg, "⏳ Уже идёт расчёт — дождитесь его окончания.")
        return

    src = rt.store.input_path(chat_id)
    tg_file = await msg.document.get_file()
    await tg_file.download_to_drive(src)
    try:
        points = await asyncio.to_thread(read_input, src)
    except InputError as e:
        await reply(msg, f"❌ {views.esc(e)}", reply_markup=kb.MAIN_KB)
        return

    status = await msg.reply_text(f"📥 Точек в файле: {len(points)}. Начинаю расчёт маршрута…")
    rt.busy.add(chat_id)
    ctx.application.create_task(_plan_job(ctx, chat_id, msg, status, points), update=update)


async def _plan_job(ctx, chat_id: int, msg: Message, status: Message, points: list[Point]) -> None:
    rt = runtime(ctx)
    try:
        async with ProgressMessage(status) as progress:
            result = await asyncio.wait_for(
                asyncio.to_thread(rt.planner.plan, points, progress.update), PLAN_TIMEOUT
            )
        # Синхронно, без await между load и save: параллельные отметки не потеряются
        state = RouteState(points=result.points, planned_at=now_str())
        old = rt.store.load(chat_id)
        state.city = old.city
        state.merge_progress(old)
        rt.store.save(chat_id, state)
    except (PlanError, OsrmError) as e:
        await status.edit_text(f"❌ {e}")
        return
    except asyncio.TimeoutError:
        await status.edit_text("❌ Расчёт занял слишком много времени. Попробуйте ещё раз позже.")
        return
    except Exception:
        log.exception("Ошибка расчёта маршрута для чата %s", chat_id)
        await status.edit_text("❌ Непредвиденная ошибка при расчёте маршрута. Подробности в логе бота.")
        return
    finally:
        rt.busy.discard(chat_id)

    await status.delete()
    route_path = await asyncio.to_thread(write_route, rt.store.route_path(chat_id), result.route, result.missing)
    with open(route_path, "rb") as f:
        await msg.reply_document(InputFile(f, filename=ROUTE_FILENAME), caption="📎 Готовый маршрут со ссылками Google Maps и 2ГИС")
    await reply(
        msg,
        views.plan_done(result) + "\n\n" + views.progress(state) + f"\n\nНажмите «{kb.BTN_WALK}»",
        reply_markup=kb.MAIN_KB,
    )


# ═══════════════ Геопозиция: «📌 Я тут» или пересчёт ═══════════════

async def on_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    loc = update.message.location
    pin_id = ctx.user_data.pop(PIN_ID, None)
    if pin_id is not None:
        await _pin_point(update, ctx, pin_id, loc)
        return

    msg = update.message
    state = load_state(update, ctx)
    if not state.points:
        await reply(msg, views.NO_ROUTE, reply_markup=kb.MAIN_KB)
        return
    if not state.remaining():
        await reply(msg, "✅ Все точки уже пройдены — пересчитывать нечего.", reply_markup=kb.MAIN_KB)
        return
    pending = [p for p in state.remaining() if p.has_coords and p.status is Status.PENDING]
    skipped = [p for p in state.remaining() if p.has_coords and p.status is Status.SKIPPED]
    if not pending and not skipped:
        await reply(msg, "⚠️ У оставшихся точек нет координат — откройте их карточки и отметьте «📌 Я тут».")
        return

    rt, chat_id = runtime(ctx), update.effective_chat.id
    if chat_id in rt.busy:
        await reply(msg, "⏳ Уже идёт расчёт — дождитесь его окончания.")
        return
    status = await msg.reply_text(
        f"📍 Ваша позиция: {loc.latitude:.5f}, {loc.longitude:.5f}\n"
        f"🔄 Перестраиваю порядок для {len(pending) + len(skipped)} точек…"
    )
    rt.busy.add(chat_id)
    ctx.application.create_task(
        _reroute_job(ctx, chat_id, msg, status, loc, pending, skipped), update=update
    )


async def _reroute_job(ctx, chat_id, msg: Message, status: Message, loc: Location, pending, skipped) -> None:
    rt = runtime(ctx)
    try:
        sequence = await asyncio.to_thread(rt.planner.reroute, loc.latitude, loc.longitude, pending, skipped)
        # Перечитываем: пока шёл расчёт, агент мог отметить другую точку
        state = rt.store.load(chat_id)
        rerouted = state.apply_reroute(sequence)
        rt.store.save(chat_id, state)
    except (PlanError, OsrmError) as e:
        await status.edit_text(f"❌ Ошибка пересчёта: {e}")
        return
    except Exception:
        log.exception("Ошибка пересчёта для чата %s", chat_id)
        await status.edit_text("❌ Непредвиденная ошибка при пересчёте. Подробности в логе бота.")
        return
    finally:
        rt.busy.discard(chat_id)

    first = rerouted[0] if rerouted else None
    lines = ["✅ <b>Маршрут перестроен от вашей позиции!</b>"]
    if first:
        lines.append(f"➡️ Ближайшая: <b>{views.esc(first.name)}</b> (~{round(first.mins)} мин)")
    if skipped:
        lines.append(f"⏭ Отложенные ({len(skipped)}) — в конце маршрута")
    lines += ["", views.progress(state)]
    link = views.next_stops_link(state)
    if link:
        lines.append(link)
    await status.delete()
    await reply(msg, "\n".join(lines), reply_markup=kb.MAIN_KB)
    if first:
        await send_card(msg, state, first)


# ═══════════════ «📌 Я тут» ═══════════════

async def cb_pin(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    _, p = await callback_point(update, ctx)
    if p is None:
        return
    query = update.callback_query
    await query.answer()
    ctx.user_data[PIN_ID] = p.id
    await reply(
        query.message,
        f"📌 Вы у точки <b>{views.esc(p.name)}</b>?\n"
        f"Нажмите «{kb.BTN_PIN}» — запомню её точные координаты для этого и будущих маршрутов.",
        reply_markup=kb.PIN_KB,
    )


async def _pin_point(update: Update, ctx: ContextTypes.DEFAULT_TYPE, pid: int, loc: Location) -> None:
    msg = update.message
    state = load_state(update, ctx)
    p = state.by_id(pid)
    if p is None:
        await reply(msg, "⚠️ Точка не найдена — маршрут обновился.", reply_markup=kb.MAIN_KB)
        return
    was_missing = not p.has_coords
    p.lat, p.lon, p.src, p.fail_reason = loc.latitude, loc.longitude, "на месте", ""
    save_state(update, ctx, state)
    if p.addr:
        await asyncio.to_thread(runtime(ctx).geocoder.remember, p.addr, p.lat, p.lon)

    text = f"📌 Координаты <b>{views.esc(p.name)}</b> сохранены — в следующих маршрутах адрес будет точным."
    if loc.horizontal_accuracy and loc.horizontal_accuracy > 100:
        text += f"\n⚠️ Точность GPS низкая (±{round(loc.horizontal_accuracy)} м) — при необходимости повторите."
    if was_missing and p.status.is_active:
        text += f"\nНажмите «{kb.BTN_LOCATION}», чтобы встроить точку в маршрут."
    await reply(msg, text, reply_markup=kb.MAIN_KB)


async def on_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """«❌ Отмена» / /cancel вне диалога заметки: выходим из режима «📌 Я тут»."""
    ctx.user_data.pop(PIN_ID, None)
    await reply(update.effective_message, "Отменено.", reply_markup=kb.MAIN_KB)
