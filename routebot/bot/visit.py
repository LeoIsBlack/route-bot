"""Диалог записи итога визита: категория → текст / фото / альбом."""
from __future__ import annotations

from datetime import datetime

from telegram import Message, Update
from telegram.ext import ContextTypes, ConversationHandler

from ..models import Status, now_str
from . import keyboards as kb
from . import views
from .common import callback_point, continue_walk, load_state, reply, runtime, save_state

WAITING_NOTE = 0

# Ключи user_data
VISIT_ID = "visit_id"
VISIT_TAG = "visit_tag"
ALBUM = "album"  # [media_group_id, point_id] — куда складывать остальные фото альбома


def clear_visit(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    ctx.user_data.pop(VISIT_ID, None)
    ctx.user_data.pop(VISIT_TAG, None)


async def cb_visit_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    _, p = await callback_point(update, ctx)
    if p is None:
        return ConversationHandler.END
    query = update.callback_query
    await query.answer()
    clear_visit(ctx)
    ctx.user_data[VISIT_ID] = p.id

    title = f"<b>{p.order}. {views.esc(p.name)}</b>"
    prompt = (
        f"✍️ Дополнение по точке {title} — выберите категорию или сразу отправьте текст / фото:"
        if p.status is Status.VISITED
        else f"🏷 Точка {title} — выберите причину / категорию или сразу напишите текст:"
    )
    await reply(query.message, prompt, reply_markup=kb.TAGS_KB)
    return WAITING_NOTE


async def cb_tag(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    label = kb.REASON_TAGS.get(key)
    if label:
        ctx.user_data[VISIT_TAG] = label
        text = (
            f"Категория: <b>{views.esc(label)}</b>\n\n"
            "Напишите детали разговора (можно 📸 фото или альбом) или отправьте <b>-</b>, если без комментариев:"
        )
    else:
        ctx.user_data.pop(VISIT_TAG, None)
        text = "✍️ Напишите, что сказали в магазине (можно 📸 фото с подписью):"
    await reply(query.message, text)
    return WAITING_NOTE


async def on_note_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    note = update.message.text.strip()
    return await _save_visit(update, ctx, "" if note == "-" else note)


async def on_note_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message
    pid = ctx.user_data.get(VISIT_ID)
    filename = await _download_photo(update, ctx, pid)
    if msg.media_group_id:
        # Остальные фото альбома придут отдельными апдейтами уже после конца диалога
        ctx.user_data[ALBUM] = [msg.media_group_id, pid]
    return await _save_visit(update, ctx, msg.caption or "", photo=filename)


async def on_album_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Фото вне диалога: либо продолжение альбома (догружаем к той же точке), либо подсказка."""
    msg = update.message
    album = ctx.user_data.get(ALBUM)
    if not (album and msg.media_group_id and msg.media_group_id == album[0]):
        await reply(msg, "📸 Чтобы прикрепить фото, нажмите «✅ Записать итог» на карточке точки.", reply_markup=kb.MAIN_KB)
        return
    filename = await _download_photo(update, ctx, album[1])
    state = load_state(update, ctx)
    p = state.by_id(album[1])
    if p:
        p.add_photo(filename)
        p.add_note(msg.caption or "")
        save_state(update, ctx, state)


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    clear_visit(ctx)
    await reply(update.effective_message, "❌ Запись заметки отменена.", reply_markup=kb.MAIN_KB)
    return ConversationHandler.END


def end_visit_then(handler):
    """Обёртка для fallbacks: закрыть незаконченную заметку и выполнить действие.
    Иначе следующий текст молча записался бы в точку, по которой начали заметку."""

    async def wrapped(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
        clear_visit(ctx)
        await handler(update, ctx)
        return ConversationHandler.END

    return wrapped


async def _save_visit(update: Update, ctx: ContextTypes.DEFAULT_TYPE, note: str, photo: str = "") -> int:
    msg: Message = update.message
    pid = ctx.user_data.get(VISIT_ID)
    tag = ctx.user_data.get(VISIT_TAG, "")
    clear_visit(ctx)

    state = load_state(update, ctx)
    p = state.by_id(pid)
    if p is None:
        await reply(msg, "⚠️ Точка не найдена — маршрут обновился. Откройте карточку заново.", reply_markup=kb.MAIN_KB)
        return ConversationHandler.END

    if p.status is not Status.VISITED:
        p.status, p.visit_time = Status.VISITED, now_str()
    if tag:
        p.reason_tag = tag
    p.add_note(note)  # дописываем: заметки могли остаться после «Закрыт» → «Вернуть»
    p.add_photo(photo)
    save_state(update, ctx, state)

    await reply(
        msg,
        f"✅ Заметка по точке <b>{p.order}. {views.esc(p.name)}</b> сохранена!\n\n" + views.progress(state),
        reply_markup=kb.MAIN_KB,
    )
    await continue_walk(msg, ctx, state, after_order=p.order)
    return ConversationHandler.END


async def _download_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE, pid) -> str:
    msg = update.message
    if msg.photo:
        file_id, ext = msg.photo[-1].file_id, ".jpg"
    else:  # фото, отправленное файлом (без сжатия)
        file_id = msg.document.file_id
        ext = "." + (msg.document.file_name or "photo.jpg").rsplit(".", 1)[-1].lower()
    # message_id в имени: фото одного альбома приходят в одну секунду
    filename = f"point_{pid}_{datetime.now():%Y%m%d_%H%M%S}_{msg.message_id}{ext}"
    tg_file = await ctx.bot.get_file(file_id)
    await tg_file.download_to_drive(runtime(ctx).store.photos_dir(update.effective_chat.id) / filename)
    return filename
