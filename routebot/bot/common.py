"""Общее для хэндлеров: рантайм-зависимости, состояние чата, отправка карточек и отчёта."""
from __future__ import annotations

import asyncio
import contextlib
import logging
from dataclasses import dataclass, field

from telegram import CallbackQuery, InputFile, Message, Update
from telegram.constants import ParseMode
from telegram.error import BadRequest, TelegramError
from telegram.ext import ContextTypes

from ..excel import write_report
from ..geocoding import Geocoder
from ..models import Point, RouteState
from ..planner import RoutePlanner
from ..storage import StateStore
from . import keyboards as kb
from . import views

log = logging.getLogger(__name__)

REPORT_FILENAME = "Маршрут_Астана_отчет.xlsx"
ROUTE_FILENAME = "Маршрут_Астана_готовый.xlsx"


@dataclass
class Runtime:
    """Зависимости бота. Лежит в bot_data (не персистится) — создаётся при старте."""

    store: StateStore
    geocoder: Geocoder
    planner: RoutePlanner
    busy: set[int] = field(default_factory=set)  # чаты, где идёт долгий расчёт


def runtime(ctx: ContextTypes.DEFAULT_TYPE) -> Runtime:
    return ctx.application.bot_data["runtime"]


def load_state(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> RouteState:
    return runtime(ctx).store.load(update.effective_chat.id)


def save_state(update: Update, ctx: ContextTypes.DEFAULT_TYPE, state: RouteState) -> None:
    runtime(ctx).store.save(update.effective_chat.id, state)


async def reply(message: Message, text: str, **kwargs) -> Message:
    kwargs.setdefault("disable_web_page_preview", True)
    return await message.reply_text(text, parse_mode=ParseMode.HTML, **kwargs)


async def callback_point(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> tuple[RouteState, Point | None]:
    """Точка из callback_data вида «action:id». Если её уже нет — сообщает об этом пользователю."""
    query = update.callback_query
    state = load_state(update, ctx)
    try:
        p = state.by_id(int(query.data.split(":", 1)[1]))
    except (IndexError, ValueError):
        p = None
    if p is None:
        await query.answer("Точка не найдена — маршрут обновился. Откройте карточку заново.", show_alert=True)
    return state, p


async def send_card(message: Message, state: RouteState, p: Point, prefix: str = "") -> None:
    await reply(message, views.card(state, p, prefix), reply_markup=kb.card_kb(p, len(state.points)))


async def edit_card(query: CallbackQuery, state: RouteState, p: Point) -> None:
    """Обновляет карточку на месте; если сообщение не редактируется — присылает новую."""
    try:
        await query.edit_message_text(
            views.card(state, p),
            parse_mode=ParseMode.HTML,
            disable_web_page_preview=True,
            reply_markup=kb.card_kb(p, len(state.points)),
        )
    except BadRequest as e:
        if "not modified" in str(e).lower():
            return
        await send_card(query.message, state, p)


async def continue_walk(
    message: Message, ctx: ContextTypes.DEFAULT_TYPE, state: RouteState, after_order: int | None = None
) -> None:
    """Следующая карточка или, если всё пройдено, поздравление + отчёт."""
    if not state.points:
        await reply(message, views.NO_ROUTE, reply_markup=kb.MAIN_KB)
        return
    p, from_skipped = state.next_to_visit(after_order)
    if p is None:
        await reply(message, views.all_done(state), reply_markup=kb.MAIN_KB)
        await send_report(message, ctx, state)
        return
    await send_card(message, state, p, views.BACK_TO_SKIPPED if from_skipped else "")


async def send_report(
    message: Message,
    ctx: ContextTypes.DEFAULT_TYPE,
    state: RouteState,
    caption: str = "📊 Отчёт обхода: статусы, заметки и аналитика оттока",
) -> None:
    path = await asyncio.to_thread(write_report, runtime(ctx).store.report_path(message.chat_id), state)
    with open(path, "rb") as f:
        await message.reply_document(InputFile(f, filename=REPORT_FILENAME), caption=caption)


class ProgressMessage:
    """
    Показывает прогресс долгой операции из рабочего потока: поток только меняет текст
    (update), а асинхронный цикл раз в interval секунд редактирует одно сообщение.
    Так не нужно передавать корутины между потоками и не упираемся в лимиты Telegram.
    """

    def __init__(self, message: Message, interval: float = 3.0):
        self.message = message
        self.interval = interval
        self._text = ""
        self._shown = message.text or ""
        self._task: asyncio.Task | None = None

    def update(self, text: str) -> None:
        self._text = text

    async def __aenter__(self) -> ProgressMessage:
        self._task = asyncio.create_task(self._loop())
        return self

    async def __aexit__(self, *exc) -> None:
        if self._task:
            self._task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await self._task

    async def _loop(self) -> None:
        while True:
            await asyncio.sleep(self.interval)
            text = self._text
            if text and text != self._shown:
                with contextlib.suppress(TelegramError):
                    await self.message.edit_text(text)
                    self._shown = text
