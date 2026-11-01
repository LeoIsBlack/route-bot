"""Сборка Telegram-приложения: зависимости, хэндлеры, персистентность, обработка ошибок."""
from __future__ import annotations

import logging
import sys
import warnings
from pathlib import Path

from telegram import BotCommand, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    PersistenceInput,
    PicklePersistence,
    filters,
)
from telegram.warnings import PTBUserWarning

from .. import config
from ..geocoding import Geocoder
from ..osrm import OsrmClient
from ..planner import RoutePlanner
from ..storage import StateStore
from . import keyboards as kb
from . import planning, report, visit, walk
from .common import Runtime, reply

log = logging.getLogger(__name__)

# per_message=False осознанно: диалог заметки начинается с кнопки, а продолжается текстом/фото
warnings.filterwarnings("ignore", message=r".*per_message=False.*", category=PTBUserWarning)

# (кнопка меню, команда, хэндлер) — одни и те же действия доступны кнопкой и командой
MENU = [
    (kb.BTN_WALK, "walk", walk.cmd_walk),
    (kb.BTN_ROUTE, "route", walk.cmd_route),
    (kb.BTN_REPORT, "report", report.cmd_report),
    (kb.BTN_RESET, "reset", report.cmd_reset),
]
POINT_CALLBACKS = [
    (r"^nav:\d+$", walk.cb_nav),
    (r"^skip:\d+$", walk.cb_skip),
    (r"^closed:\d+$", walk.cb_closed),
    (r"^undo:\d+$", walk.cb_undo),
    (r"^pin:\d+$", planning.cb_pin),
]
COMMANDS = [
    BotCommand("walk", "Продолжить обход"),
    BotCommand("route", "Весь маршрут"),
    BotCommand("status", "Прогресс и прогноз"),
    BotCommand("report", "Excel-отчёт"),
    BotCommand("cancel", "Отменить ввод"),
    BotCommand("help", "Как пользоваться"),
]


def _visit_conversation() -> ConversationHandler:
    end = visit.end_visit_then
    fallbacks = [
        CommandHandler("cancel", visit.cancel),
        MessageHandler(filters.Text([kb.BTN_CANCEL]), visit.cancel),
    ]
    for text, command, handler in MENU:
        fallbacks += [
            MessageHandler(filters.Text([text]), end(handler)),
            CommandHandler(command, end(handler)),
        ]
    fallbacks += [CallbackQueryHandler(end(h), pattern=p) for p, h in POINT_CALLBACKS]
    fallbacks.append(CommandHandler("status", end(walk.cmd_status)))

    return ConversationHandler(
        entry_points=[CallbackQueryHandler(visit.cb_visit_start, pattern=r"^visit:\d+$")],
        states={
            visit.WAITING_NOTE: [
                CallbackQueryHandler(visit.cb_tag, pattern=r"^tag:"),
                MessageHandler(filters.PHOTO | filters.Document.IMAGE, visit.on_note_photo),
                MessageHandler(filters.TEXT & ~filters.COMMAND & ~filters.Text(kb.MENU_TEXTS), visit.on_note_text),
            ],
        },
        fallbacks=fallbacks,
        allow_reentry=True,  # «Записать итог» на другой карточке, пока открыт диалог
        per_message=False,
        name="visit",
        persistent=True,  # диалог переживает перезапуск бота (run.bat делает git pull + рестарт)
    )


def _register(app: Application) -> None:
    app.add_handler(_visit_conversation())

    app.add_handler(CommandHandler("start", walk.cmd_start))
    app.add_handler(CommandHandler("help", walk.cmd_help))
    app.add_handler(CommandHandler("status", walk.cmd_status))
    app.add_handler(CommandHandler("cancel", planning.on_cancel))
    app.add_handler(MessageHandler(filters.Text([kb.BTN_CANCEL]), planning.on_cancel))
    for text, command, handler in MENU:
        app.add_handler(MessageHandler(filters.Text([text]), handler))
        app.add_handler(CommandHandler(command, handler))

    for pattern, handler in POINT_CALLBACKS:
        app.add_handler(CallbackQueryHandler(handler, pattern=pattern))
    app.add_handler(CallbackQueryHandler(walk.cb_city, pattern=r"^city:"))
    app.add_handler(CallbackQueryHandler(report.cb_reset, pattern=r"^reset:"))
    app.add_handler(CallbackQueryHandler(walk.cb_stale_tag, pattern=r"^tag:"))

    app.add_handler(MessageHandler(filters.Document.ALL, planning.on_document))
    app.add_handler(MessageHandler(filters.LOCATION, planning.on_location))
    app.add_handler(MessageHandler(filters.PHOTO, visit.on_album_photo))


async def _on_error(update: object, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Ошибка при обработке апдейта", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await reply(update.effective_message, "⚠️ Что-то пошло не так. Попробуйте ещё раз.")
        except Exception:  # сеть могла упасть — тогда просто остаётся запись в логе
            pass


async def _post_init(app: Application) -> None:
    await app.bot.set_my_commands(COMMANDS)


def build_application(token: str = config.BOT_TOKEN, data_dir: Path = config.DATA_DIR) -> Application:
    data_dir.mkdir(parents=True, exist_ok=True)
    geocoder = Geocoder()
    runtime = Runtime(
        store=StateStore(data_dir, legacy_state=config.LEGACY_STATE),
        geocoder=geocoder,
        planner=RoutePlanner(geocoder, OsrmClient()),
    )
    persistence = PicklePersistence(
        filepath=data_dir / "bot_persistence.pickle",
        # Персистим только user_data и диалоги; bot_data содержит рантайм-объекты
        store_data=PersistenceInput(bot_data=False, chat_data=False, user_data=True, callback_data=False),
    )
    app = Application.builder().token(token).persistence(persistence).post_init(_post_init).build()
    app.bot_data["runtime"] = runtime
    _register(app)
    app.add_error_handler(_on_error)
    return app


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    logging.basicConfig(
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        level=logging.INFO,
        stream=sys.stdout,
    )
    logging.getLogger("httpx").setLevel(logging.WARNING)  # иначе каждый getUpdates в логе

    app = build_application()
    log.info("🤖 Бот запущен, ожидаю сообщения…")
    # drop_pending_updates=False: заметки, отправленные, пока бот перезапускался, не теряются
    app.run_polling(drop_pending_updates=False, allowed_updates=Update.ALL_TYPES)
