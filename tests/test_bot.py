"""Сценарии бота без сети: поддельные Message/CallbackQuery, настоящие хэндлеры и состояние."""
import asyncio
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS

from telegram.ext import ConversationHandler

from routebot.bot import planning, report, visit, walk
from routebot.bot.common import Runtime
from routebot.models import Status
from routebot.planner import RoutePlanner
from routebot.storage import StateStore

from .fakes import FakeGeocoder, FakeOsrm
from .test_excel import make_xlsx

CHAT = 777


class FakeMessage:
    _ids = iter(range(1000, 10**6))

    def __init__(self, chat, text="", **attrs):
        self.chat = chat
        self.chat_id = CHAT
        self.message_id = next(self._ids)
        self.text = text
        self.caption = None
        self.photo = ()
        self.document = None
        self.media_group_id = None
        self.location = None
        self.deleted = False
        self.__dict__.update(attrs)

    async def reply_text(self, text, **kw):
        m = FakeMessage(self.chat, text, markup=kw.get("reply_markup"))
        self.chat.sent.append(m)
        return m

    async def reply_document(self, document, caption=None, **kw):
        self.chat.documents.append(document.filename)

    async def edit_text(self, text, **kw):
        self.text = text

    async def delete(self):
        self.deleted = True


class FakeQuery:
    def __init__(self, chat, data):
        self.data = data
        self.message = FakeMessage(chat, "card")
        self.alerts = []

    async def answer(self, text=None, show_alert=False):
        if show_alert:
            self.alerts.append(text)

    async def edit_message_text(self, text, **kw):
        self.message.text = text


class FakeFile:
    def __init__(self, src: Path | None = None):
        self.src = src

    async def download_to_drive(self, path):
        Path(path).write_bytes(self.src.read_bytes() if self.src else b"jpeg")


class Chat:
    """Один чат: копит отправленные сообщения и выполняет фоновые задачи бота."""

    def __init__(self, rt: Runtime):
        self.sent: list[FakeMessage] = []
        self.documents: list[str] = []
        self.tasks = []
        app = NS(bot_data={"runtime": rt}, create_task=lambda coro, update=None: self.tasks.append(coro))
        self.ctx = NS(user_data={}, application=app, bot=NS(get_file=self._get_file))

    async def _get_file(self, file_id):
        return FakeFile()

    def update(self, message=None, query=None):
        msg = message or (query.message if query else None)
        return NS(effective_chat=NS(id=CHAT), message=message, effective_message=msg, callback_query=query)

    async def text(self, handler, text):
        return await handler(self.update(FakeMessage(self, text)), self.ctx)

    async def press(self, handler, data):
        q = FakeQuery(self, data)
        result = await handler(self.update(query=q), self.ctx)
        return q, result

    async def run_tasks(self):
        while self.tasks:
            await self.tasks.pop(0)

    @property
    def last(self) -> str:
        return self.sent[-1].text


class BotScenarioTest(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        coords = {f"ул. {i}": (51.10 + i / 200, 71.40 + (i % 3) / 100) for i in range(6)}
        self.geo = FakeGeocoder(coords)
        self.rt = Runtime(store=StateStore(root / "data"), geocoder=self.geo, planner=RoutePlanner(self.geo, FakeOsrm()))
        self.chat = Chat(self.rt)
        rows = [["№", "Название", "Адрес", "Контакт / телефон"]]
        rows += [[i + 1, f"Магазин {i}", f"ул. {i}", f"Агент +7 700 000 00 0{i}"] for i in range(6)]
        rows.append([7, "Потерянный", "неизвестная 1", ""])
        self.xlsx = make_xlsx(root / "in.xlsx", "Маршрут", rows)

    def tearDown(self):
        self.tmp.cleanup()

    def state(self):
        return self.rt.store.load(CHAT)

    async def upload(self):
        doc = NS(file_name="route.xlsx", get_file=self._file)
        await planning.on_document(self.chat.update(FakeMessage(self.chat, document=doc)), self.chat.ctx)
        await self.chat.run_tasks()

    async def _file(self):
        return FakeFile(self.xlsx)

    async def test_full_walk(self):
        await self.upload()
        s = self.state()
        self.assertEqual(len(s.points), 7)
        self.assertTrue(s.points[0].is_start)
        self.assertFalse(s.points[-1].has_coords)  # ненайденная точка — в конце, а не потеряна
        self.assertIn("Маршрут_Астана_готовый.xlsx", self.chat.documents)

        # Начали заметку по точке 2, передумали и отложили точку 3: диалог должен закрыться
        p2, p3 = s.by_order(2), s.by_order(3)
        _, state_ = await self.chat.press(visit.cb_visit_start, f"visit:{p2.id}")
        self.assertEqual(state_, visit.WAITING_NOTE)
        _, ended = await self.chat.press(visit.end_visit_then(walk.cb_skip), f"skip:{p3.id}")
        self.assertEqual(ended, ConversationHandler.END)
        self.assertNotIn(visit.VISIT_ID, self.chat.ctx.user_data)
        self.assertEqual(self.state().by_id(p3.id).status, Status.SKIPPED)

        # Записываем итог по точке 2 с категорией — после пересчёта маршрута (номера поменялись)
        await self.chat.press(visit.cb_visit_start, f"visit:{p2.id}")
        await self.chat.press(visit.cb_tag, "tag:price")
        loc = NS(latitude=51.125, longitude=71.42, horizontal_accuracy=10)
        await planning.on_location(self.chat.update(FakeMessage(self.chat, location=loc)), self.chat.ctx)
        await self.chat.run_tasks()
        self.assertIn("перестроен", "\n".join(m.text for m in self.chat.sent))
        result = await self.chat.text(visit.on_note_text, "сказали дорого")
        self.assertEqual(result, ConversationHandler.END)
        saved = self.state().by_id(p2.id)
        self.assertEqual((saved.status, saved.notes, saved.reason_tag), (Status.VISITED, "сказали дорого", "💸 Цена/Тариф"))

        # Отложенная точка после пересчёта стоит после всех «ожидающих»
        s = self.state()
        active = [p for p in s.points if p.status.is_active and p.has_coords]
        self.assertEqual(active[-1].id, p3.id)

        # Ошибочно закрыли → вернули
        await self.chat.press(walk.cb_closed, f"closed:{p3.id}")
        self.assertEqual(self.state().by_id(p3.id).status, Status.CLOSED)
        await self.chat.press(walk.cb_undo, f"undo:{p3.id}")
        self.assertEqual(self.state().by_id(p3.id).status, Status.PENDING)

        # Кнопка от удалённой точки
        q, _ = await self.chat.press(walk.cb_closed, "closed:99999")
        self.assertTrue(q.alerts)

    async def test_pin_sets_coordinates_and_remembers_address(self):
        await self.upload()
        lost = self.state().points[-1]
        await self.chat.press(planning.cb_pin, f"pin:{lost.id}")
        loc = NS(latitude=51.2, longitude=71.5, horizontal_accuracy=300)
        await planning.on_location(self.chat.update(FakeMessage(self.chat, location=loc)), self.chat.ctx)
        p = self.state().by_id(lost.id)
        self.assertEqual((p.lat, p.lon, p.src), (51.2, 71.5, "на месте"))
        self.assertEqual(self.geo.remembered["неизвестная 1"], (51.2, 71.5))
        self.assertIn("Точность GPS низкая", self.chat.last)
        self.assertEqual(self.chat.tasks, [])  # это была отметка точки, а не пересчёт маршрута

    async def test_reset_requires_confirmation(self):
        await self.upload()
        first = self.state().points[0]
        await self.chat.press(walk.cb_closed, f"closed:{first.id}")
        await self.chat.text(report.cmd_reset, "🔄 Сбросить прогресс")
        self.assertEqual(self.state().by_id(first.id).status, Status.CLOSED)
        await self.chat.press(report.cb_reset, "reset:no")
        self.assertEqual(self.state().by_id(first.id).status, Status.CLOSED)
        await self.chat.press(report.cb_reset, "reset:yes")
        self.assertEqual(self.state().by_id(first.id).status, Status.PENDING)

    async def test_album_photos_attach_to_same_point(self):
        await self.upload()
        p = self.state().points[1]
        await self.chat.press(visit.cb_visit_start, f"visit:{p.id}")
        photo = [NS(file_id="f1")]
        first = FakeMessage(self.chat, photo=photo, media_group_id="album1", caption="витрина")
        await visit.on_note_photo(self.chat.update(first), self.chat.ctx)
        second = FakeMessage(self.chat, photo=photo, media_group_id="album1")
        await visit.on_album_photo(self.chat.update(second), self.chat.ctx)
        saved = self.state().by_id(p.id)
        self.assertEqual(len(saved.photos), 2)
        self.assertEqual(saved.notes, "витрина")

    async def test_rejects_xls_and_finishes_with_report(self):
        doc = NS(file_name="old.xls")
        await planning.on_document(self.chat.update(FakeMessage(self.chat, document=doc)), self.chat.ctx)
        self.assertIn(".xls не поддерживается", self.chat.last)

        await self.upload()
        for p in self.state().points:
            await self.chat.press(walk.cb_closed, f"closed:{p.id}")
        self.assertIn("Все точки маршрута пройдены", "\n".join(m.text for m in self.chat.sent[-3:]))
        self.assertIn("Маршрут_Астана_отчет.xlsx", self.chat.documents)


if __name__ == "__main__":
    unittest.main()
