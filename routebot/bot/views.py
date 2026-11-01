"""Тексты сообщений бота (parse_mode=HTML). Чистые функции: состояние → строка."""
from __future__ import annotations

import html
from datetime import datetime, timedelta

from .. import config, links
from ..models import Point, RouteState, Status
from ..planner import PlanResult

NOTE_PREVIEW = 1000  # длинные заметки копятся через «Дополнить» — карточка не должна превысить 4096
PAGE_LIMIT = 3800

BACK_TO_SKIPPED = "🔄 <b>Возвращаемся к отложенной точке:</b>\n\n"
NO_ROUTE = "📤 Сначала отправьте .xlsx-файл со списком точек — я рассчитаю маршрут."

HELP = (
    "📌 <b>Как работать:</b>\n"
    "1️⃣ <b>Отправьте .xlsx</b> со списком магазинов (первая строка — ваш СТАРТ). "
    "Нужны колонки «№» и «Адрес»; «Название», «Контакт», «Телефон», «Примечание» — по желанию.\n"
    "2️⃣ <b>«▶️ Начать / Продолжить обход»</b> — карточка точки с WhatsApp, 2ГИС и Google Maps.\n"
    "3️⃣ <b>Фиксируйте итог</b> — категория + текст или фото (можно альбомом).\n"
    "4️⃣ <b>«📍 Пересчитать от меня»</b> — ближайшая точка станет следующей, отложенные уйдут в конец.\n"
    "5️⃣ <b>«📌 Я тут»</b> на карточке — запомнить точные координаты магазина, если карта ошиблась "
    "или адрес не нашёлся.\n"
    "6️⃣ <b>«📊 Скачать Excel-отчет»</b> — все заметки, статусы и аналитика.\n\n"
    "Команды: /walk /route /status /report /cancel /help"
)


def esc(value: object) -> str:
    return html.escape(str(value if value is not None else ""), quote=False)


def a(url: str, text: str) -> str:
    return f'<a href="{html.escape(url, quote=True)}">{esc(text)}</a>'


def _short(s: str, limit: int) -> str:
    return s if len(s) <= limit else s[: limit - 1] + "…"


def _timing(p: Point, total: float) -> str:
    if not p.has_coords:
        return "⚠️ Адрес не найден на карте — откройте поиск в 2ГИС, а на месте нажмите «📌 Я тут»"
    if p.is_start:
        return "🏁 СТАРТ"
    if p.mins < 1:
        return f"⏱ 0 мин — в том же месте (всего в пути {round(total)} мин)"
    return f"⏱ +{round(p.mins)} мин езды (всего в пути {round(total)} мин)"


def card(state: RouteState, p: Point, prefix: str = "") -> str:
    totals = state.cumulative_minutes()
    lines = [f"<b>{p.order}.</b> {esc(p.name)} {p.status.emoji}", f"📍 {esc(p.addr)}"]

    nxt = state.by_order(p.order + 1)
    if p.has_coords and nxt and nxt.has_coords and nxt.mins < 1:
        lines.append(f"🏢 <b>В этом же здании/рядом:</b> {esc(nxt.name)}")
    if p.remark:
        lines.append(f"📝 {esc(p.remark)}")
    if p.owner:
        lines.append(f"👤 Владелец: {esc(p.owner)}")
    if p.contact:
        lines.append(f"🧑‍💼 Контакт: {esc(p.contact)}")
    if p.phone:
        wa = links.whatsapp(p.phone)
        lines.append(f"📞 {esc(p.phone)}" + (f" · {a(wa, '💬 WhatsApp')}" if wa else ""))

    lines.append(_timing(p, totals.get(p.id, 0)))
    gmaps, dgis = links.point_links(p)
    lines.append(f"🔗 {a(gmaps, 'Google Maps')} · {a(dgis, '2ГИС')}")

    if p.reason_tag:
        lines.append(f"\n🏷 Категория: <b>{esc(p.reason_tag)}</b>")
    if p.notes:
        lines.append(f"💬 Заметка: {esc(_short(p.notes, NOTE_PREVIEW))}")
    if p.photos:
        lines.append(f"📸 Фото: {len(p.photos)} шт.")
    if p.visit_time:
        lines.append(f"🕐 Время визита: {esc(p.visit_time)}")
    return prefix + "\n".join(lines)


def progress(state: RouteState) -> str:
    c = state.counts()
    rem = state.remaining()
    text = (
        f"📊 <b>Прогресс:</b> {c[Status.VISITED] + c[Status.CLOSED]} из {len(state.points)} закрыто\n"
        f"   ✅ Посещено: {c[Status.VISITED]} | ⏭ Отложено: {c[Status.SKIPPED]} | "
        f"🔒 Закрыто: {c[Status.CLOSED]} | ⏳ Ожидает: {c[Status.PENDING]}"
    )
    if rem:
        mins = round(sum(p.mins for p in rem if p.has_coords) + len(rem) * config.AVG_SERVICE_MINS)
        eta = datetime.now() + timedelta(minutes=mins)
        text += f"\n🏁 <b>Прогноз завершения ({len(rem)} точ.):</b> ≈ {eta:%H:%M} (~{mins} мин с визитами)"
    return text


def next_stops_link(state: RouteState) -> str:
    """Ссылка «вести от моего местоположения через следующие точки» или пустая строка."""
    rem = [p for p in state.remaining() if p.has_coords]
    if not rem:
        return ""
    n = min(len(rem), links.GMAPS_MAX_STOPS - 1)
    return "🧭 " + a(links.gmaps_directions(rem, from_current_location=True), f"Навигатор по следующим {n} точкам")


def route_pages(state: RouteState) -> list[str]:
    pages, buf = [], [f"🗺 <b>Маршрут · {esc(state.city)}</b>"]
    totals = state.cumulative_minutes()
    for p in state.points:
        gmaps, dgis = links.point_links(p)
        if not p.has_coords:
            time_s = "⚠️ не найден на карте"
        elif p.is_start:
            time_s = "🏁 СТАРТ"
        else:
            time_s = f"⏱ +{round(p.mins)} мин (всего {round(totals[p.id])})"
        block = (
            f"<b>{p.order}.</b> {esc(p.name)} {p.status.emoji}\n"
            f"   📍 {esc(p.addr)}\n"
            f"   {time_s}  {a(gmaps, 'GMaps')} · {a(dgis, '2ГИС')}"
        )
        if p.notes:
            block += f"\n   💬 {esc(_short(p.notes, 200))}"
        if sum(map(len, buf)) + len(block) > PAGE_LIMIT:
            pages.append("\n\n".join(buf))
            buf = ["🗺 <b>Продолжение маршрута</b>"]
        buf.append(block)

    footer = [f"⏱ <b>Чистое время за рулём: ≈ {round(state.drive_minutes())} мин</b>", progress(state)]
    link = next_stops_link(state)
    if link:
        footer.append(link)
    buf.append("\n".join(footer))
    pages.append("\n\n".join(buf))
    return pages


def plan_done(result: PlanResult) -> str:
    text = (
        f"✅ <b>Маршрут рассчитан!</b>\n"
        f"Точек в маршруте: <b>{len(result.route)}</b>, за рулём ≈ <b>{round(result.drive_minutes)} мин</b>"
    )
    if result.missing:
        names = "\n".join(f"• {esc(p.name)} — {esc(p.fail_reason)}" for p in result.missing[:10])
        more = f"\n… и ещё {len(result.missing) - 10}" if len(result.missing) > 10 else ""
        text += (
            f"\n\n⚠️ <b>Не найдены на карте ({len(result.missing)})</b> — они в конце обхода, "
            f"на месте нажмите «📌 Я тут»:\n{names}{more}"
        )
    return text


def all_done(state: RouteState) -> str:
    return "🎉 <b>Все точки маршрута пройдены!</b>\n\n" + progress(state)
