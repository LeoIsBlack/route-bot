"""Доменная модель: точка маршрута и состояние обхода.

Вся логика переходов (следующая точка, слияние прогресса при новом плане, пересчёт от GPS)
живёт здесь, а не в хэндлерах бота — её легко тестировать без Telegram.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime
from enum import StrEnum

TIME_FMT = "%Y-%m-%d %H:%M"
STATE_VERSION = 3


class Status(StrEnum):
    PENDING = "pending"
    VISITED = "visited"
    SKIPPED = "skipped"
    CLOSED = "closed"

    @property
    def emoji(self) -> str:
        return _STATUS_EMOJI[self]

    @property
    def label(self) -> str:
        return _STATUS_LABEL[self]

    @property
    def is_active(self) -> bool:
        """Точку ещё предстоит посетить."""
        return self in (Status.PENDING, Status.SKIPPED)


_STATUS_EMOJI = {
    Status.PENDING: "⏳",
    Status.VISITED: "✅",
    Status.SKIPPED: "⏭",
    Status.CLOSED: "🔒",
}
_STATUS_LABEL = {
    Status.PENDING: "Ожидает",
    Status.VISITED: "Посетил",
    Status.SKIPPED: "Отложено",
    Status.CLOSED: "Закрыто",
}


def now_str() -> str:
    return datetime.now().strftime(TIME_FMT)


@dataclass
class Point:
    # ── Из входной таблицы ──
    name: str = ""
    addr: str = ""
    num: int = 0  # № строки во входной таблице (ключ для manual_coords.csv)
    zone: str = ""
    owner: str = ""
    contact: str = ""
    phone: str = ""
    remark: str = ""  # «Примечание» из таблицы — подсказка агенту
    # ── Геокодирование ──
    lat: float | None = None
    lon: float | None = None
    src: str = ""  # откуда координаты: кэш / OSM / 2ГИС / вручную / на месте
    fail_reason: str = ""
    # ── Маршрут ──
    id: int = 0  # постоянный id: кнопки ссылаются на него, а не на порядковый номер
    order: int = 0
    mins: float = 0.0  # минут езды от предыдущей точки (или от GPS после пересчёта)
    is_start: bool = False
    # ── Прогресс обхода ──
    status: Status = Status.PENDING
    visit_time: str = ""
    reason_tag: str = ""
    notes: str = ""
    photos: list[str] = field(default_factory=list)

    @property
    def has_coords(self) -> bool:
        return self.lat is not None and self.lon is not None

    @property
    def xy(self) -> tuple[float, float]:
        if not self.has_coords:
            raise ValueError(f"У точки «{self.name}» нет координат")
        return self.lat, self.lon  # type: ignore[return-value]

    @property
    def uid(self) -> str:
        """Ключ магазина между перерасчётами маршрута (порядок и id могут поменяться)."""
        return f"{self.name.strip()}|{self.addr.strip()}"

    def add_note(self, text: str) -> None:
        text = (text or "").strip()
        if text:
            self.notes = f"{self.notes}; {text}" if self.notes else text

    def add_photo(self, filename: str) -> None:
        if filename:
            self.photos.append(filename)

    def copy_progress_from(self, other: Point) -> None:
        self.status = other.status
        self.visit_time = other.visit_time
        self.reason_tag = other.reason_tag
        self.notes = other.notes
        self.photos = list(other.photos)

    def reset_progress(self) -> None:
        self.copy_progress_from(Point())

    def to_dict(self) -> dict:
        d = asdict(self)
        d["status"] = self.status.value
        return d

    @classmethod
    def from_dict(cls, d: dict) -> Point:
        known = {f.name for f in fields(cls)}
        data = {k: v for k, v in d.items() if k in known and v is not None}
        try:
            data["status"] = Status(d.get("status") or Status.PENDING)
        except ValueError:
            data["status"] = Status.PENDING
        # Формат до 3.0: фото одной строкой «a.jpg, b.jpg», старт = первая по порядку
        if "photos" not in d and d.get("photo"):
            data["photos"] = [s.strip() for s in str(d["photo"]).split(",") if s.strip()]
        if "is_start" not in d:
            data["is_start"] = d.get("order") == 1 and not d.get("mins")
        if not (d.get("lat") and d.get("lon")):
            data["lat"] = data["lon"] = None
        data["id"] = int(d.get("id") or 0)
        data["mins"] = float(d.get("mins") or 0)
        return cls(**data)


@dataclass
class RouteState:
    points: list[Point] = field(default_factory=list)
    city: str = "Астана"
    planned_at: str = ""

    # ── Поиск ──
    def by_id(self, pid: int | None) -> Point | None:
        return next((p for p in self.points if p.id == pid), None)

    def by_order(self, order: int | None) -> Point | None:
        return next((p for p in self.points if p.order == order), None)

    def remaining(self) -> list[Point]:
        return [p for p in self.points if p.status.is_active]

    def counts(self) -> dict[Status, int]:
        c = {s: 0 for s in Status}
        for p in self.points:
            c[p.status] += 1
        return c

    def next_to_visit(self, after_order: int | None = None) -> tuple[Point | None, bool]:
        """
        Следующая точка обхода → (точка, вернулись_к_отложенной).

        Сначала «Ожидает» после after_order (агент идёт вперёд по маршруту),
        затем «Ожидает» с начала списка, и только потом — отложенные.
        """
        pending = [p for p in self.points if p.status is Status.PENDING]
        if after_order is not None:
            ahead = next((p for p in pending if p.order > after_order), None)
            if ahead:
                return ahead, False
        if pending:
            return pending[0], False
        skipped = next((p for p in self.points if p.status is Status.SKIPPED), None)
        return skipped, skipped is not None

    def cumulative_minutes(self) -> dict[int, float]:
        """id → минут за рулём от начала маршрута до этой точки."""
        total, out = 0.0, {}
        for p in self.points:
            if p.has_coords:
                total += p.mins
            out[p.id] = total
        return out

    def drive_minutes(self) -> float:
        return sum(p.mins for p in self.points if p.has_coords)

    # ── Изменения ──
    def renumber(self) -> None:
        for i, p in enumerate(self.points, 1):
            p.order = i

    def ensure_ids(self) -> None:
        """Выдаёт id точкам без id и дубликатам."""
        next_id = max((p.id for p in self.points), default=0) + 1
        seen: set[int] = set()
        for p in self.points:
            if not p.id or p.id in seen:
                p.id = next_id
                next_id += 1
            seen.add(p.id)

    def merge_progress(self, old: RouteState) -> None:
        """
        Переносит id, статусы, заметки и фото из прошлого состояния по ключу «имя|адрес»,
        чтобы повторная загрузка таблицы не стирала уже сделанную работу.
        """
        old_by_uid: dict[str, list[Point]] = defaultdict(list)
        for p in old.points:
            old_by_uid[p.uid].append(p)
        for p in self.points:
            candidates = old_by_uid.get(p.uid)
            if candidates:
                o = candidates.pop(0)
                p.id = o.id
                p.copy_progress_from(o)
        self.ensure_ids()

    def apply_reroute(self, sequence: list[tuple[int, float]]) -> list[Point]:
        """
        Применяет новый порядок оставшихся точек [(id, минут от предыдущей), …].

        Пройденные остаются в начале, затем новый порядок, затем активные точки,
        которых нет в sequence (без координат или недостижимые) — чтобы ничего не терялось.
        Точки, которые успели отметить, пока шёл расчёт, не трогаются.
        """
        by_id = {p.id: p for p in self.points}
        done = [p for p in self.points if not p.status.is_active]
        rerouted: list[Point] = []
        ids: set[int] = set()
        for pid, mins in sequence:
            p = by_id.get(pid)
            if p is None or not p.status.is_active or pid in ids:
                continue
            p.mins, p.is_start = mins, False
            rerouted.append(p)
            ids.add(pid)
        rest = [p for p in self.points if p.status.is_active and p.id not in ids]
        self.points = done + rerouted + rest
        self.renumber()
        return rerouted

    def reset_progress(self) -> None:
        for p in self.points:
            p.reset_progress()

    # ── Сериализация ──
    def to_dict(self) -> dict:
        return {
            "version": STATE_VERSION,
            "city": self.city,
            "planned_at": self.planned_at,
            "points": [p.to_dict() for p in self.points],
        }

    @classmethod
    def from_dict(cls, d: dict) -> RouteState:
        state = cls(
            points=[Point.from_dict(p) for p in d.get("points", []) if isinstance(p, dict)],
            city=d.get("city") or "Астана",
            planned_at=d.get("planned_at", ""),
        )
        state.ensure_ids()
        return state
