"""Планирование маршрута: геокодирование → матрица времени → оптимальный порядок.

Не зависит ни от Telegram, ни от Excel: на входе точки, на выходе упорядоченный маршрут.
Прогресс сообщается через колбэк — CLI печатает его, бот показывает в сообщении.
"""
from __future__ import annotations

import logging
from collections.abc import Callable, Sequence
from dataclasses import dataclass

from .geocoding import Geocoder
from .models import Point
from .osrm import Matrix, OsrmClient
from .solver import path_cost, solve, solve_subset

log = logging.getLogger(__name__)

Progress = Callable[[str], None]


class PlanError(Exception):
    """Ошибка, которую можно показать пользователю как есть."""


@dataclass
class PlanResult:
    route: list[Point]  # упорядоченные точки с координатами, route[0] — СТАРТ
    missing: list[Point]  # не найдены на карте или недостижимы по дорогам

    @property
    def points(self) -> list[Point]:
        return self.route + self.missing

    @property
    def drive_minutes(self) -> float:
        return sum(p.mins for p in self.route)


def _noop(_: str) -> None:
    pass


def reachable_indices(d: Matrix, keep: int = 0) -> list[int]:
    """
    Индексы, между которыми OSRM нашёл дороги во все стороны. Жадно выбрасываем точку
    с наибольшим числом «нет дороги», пока такие есть. Точку keep (СТАРТ / GPS)
    выбрасываем только если без неё никак — тогда PlanError.
    """
    idx = list(range(len(d)))
    while True:
        bad = {
            i: sum(1 for j in idx if j != i and (d[i][j] is None or d[j][i] is None)) for i in idx
        }
        worst = max(idx, key=lambda i: (bad[i], i != keep))
        if bad[worst] == 0:
            return idx
        if worst == keep:
            raise PlanError("Не удалось построить путь по дорогам от начальной точки — проверьте её координаты.")
        idx.remove(worst)


def _submatrix(d: Matrix, idx: Sequence[int]) -> list[list[float]]:
    return [[d[a][b] for b in idx] for a in idx]  # type: ignore[misc]


class RoutePlanner:
    def __init__(self, geocoder: Geocoder, osrm: OsrmClient):
        self.geocoder = geocoder
        self.osrm = osrm

    def plan(self, points: Sequence[Point], progress: Progress = _noop) -> PlanResult:
        """Первая точка списка — СТАРТ: она всегда первая в маршруте."""
        points = list(points)
        if not points:
            raise PlanError("В таблице нет ни одной точки.")

        self.geocoder.reload_manual()
        n = len(points)
        for i, p in enumerate(points, 1):
            progress(f"🔍 Ищу адреса на карте: {i}/{n}\n{p.name}")
            ok = self.geocoder.locate(p)
            log.info("%2d. %s: %s", p.num, p.name, f"✅ {p.xy}" if ok else f"❌ {p.fail_reason}")

        start = points[0]
        if not start.has_coords:
            raise PlanError(
                f"Не найден СТАРТ (первая строка) «{start.name}», адрес «{start.addr}». "
                "Уточните адрес или впишите координаты в manual_coords.csv."
            )
        found = [p for p in points if p.has_coords]
        missing = [p for p in points if not p.has_coords]

        progress(f"🚗 Считаю время в пути по дорогам между {len(found)} точками…")
        d = self.osrm.durations([p.xy for p in found])
        keep = reachable_indices(d)
        for i in sorted(set(range(len(found))) - set(keep)):
            found[i].fail_reason = "OSRM не нашёл дорогу до точки — проверьте координаты"
            missing.append(found[i])
        found, d = [found[i] for i in keep], _submatrix(d, keep)

        progress("🧠 Подбираю оптимальный порядок объезда…")
        order = solve(d)
        log.info("Оптимальный порядок: ~%d мин в пути.", round(path_cost(order, d) / 60))

        route = []
        for k, i in enumerate(order):
            p = found[i]
            p.mins = d[order[k - 1]][i] / 60 if k else 0.0
            p.is_start = k == 0
            route.append(p)
        for p in missing:
            p.mins, p.is_start = 0.0, False
        for k, p in enumerate(route + missing, 1):
            p.order = k
        return PlanResult(route=route, missing=missing)

    def reroute(
        self,
        lat: float,
        lon: float,
        pending: Sequence[Point],
        deferred: Sequence[Point] = (),
    ) -> list[tuple[int, float]]:
        """
        Новый порядок от GPS-позиции → [(id точки, минут от предыдущей), …].

        Первой идёт точка, до которой быстрее всего доехать по дорогам; остальные «ожидающие» —
        в оптимальном порядке после неё; отложенные (deferred) — в конце, тоже оптимально.
        Недостижимые точки в ответ не попадают (RouteState.apply_reroute оставит их в конце).
        """
        main, later = (list(pending), list(deferred)) if pending else (list(deferred), [])
        if not main:
            return []
        pts = main + later
        d = self.osrm.durations([(lat, lon)] + [p.xy for p in pts])
        keep = reachable_indices(d)  # 0 — позиция агента
        d = _submatrix(d, keep)
        # keep отсортирован, поэтому основные точки по-прежнему идут перед отложенными
        n_main = sum(1 for i in keep[1:] if i <= len(main)) or len(keep) - 1
        pts = [pts[i - 1] for i in keep[1:]]
        if not pts:
            return []

        main_idx = range(1, n_main + 1)
        later_idx = range(n_main + 1, len(d))
        closest = min(main_idx, key=lambda j: d[0][j])
        order = solve_subset(d, (0, closest), [j for j in main_idx if j != closest])
        if later_idx:
            order += solve_subset(d, (order[-1],), later_idx)[1:]
        return [(pts[i - 1].id, d[a][i] / 60) for a, i in zip(order, order[1:])]
