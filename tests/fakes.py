"""Тестовые двойники: геокодер, OSRM и HTTP без сети."""
from __future__ import annotations

import math
import random

from routebot.models import Point


def random_matrix(n: int, seed: int) -> list[list[float]]:
    """Асимметричная «дорожная» матрица: манхэттен + случайная добавка."""
    r = random.Random(seed)
    pts = [(r.random(), r.random()) for _ in range(n)]
    return [
        [0.0 if i == j else abs(a[0] - b[0]) * 900 + abs(a[1] - b[1]) * 700 + r.random() * 60 for j, b in enumerate(pts)]
        for i, a in enumerate(pts)
    ]


def travel_seconds(a: tuple[float, float], b: tuple[float, float]) -> float:
    return round(math.dist(a, b) * 100_000 / 10, 3)  # ~10 м/с


class FakeGeocoder:
    def __init__(self, coords: dict[str, tuple[float, float]]):
        self.coords = coords
        self.remembered: dict[str, tuple[float, float]] = {}

    def reload_manual(self) -> None:
        pass

    def locate(self, p: Point) -> bool:
        if p.addr in self.coords:
            (p.lat, p.lon), p.src = self.coords[p.addr], "тест"
            return True
        p.fail_reason = "нет в тестовых данных"
        return False

    def remember(self, addr, lat, lon) -> None:
        self.remembered[addr] = (lat, lon)


class FakeOsrm:
    """Время = расстояние; пары из blocked считаются «дороги нет»."""

    def __init__(self, blocked: set[tuple[float, float]] = frozenset()):
        self.blocked = blocked

    def durations(self, coords):
        return [
            [
                None if (a in self.blocked or b in self.blocked) and a != b else travel_seconds(a, b)
                for b in coords
            ]
            for a in coords
        ]


class FakeResponse:
    def __init__(self, status: int, payload: dict):
        self.status_code = status
        self._payload = payload

    def json(self):
        return self._payload


class FakeOsrmHttp:
    """Имитирует OSRM Table API (включая sources/destinations) и считает запросы."""

    def __init__(self, fail_first: int = 0):
        self.calls = 0
        self.fail_first = fail_first

    def get(self, url, params=None, timeout=None):
        self.calls += 1
        if self.calls <= self.fail_first:
            return FakeResponse(503, {})
        coords = [tuple(map(float, c.split(",")))[::-1] for c in url.rsplit("/", 1)[1].split(";")]
        src = [int(x) for x in params["sources"].split(";")] if "sources" in params else range(len(coords))
        dst = [int(x) for x in params["destinations"].split(";")] if "destinations" in params else range(len(coords))
        assert len(coords) <= 100, "публичный OSRM не принимает больше 100 точек"
        return FakeResponse(
            200, {"code": "Ok", "durations": [[travel_seconds(coords[i], coords[j]) for j in dst] for i in src]}
        )
