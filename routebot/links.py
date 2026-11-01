"""Ссылки на карты и мессенджеры. Строятся на лету из координат — в состоянии не хранятся,
поэтому после смены формата ссылки старые маршруты не остаются со «сломанными» ссылками."""
from __future__ import annotations

import re
from collections.abc import Sequence
from urllib.parse import quote

from . import config
from .models import Point

GMAPS_MAX_STOPS = 10  # origin + 8 промежуточных + destination


def gmaps_point(lat: float, lon: float) -> str:
    return f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"


def gmaps_search(addr: str) -> str:
    return f"https://www.google.com/maps/search/?api=1&query={quote(f'{config.CITY}, {addr}')}"


def dgis_point(lat: float, lon: float) -> str:
    """/geo/{долгота},{широта}: 2ГИС открывает карточку здания с пином в этой точке."""
    return f"https://2gis.kz/{config.DGIS_CITY}/geo/{lon:.6f}%2C{lat:.6f}"


def dgis_search(addr: str) -> str:
    return f"https://2gis.kz/{config.DGIS_CITY}/search/{quote(addr)}"


def point_links(p: Point) -> tuple[str, str]:
    """(Google Maps, 2ГИС): пин по координатам, а без координат — поиск по адресу."""
    if p.has_coords:
        return gmaps_point(p.lat, p.lon), dgis_point(p.lat, p.lon)
    return gmaps_search(p.addr), dgis_search(p.addr)


def gmaps_directions(points: Sequence[Point], from_current_location: bool = False) -> str:
    """
    Маршрут через несколько точек в Google Maps.
    from_current_location=True — без origin: навигатор поведёт от текущей позиции телефона.
    """
    pts = [p for p in points if p.has_coords]
    if not pts:
        raise ValueError("Нет точек с координатами")
    pts = pts[: GMAPS_MAX_STOPS - 1] if from_current_location else pts[:GMAPS_MAX_STOPS]
    fmt = lambda p: f"{p.lat},{p.lon}"  # noqa: E731
    origin = "" if from_current_location else f"&origin={fmt(pts[0])}"
    middle = pts[:-1] if from_current_location else pts[1:-1]
    waypoints = quote("|".join(fmt(p) for p in middle), safe=",")
    return (
        "https://www.google.com/maps/dir/?api=1&travelmode=driving"
        f"{origin}&destination={fmt(pts[-1])}"
        + (f"&waypoints={waypoints}" if waypoints else "")
    )


def whatsapp(phone: str) -> str | None:
    digits = re.sub(r"\D", "", phone or "")
    if digits.startswith("8") and len(digits) == 11:
        digits = "7" + digits[1:]
    return f"https://wa.me/{digits}" if len(digits) >= 10 else None
