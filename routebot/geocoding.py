"""Геокодирование адресов: ручные координаты → кэш → 2ГИС (если есть ключ) → OpenStreetMap."""
from __future__ import annotations

import csv
import logging
import threading
import time
from pathlib import Path

import requests

from . import config, net
from .addresses import address_variants, guess_fail_reason
from .models import Point
from .storage import atomic_write_json, read_json

log = logging.getLogger(__name__)

LatLon = tuple[float, float]


def in_city(lat: float, lon: float) -> bool:
    lon_min, lon_max, lat_min, lat_max = config.BBOX
    return lat_min <= lat <= lat_max and lon_min <= lon <= lon_max


def is_false_city_center(lat: float, lon: float) -> bool:
    """Геокодер вернул абстрактный центр города вместо улицы."""
    c_lat, c_lon = config.CITY_CENTER
    return abs(lat - c_lat) < 0.0003 and abs(lon - c_lon) < 0.0003


def _plausible(lat: float, lon: float) -> bool:
    return in_city(lat, lon) and not is_false_city_center(lat, lon)


class Geocoder:
    DGIS_URL = "https://catalog.api.2gis.com/3.0/items/geocode"
    NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
    NOMINATIM_INTERVAL = 1.1  # политика Nominatim: не чаще 1 запроса в секунду

    def __init__(
        self,
        cache_path: Path = config.GEOCODE_CACHE,
        manual_path: Path | None = config.MANUAL_COORDS,
        dgis_key: str | None = config.DGIS_KEY,
        session: requests.Session | None = None,
    ):
        self.cache_path = Path(cache_path)
        self.manual_path = Path(manual_path) if manual_path else None
        self.dgis_key = dgis_key
        self._session = session
        self._cache_lock = threading.Lock()  # кэш пишут и поток расчёта, и бот («📌 Я тут»)
        self._rate_lock = threading.Lock()
        self._last_osm = 0.0
        self._cache = self._load_cache()
        self._manual: dict[int, LatLon] = {}
        self.reload_manual()

    @property
    def session(self) -> requests.Session:
        return self._session or net.session()

    # ── Публичный API ──
    def locate(self, p: Point) -> bool:
        """Заполняет p.lat/lon/src, либо p.fail_reason. Возвращает True, если нашли."""
        p.fail_reason = ""
        if p.num in self._manual:
            (p.lat, p.lon), p.src = self._manual[p.num], "вручную"
            return True
        cached = self._cache.get(p.addr)
        if cached:
            (p.lat, p.lon), p.src = cached, "кэш"
            return True
        if not p.addr:
            p.fail_reason = "В таблице не указан адрес"
            return False

        for variant in address_variants(p.addr, p.name, allow_name_search=bool(self.dgis_key)):
            query = f"{config.CITY}, {variant}"
            log.info("  🔍 пробую: %s", query)
            found = self._lookup(query)
            if found:
                (p.lat, p.lon), p.src = found
                self.remember(p.addr, p.lat, p.lon)
                return True

        p.fail_reason = guess_fail_reason(p.addr)
        return False

    def remember(self, addr: str, lat: float, lon: float) -> None:
        """Сохраняет координаты адреса (в т.ч. отмеченные агентом на месте) для будущих расчётов."""
        with self._cache_lock:
            self._cache[addr] = (lat, lon)
            atomic_write_json(self.cache_path, {k: list(v) for k, v in self._cache.items()})

    def reload_manual(self) -> None:
        """manual_coords.csv: «№ строки,широта,долгота». Перечитывается перед каждым расчётом."""
        manual: dict[int, LatLon] = {}
        if self.manual_path and self.manual_path.exists():
            with open(self.manual_path, encoding="utf-8") as f:
                for row in csv.reader(f):
                    try:
                        manual[int(row[0])] = (float(row[1]), float(row[2]))
                    except (ValueError, IndexError):
                        continue  # заголовок или кривая строка
        self._manual = manual

    # ── Внутреннее ──
    def _load_cache(self) -> dict[str, LatLon]:
        try:
            raw = read_json(self.cache_path, {})
        except ValueError:
            log.error("Кэш геокодера %s повреждён — начинаю с пустого", self.cache_path)
            raw = {}
        cache = {}
        for addr, val in raw.items():
            # Пустые значения и старые ложные «центры города» выбрасываем
            if isinstance(val, list) and len(val) == 2 and _plausible(*val):
                cache[addr] = (float(val[0]), float(val[1]))
        return cache

    def _lookup(self, query: str) -> tuple[LatLon, str] | None:
        try:
            if self.dgis_key:
                res = self._dgis(query)
                if res:
                    return res, "2ГИС"
            # Отдельный поиск без рамки города не нужен: результаты вне рамки всё равно
            # отбрасываются, а это удваивало число запросов (и секунд ожидания) на адрес.
            res = self._nominatim(query)
            if res:
                return res, "OSM"
        except (requests.RequestException, ValueError) as e:
            log.warning("  ⚠ ошибка запроса для «%s»: %s", query, e)
        return None

    def _dgis(self, query: str) -> LatLon | None:
        r = self.session.get(
            self.DGIS_URL,
            params={"q": query, "fields": "items.point", "key": self.dgis_key},
            timeout=20,
        )
        for it in r.json().get("result", {}).get("items", []):
            pt = it.get("point")
            if pt and _plausible(pt["lat"], pt["lon"]):
                return pt["lat"], pt["lon"]
        return None

    def _nominatim(self, query: str) -> LatLon | None:
        self._throttle()
        lon_min, lon_max, lat_min, lat_max = config.BBOX
        r = self.session.get(
            self.NOMINATIM_URL,
            params={
                "q": query,
                "format": "json",
                "limit": 5,
                "countrycodes": "kz",
                "viewbox": f"{lon_min},{lat_max},{lon_max},{lat_min}",
                "bounded": 1,
            },
            timeout=20,
        )
        for it in r.json():
            if it.get("type") in ("administrative", "city", "state"):
                continue
            lat, lon = float(it["lat"]), float(it["lon"])
            if _plausible(lat, lon):
                return lat, lon
        return None

    def _throttle(self) -> None:
        with self._rate_lock:
            wait = self._last_osm + self.NOMINATIM_INTERVAL - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            self._last_osm = time.monotonic()
