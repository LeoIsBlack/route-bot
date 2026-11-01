"""Матрица времени в пути по дорогам (OSRM Table API)."""
from __future__ import annotations

import logging
import time
from collections.abc import Sequence

import requests

from . import config, net

log = logging.getLogger(__name__)

Matrix = list[list[float | None]]


class OsrmError(RuntimeError):
    """Сервис маршрутов недоступен или отклонил запрос."""


class OsrmClient:
    """
    d[i][j] — секунды от i до j; None — дороги нет. Матрица асимметрична
    (односторонки, развороты), поэтому d[i][j] ≠ d[j][i].

    Публичный сервер OSRM принимает не больше 100 точек в запросе — большие списки
    считаются блоками, и для вызывающего кода это прозрачно.
    """

    def __init__(
        self,
        base_url: str = config.OSRM_URL,
        *,
        profile: str = "driving",
        session: requests.Session | None = None,
        retries: int = 3,
        timeout: float = 60,
        max_coords: int = 100,
        retry_delay: float = 2.0,
    ):
        self.base_url = base_url.rstrip("/")
        self.profile = profile
        self._session = session
        self.retries = retries
        self.timeout = timeout
        self.max_coords = max_coords
        self.retry_delay = retry_delay

    @property
    def session(self) -> requests.Session:
        return self._session or net.session()

    def durations(self, coords: Sequence[tuple[float, float]]) -> Matrix:
        n = len(coords)
        if n == 0:
            return []
        if n == 1:
            return [[0.0]]
        if n <= self.max_coords:
            return self._table(coords)

        block = self.max_coords // 2
        d: Matrix = [[None] * n for _ in range(n)]
        for i0 in range(0, n, block):
            rows = range(i0, min(i0 + block, n))
            for j0 in range(0, n, block):
                cols = range(j0, min(j0 + block, n))
                idx = sorted(set(rows) | set(cols))
                pos = {g: k for k, g in enumerate(idx)}
                part = self._table(
                    [coords[g] for g in idx],
                    sources=[pos[i] for i in rows],
                    destinations=[pos[j] for j in cols],
                )
                for a, i in enumerate(rows):
                    d[i][j0 : j0 + len(cols)] = part[a]
        return d

    def _table(self, coords, sources=None, destinations=None) -> Matrix:
        url = (
            f"{self.base_url}/table/v1/{self.profile}/"
            + ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
        )
        params = {"annotations": "duration"}
        if sources is not None:
            params["sources"] = ";".join(map(str, sources))
        if destinations is not None:
            params["destinations"] = ";".join(map(str, destinations))

        last_err: Exception | None = None
        for attempt in range(self.retries):
            try:
                r = self.session.get(url, params=params, timeout=self.timeout)
                if r.status_code == 429 or r.status_code >= 500:
                    raise OsrmError(f"HTTP {r.status_code}")
                data = r.json()
            except (requests.RequestException, ValueError, OsrmError) as e:
                last_err = e
                log.warning("OSRM: попытка %d/%d не удалась: %s", attempt + 1, self.retries, e)
                if attempt + 1 < self.retries:
                    time.sleep(self.retry_delay * (attempt + 1))
                continue
            if data.get("code") != "Ok":
                raise OsrmError(f"OSRM: {data.get('message') or data.get('code')}")
            return data["durations"]
        raise OsrmError(f"Сервис маршрутов OSRM недоступен: {last_err}")
