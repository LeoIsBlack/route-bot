"""CLI: python route_planner.py [--src входной.xlsx] [--out готовый.xlsx]"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from . import config
from .excel import InputError, read_input, write_route
from .geocoding import Geocoder
from .osrm import OsrmClient, OsrmError
from .planner import PlanError, RoutePlanner


def main(argv: list[str] | None = None) -> int:
    # В pipe (например, из другого процесса) Windows отдаёт cp1251 — эмодзи в логе уронили бы скрипт
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="Оптимальный маршрут объезда точек по Астане")
    ap.add_argument("--src", type=Path, default=config.SRC_XLSX, help="входная таблица (первая строка — СТАРТ)")
    ap.add_argument("--out", type=Path, default=config.OUT_XLSX, help="куда сохранить готовый маршрут")
    ap.add_argument("-q", "--quiet", action="store_true", help="меньше вывода")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(message)s",
        stream=sys.stdout,
    )
    if not args.src.exists():
        print(f"Положите {args.src.name} рядом со скриптом или укажите --src.")
        return 1

    try:
        points = read_input(args.src)
        print(f"Точек: {len(points)}. Ищу координаты…")
        planner = RoutePlanner(Geocoder(), OsrmClient())
        result = planner.plan(points)
    except (InputError, PlanError, OsrmError) as e:
        print(f"❌ {e}")
        return 1

    saved = write_route(args.out, result.route, result.missing)
    print(
        f"Готово: {saved.name}. В маршруте {len(result.route)} точек, "
        f"~{round(result.drive_minutes)} мин за рулём. Не найдено адресов: {len(result.missing)}"
    )
    return 0
