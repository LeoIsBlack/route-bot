#!/usr/bin/env python3
"""
Маршрут по адресам Астаны из Excel (без бота).

    python route_planner.py                  # Маршрут_Астана.xlsx → Маршрут_Астана_готовый.xlsx
    python route_planner.py --src a.xlsx --out b.xlsx

Код — в пакете routebot/ (planner.py, solver.py, excel.py, …).
"""
import sys

from routebot.cli import main

if __name__ == "__main__":
    sys.exit(main())
