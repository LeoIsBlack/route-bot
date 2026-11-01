"""Настройки проекта: пути, внешние сервисы, бизнес-константы.

Всё, что зависит от окружения, читается из переменных окружения с разумными значениями по умолчанию.
"""
from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# ── Файлы CLI (python route_planner.py) ──────────────────
SRC_XLSX = ROOT / "Маршрут_Астана.xlsx"
OUT_XLSX = ROOT / "Маршрут_Астана_готовый.xlsx"

# ── Общие для CLI и бота ─────────────────────────────────
GEOCODE_CACHE = ROOT / "geocode_cache.json"
MANUAL_COORDS = ROOT / "manual_coords.csv"

# ── Бот: у каждого чата своя папка data/<chat_id>/ ───────
DATA_DIR = Path(os.environ.get("ROUTEBOT_DATA_DIR", ROOT / "data"))
LEGACY_STATE = ROOT / "visit_state.json"  # до версии 3.0 состояние было одно на всех

# ── Внешние сервисы ──────────────────────────────────────
USER_AGENT = "astana-route-planner/3.0 (field-crm)"
OSRM_URL = os.environ.get("OSRM_URL", "https://router.project-osrm.org").rstrip("/")
DGIS_KEY = os.environ.get("DGIS_KEY") or None
BOT_TOKEN = os.environ.get("BOT_TOKEN", "9093591208:AAHLuair6rfphdfsQKwM8RttPpF6qkAQW3s")

# ── Город ────────────────────────────────────────────────
CITY = "Астана"
DGIS_CITY = "astana"
# Рамка Астаны (lon_min, lon_max, lat_min, lat_max): точки вне её — ошибка геокодера
BBOX = (71.15, 71.75, 50.95, 51.40)
# Центр Астаны в OSM — ложное срабатывание на слова «Султан» / «Астана»
CITY_CENTER = (51.1282804, 71.4304708)

# ── Бизнес-константы ─────────────────────────────────────
AVG_SERVICE_MINS = int(os.environ.get("AVG_SERVICE_MINS", "12"))  # парковка + разговор в точке
