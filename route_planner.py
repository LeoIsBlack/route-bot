#!/usr/bin/env python3
"""
Автоматический маршрут по адресам Астаны.

Что делает:
1. Читает «Маршрут_Астана.xlsx» (лист «Маршрут»), первая строка = СТАРТ.
2. Находит координаты каждого адреса (2ГИС, если задан ключ, иначе OpenStreetMap).
3. Считает время в пути между всеми точками по реальным дорогам (OSRM).
4. Подбирает порядок объезда с минимальным временем, старт фиксирован.
5. Сохраняет «Маршрут_Астана_готовый.xlsx» со ссылками Google Maps и 2ГИС.

Запуск:
    pip install requests openpyxl
    python route_planner.py

Необязательно: ключ 2ГИС (точнее ищет адреса в Казахстане)
    Windows:  set DGIS_KEY=ваш_ключ
    Mac/Linux: export DGIS_KEY=ваш_ключ

Если какой-то адрес не нашёлся, впишите его координаты в файл manual_coords.csv
строкой вида:  номер,широта,долгота   (например 7,51.1283,71.4304) и запустите снова.
"""
import csv
import json
import os
import re
import sys
import time
from urllib.parse import quote

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

SRC = "Маршрут_Астана.xlsx"
OUT = "Маршрут_Астана_готовый.xlsx"
CACHE = "geocode_cache.json"
MANUAL = "manual_coords.csv"
UA = {"User-Agent": "astana-route-planner/1.0 (personal use)"}
# Рамка Астаны: точки вне её считаем ошибкой поиска
LON_MIN, LON_MAX, LAT_MIN, LAT_MAX = 71.15, 71.75, 50.95, 51.40


def in_astana(lat, lon):
    return LAT_MIN <= lat <= LAT_MAX and LON_MIN <= lon <= LON_MAX


# ---------- чтение таблицы ----------
def read_points():
    ws = load_workbook(SRC)["Маршрут"]
    pts = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not isinstance(r[0], int):
            continue
        pts.append(dict(num=r[0], name=r[2], addr=r[3], owner=r[4] or "",
                        contact=r[5] or "", phone=r[6] or ""))
    return pts


# ---------- координаты ----------

# Нормализация казахских названий улиц для Nominatim/OSM
STREET_ALIASES = {
    "Шоссе Коргалжын": "Коргалжынское шоссе",
    "Шоссе Алаш": "шоссе Алаш",
    "Шоссе Коргалжын": "Коргалжынское шоссе",
}

# Шаблоны буквенно-цифровых улиц Астаны (А 52, Е 652 и т.п.)
_LETTER_ROAD = re.compile(r"^(?:ул\.\s*)?([АЕA-E])\s+(\d+)", re.IGNORECASE)
# Шаблоны числовых улиц (23-16 и т.п.)
_NUM_STREET = re.compile(r"(?:ул\.\s*)?(\d{1,3}[-–]\d{1,3})")


def _normalize_street(addr: str) -> str:
    """Приводит адрес к более стандартной форме для поисковиков."""
    s = re.sub(r"\s+", " ", addr).strip()
    # «Шоссе Коргалжын 88» → «Коргалжынское шоссе 88»
    for raw, alias in STREET_ALIASES.items():
        if raw in s:
            s = s.replace(raw, alias)
    # «ул. А 52, дом 4» → «улица А-52, 4»
    m = _LETTER_ROAD.search(s)
    if m:
        letter = m.group(1).upper()
        # Кириллические А/Е → латинские для OSM
        tr = {"А": "A", "Е": "E"}
        letter = tr.get(letter, letter)
        road = f"улица {letter}-{m.group(2)}"
        s = s[:m.start()] + road + s[m.end():]
    # «ул. 23-16, дом 39» → «улица 23-16, 39»
    m = _NUM_STREET.search(s)
    if m:
        num_street = m.group(1).replace("–", "-")
        repl = f"улица {num_street}"
        s = s[:m.start()] + repl + s[m.end():]
    # убираем «дом »
    s = re.sub(r"\bдом\s+", "", s)
    return s


def address_variants(addr, name=""):
    """Генерирует несколько вариантов поискового запроса."""
    raw = re.sub(r"\s+", " ", addr).strip()
    norm = _normalize_street(raw)

    # Основной: нормализованный адрес
    # Без ЖК/НП/ВП/ТД/бутик — только улица и дом
    short = re.split(r",\s*(?:ЖК|НП|ВП|ТД|бутик|рынок|Алматинский)", norm)[0]
    short = re.sub(r"\s+(?:ЖК|НП|ВП)\b.*$", "", short)
    # Без префикса ул./пр.
    nopref = re.sub(r"^(?:ул\.|пр\.)\s*", "", short)
    # Извлекаем название ЖК/ТД из адреса
    jk = re.search(r"(?:ЖК|ТД)\s+([^\s,]+(?:\s*[-\d]+)?)", raw)
    jk_name = jk.group(1) if jk else ""
    # Район из адреса
    district = ""
    dm = re.search(r"(Алматинский\s+район)", raw)
    if dm:
        district = dm.group(1)

    seen, out = set(), []
    candidates = [
        norm,                       # полный нормализованный
        short,                      # до ЖК/НП...
        nopref,                     # без ул./пр.
    ]
    if district and nopref:
        candidates.append(f"{district}, {nopref}")
    if jk_name:
        candidates.append(f"ЖК {jk_name}")
    # Название магазина (без «GS», «бутик» и т.п.)
    if name:
        clean_name = re.sub(r"\s*(GS|бутик|маркет|market|СТАРТ|\(.*?\))\s*", " ", name,
                            flags=re.IGNORECASE).strip()
        if clean_name and len(clean_name) > 2:
            candidates.append(clean_name)

    for v in candidates:
        v = v.strip().rstrip(",")
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def geocode_2gis(q, key):
    r = requests.get("https://catalog.api.2gis.com/3.0/items/geocode",
                     params={"q": q, "fields": "items.point", "key": key},
                     headers=UA, timeout=20)
    for it in r.json().get("result", {}).get("items", []):
        p = it.get("point")
        if p and in_astana(p["lat"], p["lon"]):
            return p["lat"], p["lon"]
    return None


def geocode_osm(q, bounded=True):
    """Ищет координаты через Nominatim. bounded=False расширяет поиск."""
    time.sleep(1.1)  # правило Nominatim: не чаще 1 запроса в секунду
    params = {"q": q, "format": "json", "limit": 5, "countrycodes": "kz"}
    if bounded:
        params.update({
            "viewbox": f"{LON_MIN},{LAT_MAX},{LON_MAX},{LAT_MIN}",
            "bounded": 1,
        })
    r = requests.get("https://nominatim.openstreetmap.org/search",
                     params=params, headers=UA, timeout=20)
    for it in r.json():
        lat, lon = float(it["lat"]), float(it["lon"])
        if in_astana(lat, lon):
            return lat, lon
    return None


def _try_geocode(q, key):
    """Пробует 2ГИС (если есть ключ), затем OSM bounded, затем OSM без ограничения."""
    res = None
    try:
        if key:
            res = geocode_2gis(q, key)
        if not res:
            res = geocode_osm(q, bounded=True)
        if not res:
            res = geocode_osm(q, bounded=False)
    except Exception as e:
        print(f"  ⚠ ошибка запроса для «{q}»: {e}")
    return res


def load_json(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def find_coords(pts):
    cache = load_json(CACHE)
    # Сбрасываем кэш для ранее ненайденных адресов, чтобы попробовать снова
    for addr, val in list(cache.items()):
        if val is None:
            del cache[addr]
    key = os.environ.get("DGIS_KEY")
    manual = {}
    if os.path.exists(MANUAL):
        with open(MANUAL, encoding="utf-8") as f:
            for row in csv.reader(f):
                if len(row) >= 3 and row[0].strip().isdigit():
                    manual[int(row[0])] = (float(row[1]), float(row[2]))
    for p in pts:
        if p["num"] in manual:
            p["xy"], p["src"] = manual[p["num"]], "вручную"
            continue
        if p["addr"] in cache and cache[p["addr"]]:
            p["xy"], p["src"] = tuple(cache[p["addr"]]), "кэш"
            continue
        p["xy"], p["src"] = None, ""
        variants = address_variants(p["addr"], p.get("name", ""))
        for v in variants:
            q = "Астана, " + v
            print(f"  🔍 пробую: {q}")
            res = _try_geocode(q, key)
            if res:
                p["xy"], p["src"] = res, ("2ГИС" if key else "OSM")
                break
        if not p["xy"]:
            # Причина ненахождения
            reason = _guess_reason(p["addr"])
            p["fail_reason"] = reason
            print(f"  ❌ причина: {reason}")
        cache[p["addr"]] = list(p["xy"]) if p["xy"] else None
        status = "✅ " + str(p["xy"]) if p["xy"] else "❌ НЕ НАЙДЕН"
        print(f"{p['num']:>2}. {p['name']}: {status}")
        with open(CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    return pts


def _guess_reason(addr):
    """Пытается определить причину, почему адрес не нашёлся."""
    if _LETTER_ROAD.search(addr):
        return "Буквенно-цифровая улица (А-52, Е-652) — плохо покрыта в OSM"
    if _NUM_STREET.search(addr):
        return "Числовая улица (23-16) — плохо покрыта в OSM"
    if re.search(r"Шоссе", addr, re.IGNORECASE):
        return "Шоссе — может быть записано иначе в OSM"
    if re.search(r",\s*\d+$", addr):
        return "Возможно некорректный номер дома или формат адреса"
    return "Адрес не распознан в OpenStreetMap — откройте в 2ГИС и впишите координаты"


# ---------- время в пути ----------
def duration_matrix(coords):
    s = ";".join(f"{lon},{lat}" for lat, lon in coords)
    r = requests.get(f"https://router.project-osrm.org/table/v1/driving/{s}",
                     params={"annotations": "duration"}, headers=UA, timeout=60)
    data = r.json()
    if data.get("code") != "Ok":
        raise RuntimeError(f"OSRM: {data.get('message', data.get('code'))}")
    return data["durations"]


# ---------- порядок объезда (открытый маршрут, старт = точка 0) ----------
def path_cost(order, d):
    return sum(d[a][b] for a, b in zip(order, order[1:]))


def solve(d):
    n = len(d)
    order, left = [0], set(range(1, n))
    while left:  # ближайший сосед
        nxt = min(left, key=lambda j: d[order[-1]][j])
        order.append(nxt)
        left.remove(nxt)
    improved = True
    while improved:  # 2-opt (старт не двигаем)
        improved = False
        for i in range(1, n - 1):
            for j in range(i + 1, n):
                new = order[:i] + order[i:j + 1][::-1] + order[j + 1:]
                if path_cost(new, d) < path_cost(order, d) - 1e-9:
                    order, improved = new, True
    improved = True
    while improved:  # перенос одной точки
        improved = False
        for i in range(1, n):
            node = order[i]
            rest = order[:i] + order[i + 1:]
            for k in range(1, n):
                new = rest[:k] + [node] + rest[k:]
                if path_cost(new, d) < path_cost(order, d) - 1e-9:
                    order, improved = new, True
                    break
            if improved:
                break
    return order


# ---------- ссылки ----------
def gmaps_url(chunk):
    f = lambda p: f"{p['xy'][0]},{p['xy'][1]}"
    wp = quote("|".join(f(p) for p in chunk[1:-1]), safe=",")
    return ("https://www.google.com/maps/dir/?api=1&travelmode=driving"
            f"&origin={f(chunk[0])}&destination={f(chunk[-1])}"
            + (f"&waypoints={wp}" if wp else ""))


def dgis_url(p):
    return f"https://2gis.kz/astana?m={p['xy'][1]},{p['xy'][0]}/17"


# ---------- сохранение ----------
def save(ordered, minutes, missing):
    F = "Arial"
    hf = PatternFill("solid", fgColor="1F3A5F")
    wb = Workbook()
    ws = wb.active
    ws.title = "Маршрут"
    ws.append(["Порядок", "Точка", "Адрес", "Владелец", "Контакт", "Телефон",
               "Минут от предыдущей", "Всего минут", "Откуда координаты", "Google Maps", "2ГИС"])
    total = 0.0
    for i, (p, m) in enumerate(zip(ordered, minutes), 1):
        total += m
        ws.append([i, p["name"], p["addr"], p["owner"], p["contact"], p["phone"],
                   round(m), round(total), p["src"], "Открыть", "Открыть"])
        r = ws.max_row
        ws.cell(r, 10).hyperlink = "https://www.google.com/maps/search/?api=1&query=" + \
            f"{p['xy'][0]},{p['xy'][1]}"
        ws.cell(r, 11).hyperlink = dgis_url(p)
        for c in (10, 11):
            ws.cell(r, c).font = Font(name=F, color="0563C1", underline="single")
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = hf
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_col=9):
        for c in row:
            c.font = Font(name=F, size=11)
            c.alignment = Alignment(vertical="center", wrap_text=True)
    for k, w in zip("ABCDEFGHIJK", [8, 24, 38, 26, 24, 18, 12, 11, 14, 13, 8]):
        ws.column_dimensions[k].width = w
    ws.freeze_panes = "C2"
    n = ws.max_row + 2
    ws.cell(n, 1, f"Время в пути по дорогам без учёта пробок и стоянок: около {round(total)} мин.").font = Font(name=F, size=10)

    gs = wb.create_sheet("Ссылки Google Maps")
    gs.append(["Часть", "Откуда", "Куда", "Точек", "Ссылка"])
    i = k = 0
    while i < len(ordered) - 1:
        chunk = ordered[i:i + 10]
        k += 1
        gs.append([k, chunk[0]["name"], chunk[-1]["name"], len(chunk), "Открыть маршрут"])
        gs.cell(gs.max_row, 5).hyperlink = gmaps_url(chunk)
        gs.cell(gs.max_row, 5).font = Font(name=F, color="0563C1", underline="single")
        i += 9
    for c in gs[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = hf
    for kk, w in zip("ABCDE", [8, 26, 26, 8, 20]):
        gs.column_dimensions[kk].width = w

    if missing:
        ms = wb.create_sheet("Не найдены")
        ms.append(["№ в исходной таблице", "Точка", "Адрес", "Причина", "Что делать"])
        for p in missing:
            reason = p.get("fail_reason", "Не распознан")
            ms.append([p["num"], p["name"], p["addr"], reason,
                       f"Откройте адрес в 2ГИС, скопируйте координаты → {MANUAL}: {p['num']},широта,долгота"])
        for c in ms[1]:
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = hf
        for kk, w in zip("ABCDE", [12, 24, 40, 44, 60]):
            ms.column_dimensions[kk].width = w
    try:
        wb.save(OUT)
    except PermissionError:
        # Файл открыт в другой программе — сохраняем под другим именем
        base, ext = os.path.splitext(OUT)
        for i in range(1, 100):
            alt = f"{base}_{i}{ext}"
            try:
                wb.save(alt)
                print(f"⚠ Файл {OUT} занят, сохранён как {alt}")
                return alt
            except PermissionError:
                continue
        sys.exit(f"Не удалось сохранить: {OUT} и все альтернативы заняты. Закройте файл.")
    return OUT


def main():
    if not os.path.exists(SRC):
        sys.exit(f"Положите {SRC} рядом со скриптом.")
    pts = read_points()
    print(f"Точек: {len(pts)}. Ищу координаты…")
    pts = find_coords(pts)
    if pts[0]["xy"] is None:
        sys.exit("Не найден СТАРТ (первая строка). Впишите его в manual_coords.csv.")
    found = [p for p in pts if p["xy"]]
    missing = [p for p in pts if not p["xy"]]
    print("Считаю время в пути по дорогам…")
    d = duration_matrix([p["xy"] for p in found])
    for row in d:
        if any(x is None for x in row):
            sys.exit("OSRM не нашёл дорогу до одной из точек. Проверьте координаты.")
    order = solve(d)
    ordered = [found[i] for i in order]
    minutes = [0.0] + [d[a][b] / 60 for a, b in zip(order, order[1:])]
    saved_as = save(ordered, minutes, missing)
    print(f"Готово: {saved_as}. Не найдено адресов: {len(missing)}")


if __name__ == "__main__":
    main()



def reroute_remaining(start_lat, start_lon, remaining_points):
    """
    Пересчитывает маршрут для оставшихся точек от текущей позиции.

    Args:
        start_lat, start_lon: текущие координаты пользователя.
        remaining_points: список dict с ключами order, name, addr, owner,
                          contact, phone, xy=(lat,lon), src, ...

    Returns:
        (ordered_points, minutes) — пересортированные точки и время между ними.
        Порядковые номера (order) пересчитываются с 1.
    """
    if not remaining_points:
        return [], []
    # Виртуальная стартовая точка + оставшиеся
    all_coords = [(start_lat, start_lon)] + [p["xy"] for p in remaining_points]
    d = duration_matrix(all_coords)
    for row in d:
        if any(x is None for x in row):
            raise RuntimeError("OSRM не нашёл дорогу до одной из точек.")
    order = solve(d)
    # order[0] == 0 (виртуальный старт), остальные — индексы в remaining_points
    ordered = []
    minutes = []
    for pos, idx in enumerate(order):
        if idx == 0:
            continue  # пропускаем виртуальный старт
        prev_idx = order[pos - 1] if pos > 0 else 0
        m = d[prev_idx][idx] / 60
        ordered.append(remaining_points[idx - 1])
        minutes.append(m)
    # Пересчитываем order начиная с 1
    for i, p in enumerate(ordered, 1):
        p["order"] = i
    return ordered, minutes



def reroute_remaining(start_lat, start_lon, remaining_points):
    """
    Пересчитывает маршрут для оставшихся точек от текущей позиции.

    Args:
        start_lat, start_lon: текущие координаты пользователя.
        remaining_points: список dict с ключами order, name, addr, owner,
                          contact, phone, xy=(lat,lon), src, ...

    Returns:
        (ordered_points, minutes) — пересортированные точки и время между ними.
        Порядковые номера (order) пересчитываются с 1.
    """
    if not remaining_points:
        return [], []
    # Виртуальная стартовая точка + оставшиеся
    all_coords = [(start_lat, start_lon)] + [p["xy"] for p in remaining_points]
    d = duration_matrix(all_coords)
    for row in d:
        if any(x is None for x in row):
            raise RuntimeError("OSRM не нашёл дорогу до одной из точек.")
    order = solve(d)
    # order[0] == 0 (виртуальный старт), остальные — индексы в remaining_points
    ordered = []
    minutes = []
    for pos, idx in enumerate(order):
        if idx == 0:
            continue  # пропускаем виртуальный старт
        prev_idx = order[pos - 1] if pos > 0 else 0
        m = d[prev_idx][idx] / 60
        ordered.append(remaining_points[idx - 1])
        minutes.append(m)
    # Пересчитываем order начиная с 1
    for i, p in enumerate(ordered, 1):
        p["order"] = i
    return ordered, minutes
