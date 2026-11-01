"""Нормализация адресов Астаны и генерация вариантов поискового запроса для геокодеров.

Чистые функции без сети — их удобно покрывать тестами на реальных адресах из таблиц.
"""
from __future__ import annotations

import re

STREET_ALIASES = {
    "Шоссе Коргалжын": "Коргалжынское шоссе",
    "шоссе Коргалжын": "Коргалжынское шоссе",
    "Шоссе Алаш": "шоссе Алаш",
    "ул. Шакарим Кудайберди": "проспект Шакарима Кудайбердыулы",
    "ул. А. Байтұрсынов": "улица Ахмета Байтурсынова",
    "ул. Байтурсынова": "улица Ахмета Байтурсынова",
    "ул. Мухамедханов": "улица Каюма Мухамедханова",
    "ул. Аль-Фараби": "проспект Аль-Фараби",
}

LETTER_ROAD = re.compile(r"^(?:ул\.\s*)?([АЕA-E])\s+(\d+)", re.IGNORECASE)
NUM_STREET = re.compile(r"(?:ул\.\s*)?(\d{1,3}[-–]\d{1,3})")
PHONE = re.compile(r"\+?\d[\d\s()\-]{8,}\d")
_CITY_PREFIX = re.compile(r"^\s*(?:г\.?\s*|город\s+)?(?:Астана|Нур-Султан)\s*,\s*", re.IGNORECASE)
# Ориентиры после адреса: «…, ЖК Сатурн», «… Рынок Шарын», «…, бутик А6» — геокодеры их не знают
_LANDMARK = re.compile(r"(?:,\s*|\s+)(?:ЖК|НП|ВП|ТД|ТЦ|ТРЦ|бутик|рынок)\b.*$", re.IGNORECASE)
_DISTRICT = re.compile(r"((?:Алматинский|Есильский|Сарыаркинский|Байконурский|Нура)\s+район)\s*,?\s*", re.IGNORECASE)
_INVISIBLE = "⁠​ "


def clean_text(value: object) -> str:
    s = str(value if value is not None else "")
    for ch in _INVISIBLE:
        s = s.replace(ch, " " if ch == " " else "")
    return re.sub(r"\s+", " ", s).strip()


def clean_addr(addr: str) -> str:
    """«г.Астана, ул.Толе би 46» → «ул. Толе би 46» (так же, как ключи кэша геокодера)."""
    s = _CITY_PREFIX.sub("", clean_text(addr))
    s = re.sub(r"\b(ул|пр|мкр)\.(?=\S)", r"\1. ", s)
    return re.sub(r"\s+", " ", s).strip()


def split_contact(contact: str, phone: str) -> tuple[str, str]:
    """«Жамал Манас +7 702 222 0723» в одной ячейке → («Жамал Манас», «+7 702 222 0723»)."""
    if phone or not contact:
        return contact, phone
    m = PHONE.search(contact)
    if not m:
        return contact, phone
    rest = (contact[: m.start()] + contact[m.end() :]).strip(" ,;/")
    return rest, m.group(0).strip()


def normalize_street(addr: str) -> str:
    """Приводит адрес к форме, которую понимают геокодеры."""
    s = re.sub(r"\s+", " ", addr).strip()
    for raw, alias in STREET_ALIASES.items():
        if raw in s:
            s = s.replace(raw, alias)

    # «ул. А 52, дом 4» → «улица A-52, 4»
    numbered_road = False
    m = LETTER_ROAD.search(s)
    if m:
        letter = {"А": "A", "Е": "E"}.get(m.group(1).upper(), m.group(1).upper())
        s = s[: m.start()] + f"улица {letter}-{m.group(2)}" + s[m.end() :]
        numbered_road = True

    # «ул. 23-16, дом 39» → «улица 23-16, 39»
    m = NUM_STREET.search(s)
    if m:
        s = s[: m.start()] + f"улица {m.group(1).replace('–', '-')}" + s[m.end() :]
        numbered_road = True

    s = re.sub(r",?\s*\bдом\s+", ", ", s)  # «23-16 дом 39» и «23-16, дом 39» → «23-16, 39»
    # Висящий номер квартиры/офиса: «ул. Аль-Фараби 35, 14» → «ул. Аль-Фараби 35».
    # Но у «номерных» улиц (Е-652, 23-16) число до запятой — это название улицы,
    # а после запятой — номер дома, его отрезать нельзя.
    if not numbered_road:
        s = re.sub(r"^(.+?\d+[а-яА-Яa-zA-Z]?(?:/\d+)?),\s*\d+$", r"\1", s)
    return s.strip(" ,")


def address_variants(addr: str, name: str = "", allow_name_search: bool = False) -> list[str]:
    """Несколько вариантов запроса: от точного к общему. Без дубликатов, порядок важен."""
    raw = re.sub(r"\s+", " ", addr).strip()
    norm = normalize_street(raw)

    # Район вынимаем из адреса: он полезен только как уточнение к улице
    district = _DISTRICT.search(norm)
    street = _DISTRICT.sub("", norm).strip(" ,") if district else norm
    short = _LANDMARK.sub("", street).strip(" ,")
    nopref = re.sub(r"^(?:ул\.|пр\.|улица|проспект)\s*", "", short, flags=re.IGNORECASE)

    candidates = [norm, short]
    # Для номерных улиц «23-16, 39» без слова «улица» OSM находит «39-ю улицу» в другом
    # районе — ложный пин хуже честного «не найдено», поэтому такой вариант не пробуем
    if not re.match(r"^(?:[A-ZА-Я]-)?\d", nopref):
        candidates.append(nopref)

    # Дробь в номере дома (38/2) — пробуем и базовый номер
    slash = re.match(r"^(.+?)\s+(\d+)/\d+$", short)
    if slash:
        candidates.append(f"{slash.group(1)} {slash.group(2)}")

    if district and short:
        candidates.append(f"{district.group(1)}, {short}")

    jk = re.search(r"(?:ЖК|ТД)\s+([^\s,]+(?:\s*[-\d]+)?)", raw)
    if jk:
        candidates.append(f"ЖК {jk.group(1)}")

    # Поиск по названию магазина — только в 2ГИС: OSM путает «Султан» с городом Нур-Султан
    if allow_name_search and name:
        clean_name = re.sub(
            r"\s*(GS|бутик|маркет|market|СТАРТ|\(.*?\))\s*", " ", name, flags=re.IGNORECASE
        ).strip()
        if len(clean_name) > 2:
            candidates.append(clean_name)

    seen, out = set(), []
    for v in candidates:
        v = v.strip().rstrip(",")
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def guess_fail_reason(addr: str) -> str:
    if LETTER_ROAD.search(addr):
        return "Буквенно-цифровая улица (А-52, Е-652) — плохо покрыта в OSM"
    if NUM_STREET.search(addr):
        return "Числовая улица (23-16) — плохо покрыта в OSM"
    if re.search(r"Шоссе", addr, re.IGNORECASE):
        return "Шоссе — может быть записано иначе в OSM"
    if re.search(r",\s*\d+$", addr):
        return "Возможно некорректный номер дома или формат адреса"
    return "Адрес не распознан — откройте его в 2ГИС или отметьте точку кнопкой «📌 Я тут» на месте"
