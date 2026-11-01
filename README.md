# Маршрут-бот по Астане

Строит самый быстрый маршрут объезда торговых точек по дорогам и ведёт агента по ним
в Telegram: карточки точек, заметки и фото, пересчёт от текущей позиции, Excel-отчёт.

## Запуск

```bat
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
run.bat                                   :: бот (git pull + запуск)
.venv\Scripts\python route_planner.py     :: только маршрут: Маршрут_Астана.xlsx → Маршрут_Астана_готовый.xlsx
.venv\Scripts\python -m unittest discover -s tests -t .
```

Переменные окружения (необязательные): `BOT_TOKEN`, `DGIS_KEY` (поиск адресов через 2ГИС),
`OSRM_URL`, `AVG_SERVICE_MINS` (минут на точку для прогноза), `ROUTEBOT_DATA_DIR`.

## Входная таблица

Первая строка с данными — СТАРТ. Колонки ищутся по заголовкам: обязательны «№» и «Адрес»,
остальные — «Название/Точка», «Контакт», «Телефон» (или «Контакт / телефон» в одной ячейке),
«Владелец», «Зона», «Примечание». Можно загрузить обратно готовый маршрут или отчёт.

## Структура

```
routebot/
  config.py      пути, сервисы, константы
  models.py      Point, Status, RouteState — вся логика обхода (следующая точка, слияние, пересчёт)
  storage.py     атомарная запись JSON, состояние по чатам data/<chat_id>/
  addresses.py   нормализация адресов и варианты запросов
  geocoding.py   ручные координаты → кэш → 2ГИС → OSM
  osrm.py        матрица времени в пути (блоками, с повторами)
  solver.py      задача коммивояжёра: Held-Karp ≤12 точек, иначе ILS
  planner.py     геокодирование → матрица → порядок; пересчёт от GPS
  links.py       ссылки Google Maps / 2ГИС / WhatsApp
  excel.py       чтение входной таблицы, готовый маршрут, отчёт
  cli.py         python route_planner.py
  bot/
    app.py       сборка приложения и хэндлеров
    keyboards.py кнопки и callback_data
    views.py     тексты сообщений (HTML)
    common.py    общие помощники хэндлеров
    walk.py      обход и навигация
    visit.py     диалог заметки (текст, фото, альбомы)
    planning.py  расчёт из xlsx, пересчёт от GPS, «📌 Я тут»
    report.py    отчёт и сброс
tests/           unittest без сети (фейковые OSRM / геокодер / Telegram)
```
