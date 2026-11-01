"""Хранение данных: атомарная запись JSON и состояние обхода отдельно для каждого чата."""
from __future__ import annotations

import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import RouteState

log = logging.getLogger(__name__)


def read_json(path: Path, default: Any) -> Any:
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def atomic_write_json(path: Path, data: Any) -> None:
    """Пишет во временный файл и подменяет: сбой посреди записи не портит исходный файл."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


class StateStore:
    """
    Состояние обхода по чатам: data/<chat_id>/state.json (+ input/route/report xlsx и photos/).
    Раньше состояние было одно на всех — два агента затирали маршруты друг друга.
    """

    def __init__(self, data_dir: Path, legacy_state: Path | None = None):
        self.data_dir = Path(data_dir)
        self.legacy_state = legacy_state

    # ── Пути ──
    def chat_dir(self, chat_id: int) -> Path:
        d = self.data_dir / str(chat_id)
        d.mkdir(parents=True, exist_ok=True)
        return d

    def state_path(self, chat_id: int) -> Path:
        return self.chat_dir(chat_id) / "state.json"

    def input_path(self, chat_id: int) -> Path:
        return self.chat_dir(chat_id) / "input.xlsx"

    def route_path(self, chat_id: int) -> Path:
        return self.chat_dir(chat_id) / "route.xlsx"

    def report_path(self, chat_id: int) -> Path:
        return self.chat_dir(chat_id) / "report.xlsx"

    def photos_dir(self, chat_id: int) -> Path:
        d = self.chat_dir(chat_id) / "photos"
        d.mkdir(exist_ok=True)
        return d

    # ── Состояние ──
    def load(self, chat_id: int) -> RouteState:
        path = self.state_path(chat_id)
        if not path.exists() and self._migrate_legacy(path):
            log.info("Состояние %s перенесено в чат %s", self.legacy_state, chat_id)
        try:
            data = read_json(path, None)
        except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
            broken = path.with_name(f"state.broken-{datetime.now():%Y%m%d-%H%M%S}.json")
            log.error("Повреждён %s (%s) — сохранён как %s", path, e, broken.name)
            path.replace(broken)  # не даём следующему save затереть заметки
            return RouteState()
        if not isinstance(data, dict):
            return RouteState()
        return RouteState.from_dict(data)

    def save(self, chat_id: int, state: RouteState) -> None:
        atomic_write_json(self.state_path(chat_id), state.to_dict())

    def _migrate_legacy(self, target: Path) -> bool:
        """Общий visit_state.json старой версии достаётся первому чату, который обратится к боту."""
        src = self.legacy_state
        if not src or not src.exists():
            return False
        try:
            data = read_json(src, None)
            if not isinstance(data, dict):
                return False
            atomic_write_json(target, RouteState.from_dict(data).to_dict())
            src.replace(src.with_name(src.stem + ".migrated.json"))
            return True
        except (json.JSONDecodeError, OSError) as e:
            log.error("Не удалось перенести %s: %s", src, e)
            return False
