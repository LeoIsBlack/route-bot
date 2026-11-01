"""HTTP-сессии: по одной на поток (requests.Session не гарантирует потокобезопасность)."""
from __future__ import annotations

import threading

import requests

from . import config

_local = threading.local()


def session() -> requests.Session:
    s = getattr(_local, "session", None)
    if s is None:
        s = requests.Session()
        s.headers["User-Agent"] = config.USER_AGENT
        _local.session = s
    return s
