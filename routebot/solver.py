"""
Решатель задачи коммивояжёра для открытого маршрута (начало закреплено, конец свободный).

Матрица d асимметрична (односторонние улицы, развороты, мосты через Есиль),
поэтому все дельты считаются с учётом направления рёбер.

  ≤ EXACT_MAX свободных точек — точный оптимум (Held-Karp, O(2^n·n²));
  больше — Iterated Local Search: NN → 2-opt + Or-opt → встряски double-bridge.
"""
from __future__ import annotations

import random
import time
from collections.abc import Sequence

Matrix = Sequence[Sequence[float]]

EXACT_MAX = 12  # до стольких свободных точек — точный перебор (Held-Karp)
ILS_ITERS = 400  # итерации «встряска + локальный поиск» для больших маршрутов
ILS_TIME_LIMIT = 4.0  # сек, страховка для очень длинных списков


def path_cost(order: Sequence[int], d: Matrix) -> float:
    return sum(d[a][b] for a, b in zip(order, order[1:]))


def _nearest_neighbor(prefix, free, d):
    order, left = list(prefix), set(free)
    while left:
        nxt = min(left, key=lambda j: d[order[-1]][j])
        order.append(nxt)
        left.remove(nxt)
    return order


def _held_karp(prefix, free, d):
    """Точный оптимум открытого пути из prefix[-1] через все free (конец свободный)."""
    m, src = len(free), prefix[-1]
    full = (1 << m) - 1
    INF = float("inf")
    dp = [[INF] * m for _ in range(1 << m)]
    parent = [[-1] * m for _ in range(1 << m)]
    for j in range(m):
        dp[1 << j][j] = d[src][free[j]]
    for mask in range(1, full + 1):
        row = dp[mask]
        for j in range(m):
            cur = row[j]
            if cur == INF or not (mask >> j) & 1:
                continue
            fj = free[j]
            for k in range(m):
                if (mask >> k) & 1:
                    continue
                nm = mask | (1 << k)
                c = cur + d[fj][free[k]]
                if c < dp[nm][k]:
                    dp[nm][k], parent[nm][k] = c, j
    j = min(range(m), key=lambda x: dp[full][x])
    tail, mask = [], full
    while j != -1:
        tail.append(free[j])
        mask, j = mask ^ (1 << j), parent[mask][j]
    return list(prefix) + tail[::-1]


def _two_opt(order, d, lo):
    """Разворот отрезков order[i..j], i ≥ lo. Дельта считается за O(1) с учётом асимметрии d."""
    n, improved = len(order), False
    for i in range(lo, n - 1):
        a = order[i - 1]
        fwd = rev = 0.0
        for j in range(i + 1, n):
            fwd += d[order[j - 1]][order[j]]
            rev += d[order[j]][order[j - 1]]
            old = d[a][order[i]] + fwd
            new = d[a][order[j]] + rev
            if j + 1 < n:
                c = order[j + 1]
                old += d[order[j]][c]
                new += d[order[i]][c]
            if new < old - 1e-9:
                order[i : j + 1] = order[i : j + 1][::-1]
                improved = True
                break
    return improved


def _or_opt(order, d, lo):
    """Перенос цепочки из 1–3 точек (в прямом или обратном порядке) на лучшее место."""
    n = len(order)
    for seg_len in (1, 2, 3):
        for i in range(lo, n - seg_len + 1):
            seg = order[i : i + seg_len]
            rseg = seg[::-1]
            fwd = path_cost(seg, d)
            rev = path_cost(rseg, d)
            prev = order[i - 1]
            nxt = order[i + seg_len] if i + seg_len < n else None
            saved = d[prev][seg[0]] + fwd
            if nxt is not None:
                saved += d[seg[-1]][nxt] - d[prev][nxt]
            rest = order[:i] + order[i + seg_len :]
            best = None
            for k in range(lo, len(rest) + 1):
                p = rest[k - 1]
                q = rest[k] if k < len(rest) else None
                for cand, inner in ((seg, fwd), (rseg, rev)):
                    if k == i and cand is seg:
                        continue
                    added = d[p][cand[0]] + inner
                    if q is not None:
                        added += d[cand[-1]][q] - d[p][q]
                    if added < saved - 1e-9 and (best is None or added < best[0]):
                        best = (added, k, cand)
            if best:
                _, k, cand = best
                order[:] = rest[:k] + cand + rest[k:]
                return True
    return False


def _local_search(order, d, lo):
    while _two_opt(order, d, lo) or _or_opt(order, d, lo):
        pass
    return order


def _double_bridge(order, lo, rng):
    """Встряска: A B C D → A C B D на свободной части маршрута."""
    free = order[lo:]
    a, b, c = sorted(rng.sample(range(1, len(free)), 3))
    return order[:lo] + free[:a] + free[b:c] + free[a:b] + free[c:]


def solve(d: Matrix, prefix: Sequence[int] = (0,), seed: int = 42) -> list[int]:
    """
    Порядок объезда с минимальным суммарным временем.

    prefix — закреплённое начало маршрута (по умолчанию только СТАРТ = индекс 0);
    оно никогда не переставляется. Конец маршрута свободный.
    До EXACT_MAX свободных точек — точный оптимум, дальше — Iterated Local Search
    (2-opt + Or-opt + double-bridge), что убирает лишние переезды между берегами.
    """
    prefix = list(prefix)
    fixed = set(prefix)
    free = [i for i in range(len(d)) if i not in fixed]
    if len(free) <= 1:
        return prefix + free
    if len(free) <= EXACT_MAX:
        return _held_karp(prefix, free, d)

    lo = len(prefix)
    best = _local_search(_nearest_neighbor(prefix, free, d), d, lo)
    best_cost = path_cost(best, d)
    rng = random.Random(seed)
    deadline = time.monotonic() + ILS_TIME_LIMIT
    for _ in range(ILS_ITERS):
        if time.monotonic() > deadline:
            break
        cand = _local_search(_double_bridge(best, lo, rng), d, lo)
        cost = path_cost(cand, d)
        if cost < best_cost - 1e-9:
            best, best_cost = cand, cost
    return best



def solve_subset(d: Matrix, prefix: Sequence[int], nodes: Sequence[int]) -> list[int]:
    """solve() на подматрице: prefix закреплён, nodes упорядочиваются. Индексы — исходные."""
    sub = list(prefix) + list(nodes)
    dd = [[d[a][b] for b in sub] for a in sub]
    return [sub[i] for i in solve(dd, prefix=range(len(prefix)))]
