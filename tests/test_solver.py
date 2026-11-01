import itertools
import random
import unittest

from routebot import solver
from routebot.solver import path_cost, solve, solve_subset

from .fakes import random_matrix


def brute_force(d, prefix):
    free = [i for i in range(len(d)) if i not in prefix]
    return min(path_cost(list(prefix) + list(p), d) for p in itertools.permutations(free))


class SolverTest(unittest.TestCase):
    def test_exact_matches_brute_force(self):
        for seed in range(25):
            n = random.Random(seed).randint(3, 8)
            d = random_matrix(n, seed)
            for prefix in [(0,), (0, n - 1)]:
                order = solve(d, prefix=prefix)
                self.assertEqual(order[: len(prefix)], list(prefix))
                self.assertEqual(sorted(order), list(range(n)))
                self.assertAlmostEqual(path_cost(order, d), brute_force(d, prefix), places=6)

    def test_ils_keeps_start_and_is_near_optimal(self):
        d = random_matrix(14, 7)
        ils = solve(d)  # 13 свободных > EXACT_MAX → эвристика
        self.assertEqual(ils[0], 0)
        old = solver.EXACT_MAX
        solver.EXACT_MAX = 13
        try:
            exact = solve(d)
        finally:
            solver.EXACT_MAX = old
        self.assertLessEqual(path_cost(ils, d), path_cost(exact, d) * 1.02)

    def test_large_route_is_permutation(self):
        d = random_matrix(40, 3)
        order = solve(d)
        self.assertEqual(order[0], 0)
        self.assertEqual(sorted(order), list(range(40)))

    def test_solve_subset_uses_original_indices(self):
        d = random_matrix(8, 1)
        order = solve_subset(d, (5,), [1, 3, 7])
        self.assertEqual(order[0], 5)
        self.assertEqual(sorted(order[1:]), [1, 3, 7])

    def test_trivial_sizes(self):
        self.assertEqual(solve([[0]]), [0])
        self.assertEqual(solve([[0, 1], [1, 0]]), [0, 1])


if __name__ == "__main__":
    unittest.main()
