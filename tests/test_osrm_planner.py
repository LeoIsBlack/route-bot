import unittest

from routebot.models import Point
from routebot.osrm import OsrmClient, OsrmError
from routebot.planner import PlanError, RoutePlanner, reachable_indices

from .fakes import FakeGeocoder, FakeOsrm, FakeOsrmHttp, travel_seconds


def grid(n: int) -> list[tuple[float, float]]:
    return [(51.0 + (i % 13) / 100, 71.3 + (i // 13) / 100) for i in range(n)]


class OsrmClientTest(unittest.TestCase):
    def test_large_matrix_is_requested_in_blocks(self):
        coords = grid(130)
        http = FakeOsrmHttp()
        d = OsrmClient(session=http, retry_delay=0).durations(coords)
        self.assertGreater(http.calls, 1)
        for i in (0, 64, 129):
            for j in (0, 77, 128):
                self.assertEqual(d[i][j], travel_seconds(coords[i], coords[j]))

    def test_retries_temporary_errors(self):
        http = FakeOsrmHttp(fail_first=2)
        d = OsrmClient(session=http, retries=3, retry_delay=0).durations(grid(3))
        self.assertEqual(len(d), 3)
        self.assertEqual(http.calls, 3)

    def test_gives_up_with_clear_error(self):
        with self.assertRaises(OsrmError):
            OsrmClient(session=FakeOsrmHttp(fail_first=10), retries=2, retry_delay=0).durations(grid(3))


class ReachableTest(unittest.TestCase):
    def test_drops_only_isolated_point(self):
        d = [[0, 1, None], [1, 0, None], [None, None, 0]]
        self.assertEqual(reachable_indices(d), [0, 1])

    def test_isolated_start_is_error(self):
        d = [[0, None, None], [None, 0, 1], [None, 1, 0]]
        with self.assertRaises(PlanError):
            reachable_indices(d)


class PlannerTest(unittest.TestCase):
    def setUp(self):
        self.coords = {f"ул. {i}": c for i, c in enumerate(grid(8))}
        self.points = [Point(num=i + 1, name=f"P{i}", addr=f"ул. {i}") for i in range(8)]

    def test_plan_keeps_start_and_reports_missing(self):
        self.points.append(Point(num=9, name="Нет", addr="нигде"))
        blocked = {self.coords["ул. 5"]}
        result = RoutePlanner(FakeGeocoder(self.coords), FakeOsrm(blocked)).plan(self.points)
        self.assertEqual(result.route[0].name, "P0")
        self.assertTrue(result.route[0].is_start)
        self.assertEqual(sorted(p.name for p in result.missing), ["P5", "Нет"])
        self.assertIn("OSRM", next(p for p in result.missing if p.name == "P5").fail_reason)
        self.assertEqual([p.order for p in result.points], list(range(1, 10)))

    def test_missing_start_is_plan_error(self):
        self.points[0].addr = "нигде"
        with self.assertRaises(PlanError):
            RoutePlanner(FakeGeocoder(self.coords), FakeOsrm()).plan(self.points)

    def test_reroute_closest_first_and_deferred_last(self):
        geo = FakeGeocoder(self.coords)
        for i, p in enumerate(self.points):
            geo.locate(p)
            p.id = i + 1
        pending, deferred = self.points[1:6], self.points[6:]
        here = self.coords["ул. 4"]
        seq = RoutePlanner(geo, FakeOsrm()).reroute(here[0] + 1e-5, here[1], pending, deferred)
        ids = [pid for pid, _ in seq]
        self.assertEqual(ids[0], 5)  # P4 — ближайшая к позиции
        self.assertEqual(sorted(ids[:5]), [2, 3, 4, 5, 6])
        self.assertEqual(sorted(ids[5:]), [7, 8])

    def test_reroute_skips_unreachable(self):
        geo = FakeGeocoder(self.coords)
        for i, p in enumerate(self.points):
            geo.locate(p)
            p.id = i + 1
        blocked = {self.coords["ул. 2"]}
        seq = RoutePlanner(geo, FakeOsrm(blocked)).reroute(51.0, 71.3, self.points[1:4])
        self.assertEqual(sorted(pid for pid, _ in seq), [2, 4])


if __name__ == "__main__":
    unittest.main()
