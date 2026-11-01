import unittest

from routebot.models import Point, RouteState, Status


def make_state(n: int = 5) -> RouteState:
    pts = [Point(name=f"P{i}", addr=f"ул. {i}", lat=51.1, lon=71.4 + i / 100, mins=5) for i in range(1, n + 1)]
    pts[0].is_start, pts[0].mins = True, 0
    state = RouteState(points=pts)
    state.renumber()
    state.ensure_ids()
    return state


class NextToVisitTest(unittest.TestCase):
    def test_goes_forward_then_wraps_then_skipped(self):
        s = make_state()
        s.points[0].status = Status.SKIPPED
        s.points[2].status = Status.VISITED
        self.assertEqual(s.next_to_visit(after_order=3)[0].order, 4)  # вперёд по маршруту
        s.points[3].status = s.points[4].status = Status.VISITED
        self.assertEqual(s.next_to_visit(after_order=5)[0].order, 2)  # с начала
        s.points[1].status = Status.CLOSED
        p, from_skipped = s.next_to_visit()
        self.assertEqual((p.order, from_skipped), (1, True))  # остались только отложенные
        s.points[0].status = Status.VISITED
        self.assertEqual(s.next_to_visit(), (None, False))


class MergeProgressTest(unittest.TestCase):
    def test_keeps_ids_and_notes_by_name_and_address(self):
        old = make_state()
        old.points[1].status, old.points[1].notes = Status.VISITED, "дорого"
        old.points[1].photos = ["a.jpg"]
        new = make_state()
        new.points.reverse()
        for p in new.points:
            p.id = 0
        new.merge_progress(old)
        p2 = next(p for p in new.points if p.name == "P2")
        self.assertEqual((p2.id, p2.status, p2.notes, p2.photos), (old.points[1].id, Status.VISITED, "дорого", ["a.jpg"]))
        self.assertEqual(len({p.id for p in new.points}), 5)

    def test_duplicate_shops_get_distinct_ids(self):
        old = RouteState(points=[Point(name="X", addr="A"), Point(name="X", addr="A")])
        old.ensure_ids()
        old.points[0].status = Status.VISITED
        new = RouteState(points=[Point(name="X", addr="A"), Point(name="X", addr="A"), Point(name="Y")])
        new.merge_progress(old)
        self.assertEqual([p.status for p in new.points], [Status.VISITED, Status.PENDING, Status.PENDING])
        self.assertEqual(len({p.id for p in new.points}), 3)


class RerouteTest(unittest.TestCase):
    def test_done_first_then_new_order_then_leftovers(self):
        s = make_state()
        s.points.append(Point(name="NoCoords", addr="?", id=99))
        s.renumber()
        s.points[0].status = Status.VISITED
        ids = [p.id for p in s.points]
        # пока шёл расчёт, точку 3 успели закрыть — её нельзя перемещать
        s.points[2].status = Status.CLOSED
        rerouted = s.apply_reroute([(ids[4], 2.0), (ids[2], 1.0), (ids[1], 3.0)])
        self.assertEqual([p.id for p in rerouted], [ids[4], ids[1]])
        self.assertEqual(
            [p.id for p in s.points], [ids[0], ids[2], ids[4], ids[1], ids[3], 99]
        )
        self.assertEqual([p.order for p in s.points], [1, 2, 3, 4, 5, 6])
        self.assertFalse(any(p.is_start for p in rerouted))


class SerializationTest(unittest.TestCase):
    def test_roundtrip(self):
        s = make_state()
        s.points[0].photos = ["x.jpg"]
        s.points[1].status = Status.CLOSED
        again = RouteState.from_dict(s.to_dict())
        self.assertEqual(again.to_dict(), s.to_dict())

    def test_legacy_v2_state(self):
        legacy = {
            "city": "Астана",
            "points": [
                {"order": 1, "name": "A", "addr": "a", "mins": 0, "total": 0, "lat": 51.1, "lon": 71.4,
                 "gmaps": "…", "dgis": "…", "status": "visited", "photo": "p1.jpg, p2.jpg"},
                {"order": 2, "name": "B", "addr": "b", "mins": 4, "lat": None, "lon": None, "status": "weird"},
            ],
        }
        s = RouteState.from_dict(legacy)
        a, b = s.points
        self.assertTrue(a.is_start)
        self.assertEqual(a.photos, ["p1.jpg", "p2.jpg"])
        self.assertEqual(b.status, Status.PENDING)
        self.assertFalse(b.has_coords)
        self.assertNotEqual(a.id, b.id)


if __name__ == "__main__":
    unittest.main()
