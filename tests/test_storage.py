import json
import tempfile
import unittest
from pathlib import Path

from routebot.models import Point, RouteState, Status
from routebot.storage import StateStore


class StateStoreTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_chats_are_isolated(self):
        store = StateStore(self.root / "data")
        store.save(1, RouteState(points=[Point(name="A", id=1, status=Status.VISITED)]))
        store.save(2, RouteState(points=[Point(name="B", id=1)]))
        self.assertEqual(store.load(1).points[0].name, "A")
        self.assertEqual(store.load(2).points[0].name, "B")
        self.assertEqual(store.load(3).points, [])

    def test_corrupted_state_is_kept_aside(self):
        store = StateStore(self.root / "data")
        store.state_path(5).write_text("{broken", encoding="utf-8")
        self.assertEqual(store.load(5).points, [])
        self.assertEqual(len(list(store.chat_dir(5).glob("state.broken-*.json"))), 1)

    def test_legacy_state_migrates_once(self):
        legacy = self.root / "visit_state.json"
        legacy.write_text(
            json.dumps({"points": [{"order": 1, "name": "Старт", "addr": "x", "status": "visited"}]}),
            encoding="utf-8",
        )
        store = StateStore(self.root / "data", legacy_state=legacy)
        self.assertEqual(store.load(42).points[0].status, Status.VISITED)
        self.assertFalse(legacy.exists())
        self.assertTrue((self.root / "visit_state.migrated.json").exists())
        self.assertEqual(store.load(43).points, [])  # второй чат получает пустое состояние


if __name__ == "__main__":
    unittest.main()
