# -*- coding: utf-8 -*-
"""配置读写（config_store）的单元测试。"""

import os
import tempfile
import unittest

from stockwidget.core.config_store import config_paths, load_file, save_file


class TestConfigStore(unittest.TestCase):
    def setUp(self):
        self._old_appdata = os.environ.get("APPDATA")
        self._tmp = tempfile.TemporaryDirectory()
        os.environ["APPDATA"] = self._tmp.name

    def tearDown(self):
        self._tmp.cleanup()
        if self._old_appdata is None:
            os.environ.pop("APPDATA", None)
        else:
            os.environ["APPDATA"] = self._old_appdata

    def test_config_paths(self):
        self.assertEqual(
            config_paths("StockWidget"),
            os.path.join(self._tmp.name, "StockWidget"),
        )

    def test_save_and_load_roundtrip(self):
        save_file({"a": 1, "中文": "值"}, "StockWidget", "c.json")
        self.assertEqual(load_file("StockWidget", "c.json"), {"a": 1, "中文": "值"})

    def test_load_missing_returns_default(self):
        self.assertEqual(load_file("StockWidget", "nope.json"), {})
        self.assertEqual(load_file("StockWidget", "nope.json", {"d": 1}), {"d": 1})

    def test_load_corrupt_returns_default(self):
        d = os.path.join(self._tmp.name, "StockWidget")
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "bad.json"), "w", encoding="utf-8") as f:
            f.write("{not valid json")
        self.assertEqual(load_file("StockWidget", "bad.json"), {})


if __name__ == "__main__":
    unittest.main()
