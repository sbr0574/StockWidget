# -*- coding: utf-8 -*-
"""配置读写测试。"""

import os
import tempfile
import unittest

from stockwidget.constants import APP_NAME
from stockwidget.core.config_store import config_paths, load_file, save_file


class ConfigStoreTests(unittest.TestCase):
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
        self.assertEqual(config_paths(), os.path.join(self._tmp.name, APP_NAME))

    def test_save_and_load_roundtrip(self):
        save_file({"a": 1, "中文": "值"}, "c.json")
        self.assertEqual(load_file("c.json"), {"a": 1, "中文": "值"})

    def test_load_missing_returns_fallback(self):
        self.assertEqual(load_file("nope.json"), {})
        self.assertEqual(load_file("nope.json", {"d": 1}), {"d": 1})

    def test_load_corrupt_returns_fallback(self):
        directory = config_paths()
        os.makedirs(directory, exist_ok=True)
        with open(os.path.join(directory, "bad.json"), "w", encoding="utf-8") as file:
            file.write("{not valid json")
        self.assertEqual(load_file("bad.json"), {})


if __name__ == "__main__":
    unittest.main()
