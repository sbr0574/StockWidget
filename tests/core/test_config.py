"""配置保存与分页 / 同步参数迁移。"""

import os
import tempfile
import unittest

from stockwidget.constants import APP_NAME
from stockwidget.core.config_store import config_paths, load_file, save_file
from stockwidget.core.view_options import ViewOptions, column_ranges, page_slice


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

    def test_missing_file_does_not_share_mutable_default(self):
        first = load_file("missing.json")
        first["changed"] = True
        self.assertEqual(load_file("missing.json"), {})

    def test_non_object_json_returns_fallback(self):
        directory = config_paths()
        os.makedirs(directory, exist_ok=True)
        for content in ("[]", "null", '"text"', "123"):
            with self.subTest(content=content):
                with open(os.path.join(directory, "bad.json"), "w", encoding="utf-8") as file:
                    file.write(content)
                self.assertEqual(load_file("bad.json", {"default": True}), {"default": True})

    def test_load_corrupt_returns_fallback(self):
        directory = config_paths()
        os.makedirs(directory, exist_ok=True)
        with open(os.path.join(directory, "bad.json"), "w", encoding="utf-8") as file:
            file.write("{not valid json")
        self.assertEqual(load_file("bad.json"), {})


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


class PageMathTests(unittest.TestCase):
    def test_split_defaults_roundtrip_and_odd_item_distribution(self):
        defaults = ViewOptions.from_config({})
        self.assertFalse(defaults.float_split_enabled)
        self.assertTrue(defaults.taskbar_sync_split)
        options = ViewOptions.from_config({"float_split_enabled": True, "float_split_separator": False,
                                           "taskbar_sync_split": False, "taskbar_split_enabled": True})
        self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        for total in (0, 1, 2, 5, 6):
            ranges = column_ranges(total, True)
            left, right = (stop - start for start, stop in ranges)
            self.assertIn(left - right, (0, 1))
            self.assertEqual([i for start, stop in ranges for i in range(start, stop)], list(range(total)))
            self.assertEqual(column_ranges(total, False), ((0, total),))

    def test_unified_sync_preferences_migrate_legacy_appearance_and_paging(self):
        defaults = ViewOptions.from_config({})
        self.assertTrue(defaults.taskbar_sync_appearance)
        self.assertTrue(defaults.taskbar_sync_paging)
        legacy = ViewOptions.from_config({"taskbar_sync_font": True, "taskbar_sync_color": False,
                                          "taskbar_page_mode": "auto"})
        self.assertFalse(legacy.taskbar_sync_appearance)
        self.assertFalse(legacy.taskbar_sync_paging)
        explicit = ViewOptions.from_config({**legacy.to_config(), "taskbar_sync_appearance": True,
                                            "taskbar_sync_paging": True, "taskbar_opacity_pct": 150})
        self.assertTrue(explicit.taskbar_sync_appearance)
        self.assertTrue(explicit.taskbar_sync_paging)
        self.assertEqual(explicit.taskbar_opacity_pct, 100)
        self.assertEqual(ViewOptions.from_config({"taskbar_opacity_pct": -10}).taskbar_opacity_pct, 0)

    def test_legacy_enabled_modes_and_paging_migrate_without_overriding_explicit_switches(self):
        self.assertFalse(ViewOptions.from_config({}).taskbar_enabled)
        self.assertFalse(ViewOptions.from_config({}).float_paging_enabled)
        old = ViewOptions.from_config({"float_max_rows": 4, "display_mode": "taskbar"})
        self.assertTrue(old.float_paging_enabled)
        self.assertTrue(old.taskbar_enabled)
        self.assertTrue(old.taskbar_sync_metrics)
        explicit = ViewOptions.from_config({"float_max_rows": 4, "display_mode": "both",
                                            "float_paging_enabled": False, "taskbar_enabled": False})
        self.assertFalse(explicit.float_paging_enabled)
        self.assertFalse(explicit.taskbar_enabled)

    def test_last_page_and_clamping(self):
        page = page_slice(5, 3, "manual", 1)
        self.assertEqual((page.start, page.stop, page.count, page.index), (3, 5, 2, 1))
        self.assertTrue(page.controls)
        self.assertEqual(page_slice(2, 3, "manual", 9).index, 0)
        self.assertFalse(page_slice(0, 3, "auto").controls)

    def test_first_rows_and_unlimited(self):
        page = page_slice(5, 3, "first", 8)
        self.assertEqual((page.start, page.stop, page.index), (0, 3, 0))
        self.assertFalse(page.controls)
        self.assertEqual(page_slice(5, 0, "auto", 1).stop, 5)
        self.assertFalse(page_slice(5, 0, "auto").controls)

    def test_options_normalize_and_migrate(self):
        opts = ViewOptions.from_config({"taskbar_rows": 99, "float_max_rows": -1,
            "taskbar_page_interval": "bad", "taskbar_metrics": ["name", "price", "name", "kline", "volume"],
            "display_mode": "both"})
        self.assertEqual(opts.taskbar_rows, 4)
        self.assertEqual(opts.float_max_rows, 1)
        self.assertFalse(opts.float_paging_enabled)
        self.assertEqual(opts.taskbar_page_interval, 5)
        self.assertEqual(opts.taskbar_metrics, ["name", "price", "kline", "volume"])
        self.assertTrue(opts.taskbar_dual_open)
        self.assertTrue(opts.taskbar_enabled)
        self.assertFalse(opts.taskbar_sync_metrics)
        self.assertEqual(ViewOptions.from_config(opts.to_config()).to_config(), opts.to_config())
