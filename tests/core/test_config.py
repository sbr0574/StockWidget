"""配置保存与分页 / 同步参数迁移。"""

import os
import tempfile
import unittest
from unittest.mock import patch

from stockwidget.constants import APP_NAME
from stockwidget.core.config_store import config_paths, history_cache_dir, load_file, normalize_cache_directory, save_file
from stockwidget.core.view_options import ViewOptions, column_ranges, page_slice, taskbar_font_size_limit, taskbar_font_pixels


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

    def test_history_directory_uses_macos_caches_and_preserves_other_platform_defaults(self):
        for platform in ("win32", "linux", "darwin"):
            with self.subTest(platform=platform), patch("stockwidget.core.config_store.sys.platform", platform):
                expected = (os.path.join(os.path.expanduser("~"), "Library", "Caches", "com.sbr0574.StockWidget")
                            if platform == "darwin" else os.path.join(config_paths(), "cache"))
                self.assertEqual(history_cache_dir(), expected)
                custom = os.path.join(self._tmp.name, "custom")
                self.assertEqual(history_cache_dir(custom), custom)
                self.assertEqual(config_paths(), os.path.join(self._tmp.name, APP_NAME))

    def test_custom_cache_path_normalizes_and_invalid_config_uses_default(self):
        for value in (None, 123, [], {}, "", "  ", "bad\0path"):
            with self.subTest(value=value):
                self.assertEqual(normalize_cache_directory(value), "")
                self.assertEqual(history_cache_dir(value), history_cache_dir())
        self.assertEqual(normalize_cache_directory("  ~/chart-cache  "),
                         os.path.abspath(os.path.expanduser("~/chart-cache")))

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
    def test_chart_display_mode_legacy_default_validation_and_roundtrip(self):
        legacy = ViewOptions.from_config({"chart_enabled": True})
        self.assertTrue(legacy.chart_enabled)
        self.assertEqual(legacy.chart_display_mode, "window")
        for mode in ("window", "large", "medium", "small", "floating", None, "invalid", True):
            with self.subTest(mode=mode):
                options = ViewOptions.from_config({"chart_display_mode": mode})
                expected = "large" if mode == "floating" else mode if mode in ("window", "large", "medium", "small") else "window"
                self.assertEqual(options.chart_display_mode, expected)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)

    def test_chart_indicator_preferences_normalize_and_survive_disabled_chart(self):
        options = ViewOptions.from_config({"chart_enabled": False, "chart_ma_periods": [60, 5, 5, "10", -1],
                                           "chart_average_enabled": False, "chart_volume_enabled": False})
        self.assertEqual(options.chart_ma_periods, [5, 60])
        self.assertFalse(options.chart_average_enabled)
        self.assertFalse(options.chart_volume_enabled)
        self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        self.assertEqual(ViewOptions.from_config({"chart_ma_periods": []}).chart_ma_periods, [])
        self.assertEqual(ViewOptions.from_config({"chart_ma_periods": None}).chart_ma_periods, [5, 10, 20, 30, 60])

    def test_taskbar_font_limit_matches_rows_and_physical_scale(self):
        for height, dpi, rows, expected in ((88, 192, 4, 7), (88, 192, 3, 10),
                                            (44, 96, 4, 6), (44, 96, 3, 8), (88, 192, 1, 30)):
            with self.subTest(height=height, dpi=dpi, rows=rows):
                self.assertEqual(taskbar_font_size_limit(height, dpi, rows), expected)
                self.assertEqual(taskbar_font_pixels(height, dpi, rows, 100), (height - 4) // rows - 2)

    def test_color_mode_and_tray_preference_defaults_validation_and_roundtrip(self):
        defaults = ViewOptions.from_config({})
        self.assertEqual(defaults.color_mode, "system")
        self.assertFalse(defaults.hide_tray_icon)
        for mode in ("system", "light", "dark", None, "invalid"):
            with self.subTest(mode=mode):
                options = ViewOptions.from_config({"color_mode": mode, "hide_tray_icon": True})
                self.assertEqual(options.color_mode, mode if mode in ("system", "light", "dark") else "system")
                self.assertTrue(options.hide_tray_icon)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)

    def test_taskbar_automatic_color_defaults_preserve_legacy_manual_color_and_roundtrip(self):
        legacy = ViewOptions.from_config({"taskbar_color": "#123456", "taskbar_unicolor": False})
        self.assertFalse(legacy.taskbar_auto_color)
        configured = ViewOptions.from_config({**legacy.to_config(), "taskbar_auto_color": True})
        self.assertEqual(ViewOptions.from_config(configured.to_config()), configured)
        self.assertEqual(configured.taskbar_color, "#123456")
        self.assertFalse(configured.taskbar_unicolor)

    def test_float_row_and_interval_ranges_normalize_saved_configuration(self):
        for rows, seconds, expected in ((0, 0, (1, 1)), (1000, 3600, (20, 60)),
                                        (20, 60, (20, 60)), ("invalid", None, (3, 5))):
            with self.subTest(rows=rows, seconds=seconds):
                options = ViewOptions.from_config({"float_max_rows": rows, "float_page_interval": seconds})
                self.assertEqual((options.float_max_rows, options.float_page_interval), expected)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)

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
