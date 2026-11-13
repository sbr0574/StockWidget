"""配置保存与分页 / 同步参数迁移。"""

import os
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from stockwidget.constants import APP_NAME, CONFIG_FILE
from stockwidget.core import config_store as store
from stockwidget.core.config_store import config_paths, history_cache_dir, load_file, save_file
from stockwidget.core.view_options import (
    ViewOptions, column_ranges, page_slice, taskbar_font_pixels, taskbar_font_point_limit,
)


class ConfigStoreTests(unittest.TestCase):
    def setUp(self):
        self._tmp = self.enterContext(tempfile.TemporaryDirectory())
        self.enterContext(patch("stockwidget.core.config_store.sys.platform", "linux"))
        self.enterContext(patch("stockwidget.core.config_store.os.path.expanduser", return_value=self._tmp))

    def test_config_paths(self):
        self.assertEqual(config_paths(), os.path.join(self._tmp, APP_NAME))

    def test_history_directory_is_within_the_fixed_program_directory(self):
        self.assertEqual(history_cache_dir(), os.path.join(config_paths(), "cache"))

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


class DataDirectoryTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory())).resolve()
        self.default, self.program = self.root / "default", self.root / "program"
        self.target = self.root / "selected"
        self.enterContext(patch.object(store, "_active_directory", None))
        self.enterContext(patch.object(store, "default_data_directory", return_value=str(self.default)))
        self.enterContext(patch.object(store, "program_directory", return_value=self.program))
        self.enterContext(patch.object(store.sys, "platform", "win32"))

    def test_windows_prefers_appdata_even_when_program_directory_contains_configuration(self):
        save_file({"portable": True}, CONFIG_FILE, directory=str(self.program))
        self.assertIsNone(store.activate_data_directory())
        self.assertEqual(Path(config_paths()), self.default)
        self.assertFalse((self.program / store._LOCATION_FILE).exists())

    def test_windows_fallback_preserves_readable_configuration_and_survives_restart(self):
        config = {"watchlist": {"sh600000": {"checked": True}}}
        save_file(config, CONFIG_FILE, directory=str(self.default))
        probe, save = store._probe_directory, store.save_file

        def probe_writable(path):
            if path == self.default:
                raise PermissionError("read-only APPDATA")
            probe(path)

        def save_writable(data, name, *, directory=None):
            if directory == str(self.default):
                raise PermissionError("read-only APPDATA")
            save(data, name, directory=directory)

        with patch.object(store, "_probe_directory", side_effect=probe_writable), \
                patch.object(store, "save_file", side_effect=save_writable):
            self.assertEqual(store.activate_data_directory(), config)
        self.assertEqual(Path(config_paths()), self.program)
        self.assertEqual(load_file(CONFIG_FILE), config)
        self.assertEqual(load_file(CONFIG_FILE, directory=str(self.default)), config)
        self.assertTrue((self.program / store._LOCATION_FILE).is_file())
        store._active_directory = None
        store.activate_data_directory()
        self.assertEqual(Path(config_paths()), self.program)

    def test_both_locations_unwritable_and_non_windows_hosts_do_not_abort_startup(self):
        with patch.object(store, "_probe_directory", side_effect=PermissionError()):
            store.activate_data_directory()
        self.assertEqual(Path(config_paths()), self.default)
        for platform in ("darwin", "linux"):
            with self.subTest(platform=platform), patch.object(store.sys, "platform", platform), \
                    patch.object(store, "_probe_directory") as probe:
                store.activate_data_directory()
                probe.assert_not_called()
                self.assertEqual(Path(config_paths()), self.default)

    def test_location_record_write_failure_falls_back_only_on_windows(self):
        original_save = store.save_file

        def save_writable(data, name, *, directory=None):
            if directory == str(self.default):
                raise PermissionError("read-only default directory")
            original_save(data, name, directory=directory)

        for platform in ("win32", "darwin", "linux"):
            program = self.root / (platform + "-program")
            with self.subTest(platform=platform), patch.object(store.sys, "platform", platform), \
                    patch.object(store, "program_directory", return_value=program), \
                    patch.object(store, "save_file", side_effect=save_writable):
                if platform == "win32":
                    store._save_location(self.target)
                    record = load_file(store._LOCATION_FILE, directory=str(program))
                    self.assertEqual(record["directory"], str(self.target))
                else:
                    with self.assertRaises(OSError):
                        store._save_location(self.target)
                    self.assertFalse(program.exists())

    def test_default_location_records_survive_restart_on_all_platforms(self):
        configuration = {"watchlist": {"sh600000": {"checked": True}}}
        save_file(configuration, CONFIG_FILE, directory=str(self.target))
        for platform in ("win32", "darwin", "linux"):
            with self.subTest(platform=platform), patch.object(store.sys, "platform", platform), \
                    patch.object(store, "_active_directory", None):
                store._save_location(self.target)
                store.activate_data_directory()
                self.assertEqual(Path(config_paths()), self.target)
                self.assertEqual(load_file(CONFIG_FILE), configuration)
                self.assertFalse(self.program.exists())

    def test_legacy_program_location_records_remain_readable_on_non_windows_hosts(self):
        configuration = {"watchlist": {"sh600000": {"checked": True}}}
        for platform in ("darwin", "linux"):
            for portable in (False, True):
                selected = self.program if portable else self.target
                save_file(configuration, CONFIG_FILE, directory=str(selected))
                save_file({"version": 1, "directory": str(self.target), "portable": portable},
                          store._LOCATION_FILE, directory=str(self.program))
                record = self.program / store._LOCATION_FILE
                original = record.read_bytes()
                with self.subTest(platform=platform, portable=portable), \
                        patch.object(store.sys, "platform", platform), \
                        patch.object(store, "_active_directory", None):
                    store.activate_data_directory()
                    self.assertEqual(Path(config_paths()), selected)
                    self.assertEqual(load_file(CONFIG_FILE), configuration)
                    self.assertEqual(record.read_bytes(), original)
                    self.assertFalse(self.default.exists())

    def test_non_windows_location_write_failure_rolls_back_directory_selection(self):
        original_save = store.save_file

        def save_writable(data, name, *, directory=None):
            if name == store._LOCATION_FILE and directory == str(self.default):
                raise PermissionError("read-only default directory")
            original_save(data, name, directory=directory)

        save_file({"source": True}, CONFIG_FILE, directory=str(self.default))
        save_file({"destination": True}, CONFIG_FILE, directory=str(self.target))
        for root, value in ((self.default, "source"), (self.target, "destination")):
            (root / "data").mkdir()
            (root / "data" / "stock.json").write_text(value, encoding="utf-8")
        for platform in ("darwin", "linux"):
            with self.subTest(platform=platform), patch.object(store.sys, "platform", platform), \
                    patch.object(store, "save_file", side_effect=save_writable):
                store.activate_data_directory()
                with self.assertRaises(OSError):
                    store.prepare_data_directory(str(self.target), {"replacement": True})
                self.assertEqual(load_file(CONFIG_FILE, directory=str(self.target)), {"destination": True})
                self.assertEqual((self.target / "data" / "stock.json").read_text(), "destination")
                self.assertEqual((self.default / "data" / "stock.json").read_text(), "source")
                self.assertEqual(Path(config_paths()), self.default)
                self.assertFalse(self.program.exists())

    def test_failed_first_fallback_save_does_not_hide_the_readable_original_on_restart(self):
        config = {"watchlist": {"kept": {}}}
        save_file(config, CONFIG_FILE, directory=str(self.default))
        probe = store._probe_directory

        def probe_writable(path):
            if path == self.default:
                raise PermissionError()
            probe(path)

        with patch.object(store, "_probe_directory", side_effect=probe_writable), \
                patch.object(store, "save_file", side_effect=PermissionError()):
            for _ in range(2):
                self.assertEqual(store.activate_data_directory(), config)
                self.assertEqual(Path(config_paths()), self.program)
                self.assertFalse((self.program / store._LOCATION_FILE).exists())

    def test_invalid_location_records_do_not_change_the_default(self):
        for value in ("", "relative", "\0", None):
            with self.subTest(value=value):
                save_file({"version": 1, "directory": value}, store._LOCATION_FILE, directory=str(self.default))
                store.activate_data_directory()
                self.assertEqual(Path(config_paths()), self.default)

    def test_directory_probe_preserves_existing_files_and_cleans_failed_temporaries(self):
        self.default.mkdir()
        sentinel = self.default / ".stockwidget-write-test"
        sentinel.write_bytes(b"user file")
        store._probe_directory(self.default)
        with patch.object(store.os, "replace", side_effect=PermissionError()):
            with self.assertRaises(OSError):
                store._probe_directory(self.default)
        self.assertEqual(list(self.default.iterdir()), [sentinel])
        self.assertEqual(sentinel.read_bytes(), b"user file")

    def test_selection_copies_managed_data_and_wal_snapshot_without_deleting_source(self):
        store.activate_data_directory()
        (self.default / "data").mkdir()
        (self.default / "data" / "stock.json").write_text('{"codes":{}}', encoding="utf-8")
        (self.default / "unrelated.exe").write_bytes(b"program")
        (self.default / "cache").mkdir()
        database = sqlite3.connect(self.default / "cache" / "history.sqlite3")
        self.addCleanup(database.close)
        database.execute("PRAGMA journal_mode=WAL")
        database.execute("CREATE TABLE saved (value TEXT)")
        database.execute("INSERT INTO saved VALUES ('committed in WAL')")
        database.commit()
        config = {"watchlist": {"sh600000": {"cost": 15}}, "taskbar_font_size": 11}
        icon = self.root / "mine.png"
        icon.write_bytes(b"image")
        config["custom_icon_path"] = str(icon)
        copied = store.prepare_data_directory(str(self.target), config)
        self.assertEqual(Path(config_paths()), self.default)
        self.assertTrue((self.default / "cache" / "history.sqlite3").exists())
        self.assertEqual((self.target / "data" / "stock.json").read_bytes(), (self.default / "data" / "stock.json").read_bytes())
        self.assertFalse((self.target / "unrelated.exe").exists())
        self.assertEqual(Path(copied["custom_icon_path"]).read_bytes(), b"image")
        copy = sqlite3.connect(self.target / "cache" / "history.sqlite3")
        try:
            self.assertEqual(copy.execute("SELECT value FROM saved").fetchone(), ("committed in WAL",))
        finally:
            copy.close()
        self.assertFalse((self.target / "cache" / "history.sqlite3-wal").exists())
        store.activate_data_directory()
        self.assertEqual(Path(config_paths()), self.target)
        self.assertEqual(load_file(CONFIG_FILE), copied)
        self.assertEqual(Path(store.data_cache_dir()), self.target / "data")
        self.assertEqual(Path(history_cache_dir()), self.target / "cache")

    def test_existing_destination_is_backed_up_before_replacement(self):
        store.activate_data_directory()
        save_file({"old": True}, CONFIG_FILE, directory=str(self.target))
        (self.target / "data").mkdir()
        (self.target / "data" / "stock.json").write_text("old", encoding="utf-8")
        store.prepare_data_directory(str(self.target), {"watchlist": {"new": {}}})
        backup, = self.target.glob("previous-data-*")
        self.assertEqual(load_file(CONFIG_FILE, directory=str(backup)), {"old": True})
        self.assertEqual((backup / "data" / "stock.json").read_text(encoding="utf-8"), "old")

    def test_location_write_failure_rolls_back_installed_files(self):
        store.activate_data_directory()
        save_file({"old": True}, CONFIG_FILE, directory=str(self.target))
        for root, value in ((self.default, "source"), (self.target, "destination")):
            (root / "data").mkdir()
            (root / "data" / "stock.json").write_text(value, encoding="utf-8")
        with patch.object(store, "_save_location", side_effect=PermissionError()):
            with self.assertRaises(OSError):
                store.prepare_data_directory(str(self.target), {"new": True})
        self.assertEqual(load_file(CONFIG_FILE, directory=str(self.target)), {"old": True})
        self.assertEqual((self.target / "data" / "stock.json").read_text(encoding="utf-8"), "destination")
        self.assertEqual((self.default / "data" / "stock.json").read_text(encoding="utf-8"), "source")
        self.assertEqual(Path(config_paths()), self.default)

    def test_selection_rejects_unknown_destination_and_recursive_paths(self):
        store.activate_data_directory()
        (self.target / "data").mkdir(parents=True)
        for path in (self.target, self.default / "data" / "nested", self.default.parent):
            with self.subTest(path=path), self.assertRaises(OSError):
                store.prepare_data_directory(str(path), {})
        self.assertEqual(Path(config_paths()), self.default)


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


class PageMathTests(unittest.TestCase):
    def test_chart_opacity_preferences_roundtrip_without_overwriting_independent_values(self):
        defaults = ViewOptions.from_config({})
        self.assertTrue(defaults.chart_enabled)
        self.assertTrue(defaults.chart_sync_opacity)
        options = ViewOptions.from_config({"chart_enabled": False, "chart_sync_opacity": False,
                                           "chart_bg_alpha_pct": 35, "chart_opacity_pct": 45})
        synced = ViewOptions.from_config({**options.to_config(), "chart_sync_opacity": True})
        self.assertEqual((synced.chart_bg_alpha_pct, synced.chart_opacity_pct), (35, 45))
        self.assertFalse(ViewOptions.from_config(synced.to_config()).chart_enabled)
        self.assertEqual(ViewOptions.from_config(options.to_config()), options)

    def test_chart_sizes_migrate_presets_and_roundtrip_without_display_mode(self):
        legacy = ViewOptions.from_config({})
        self.assertTrue(legacy.chart_enabled)
        self.assertEqual(legacy.chart_popup_size, [480, 340])
        for mode in ("window", "large", "medium", "small", "floating", None, "invalid", True):
            with self.subTest(mode=mode):
                options = ViewOptions.from_config({"chart_display_mode": mode})
                expected = {"large": [560, 400], "floating": [560, 400], "medium": [420, 300], "small": [320, 230]}.get(mode, [480, 340])
                self.assertEqual(options.chart_popup_size, expected)
                self.assertNotIn("chart_display_mode", options.to_config())
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)

    def test_saved_chart_sizes_validate_and_override_legacy_preset(self):
        options = ViewOptions.from_config({"chart_popup_size": [530, 385], "chart_display_mode": "small",
                                           "chart_window_size": [1030, 650], "chart_order_book_enabled": True})
        self.assertEqual(options.chart_popup_size, [530, 385])
        self.assertEqual(options.chart_window_size, [1030, 650])
        self.assertTrue(options.chart_order_book_enabled)
        self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        for invalid in (None, True, "large", [400], {"w": 400}):
            self.assertEqual(ViewOptions.from_config({"chart_popup_size": invalid}).chart_popup_size, [480, 340])
        self.assertEqual(ViewOptions.from_config({"chart_popup_size": [None, -10]}).chart_popup_size, [480, 200])

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
                                            (44, 96, 4, 6), (44, 96, 3, 8), (88, 192, 1, 15)):
            with self.subTest(height=height, dpi=dpi, rows=rows):
                self.assertEqual(taskbar_font_point_limit(height, dpi, rows), expected)
                self.assertLessEqual(taskbar_font_pixels(height, dpi, rows, 100), (height - 4) // rows - 2)

    def test_taskbar_font_range_clamps_saved_preferences_and_rendering(self):
        for stored, expected in ((0, 5), (5, 5), (8, 8), (15, 15), (30, 15), ("invalid", 8)):
            with self.subTest(stored=stored):
                options = ViewOptions.from_config({"taskbar_font_size": stored})
                self.assertEqual(options.taskbar_font_size, expected)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        self.assertEqual(taskbar_font_pixels(88, 192, 1, 0), 13)
        self.assertEqual(taskbar_font_pixels(88, 192, 1, 30), 40)

    def test_taskbar_size_roundtrip_does_not_reintroduce_independent_font_family(self):
        options = ViewOptions.from_config({"taskbar_font_family": "Arial", "taskbar_font_size": 12})
        self.assertEqual(ViewOptions.from_config(options.to_config()).taskbar_font_size, 12)
        self.assertNotIn("taskbar_font_family", options.to_config())

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
        self.assertNotIn("taskbar_unicolor", configured.to_config())

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
        self.assertEqual(defaults.taskbar_opacity_pct, 50)
        self.assertTrue(defaults.taskbar_sync_paging)
        legacy = ViewOptions.from_config({"taskbar_sync_font": True, "taskbar_sync_color": False,
                                          "taskbar_page_mode": "auto"})
        self.assertFalse(legacy.taskbar_sync_paging)
        explicit = ViewOptions.from_config({**legacy.to_config(), "taskbar_sync_appearance": True,
                                            "taskbar_sync_paging": True, "taskbar_auto_color": True,
                                            "fg": "#123456", "opacity_pct": 42})
        self.assertTrue(explicit.taskbar_sync_paging)
        self.assertEqual(explicit.taskbar_opacity_pct, 42)
        self.assertEqual(explicit.taskbar_color, "#123456")
        self.assertFalse(explicit.taskbar_auto_color)
        self.assertNotIn("taskbar_sync_appearance", explicit.to_config())
        self.assertEqual(ViewOptions.from_config(explicit.to_config()), explicit)
        self.assertEqual(ViewOptions.from_config({"taskbar_opacity_pct": 150}).taskbar_opacity_pct, 100)
        self.assertEqual(ViewOptions.from_config({"taskbar_opacity_pct": -10}).taskbar_opacity_pct, 10)

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
