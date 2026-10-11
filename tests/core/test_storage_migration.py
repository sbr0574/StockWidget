"""固定目录、macOS 数据迁移与旧目录安全清理。"""

from contextlib import closing
from pathlib import Path
import os
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from stockwidget.constants import CONFIG_FILE, CODE_LIST_FILES, CODES_STATUS_FILE
from stockwidget.core import config_store as store


class MacOSMigrationTests(unittest.TestCase):
    def setUp(self):
        self.home = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.enterContext(patch.object(store.sys, "platform", "darwin"))
        self.enterContext(patch.object(store.os.path, "expanduser", side_effect=lambda path:
                                     str(self.home) + path[1:] if path.startswith("~") else path))
        self.legacy = self.home / "StockWidget"
        self.current = self.home / "Library" / "Application Support" / "StockWidget"

    def write_database(self, directory, value="history", version=1):
        directory.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(directory / "history.sqlite3")) as db, db:
            db.execute("CREATE TABLE sample (value TEXT)")
            db.execute("INSERT INTO sample VALUES (?)", (value,))
            db.execute(f"PRAGMA user_version={version}")

    def read_database(self, directory):
        with closing(sqlite3.connect(directory / "history.sqlite3")) as db:
            return db.execute("SELECT value FROM sample").fetchone()[0]

    def test_home_configuration_codes_cache_and_extra_files_migrate_before_old_directory_is_deleted(self):
        config = {"watchlist": {"sh600519": {"checked": True, "cost": 123}}, "app_icon": "dark"}
        store.save_file(config, CONFIG_FILE, directory=str(self.legacy))
        store.save_file({"legacy": True}, CODE_LIST_FILES[0], directory=str(self.legacy))
        store.save_file({"downloaded": True}, CODE_LIST_FILES[1], directory=str(self.legacy / "data"))
        store.save_file({"run_date": "2026-10-09"}, CODES_STATUS_FILE, directory=str(self.legacy / "data"))
        (self.legacy / "extra.txt").write_text("keep this too")
        self.write_database(self.legacy / "cache")
        self.assertEqual(store.initialize_data_directory(), config)
        self.assertEqual(store.config_paths(), str(self.current))
        self.assertFalse(self.legacy.exists())
        self.assertEqual(store.load_file(CODE_LIST_FILES[0]), {"legacy": True})
        self.assertEqual(store.load_file(CODE_LIST_FILES[1], directory=store.data_cache_dir()), {"downloaded": True})
        self.assertEqual(store.load_file(CODES_STATUS_FILE, directory=store.data_cache_dir()), {"run_date": "2026-10-09"})
        self.assertEqual((self.current / "extra.txt").read_text(), "keep this too")
        self.assertEqual(self.read_database(self.current / "cache"), "history")
        (self.current / "cache" / "history.sqlite3").unlink()
        self.assertEqual(store.initialize_data_directory(), config)
        self.assertFalse((self.current / "cache" / "history.sqlite3").exists())

    def test_existing_new_configuration_wins_and_entire_old_directory_is_backed_up_before_deletion(self):
        old = {"watchlist": {"old": {}}}
        current = {"watchlist": {"current": {}}}
        store.save_file(old, CONFIG_FILE, directory=str(self.legacy))
        store.save_file(current, CONFIG_FILE, directory=str(self.current))
        (self.legacy / "extra.txt").write_text("keep")
        self.write_database(self.legacy / "cache", "old history")
        self.assertEqual(store.initialize_data_directory(), current)
        self.assertFalse(self.legacy.exists())
        backup = self.current / "legacy_backup"
        self.assertEqual(store.load_file(CONFIG_FILE, directory=str(backup)), old)
        self.assertEqual((backup / "extra.txt").read_text(), "keep")
        self.assertEqual(self.read_database(backup / "cache"), "old history")
        self.assertFalse((self.current / "cache" / "history.sqlite3").exists())

    def test_old_home_code_files_are_normalized_without_overwriting_data_files_or_losing_originals(self):
        for existing_new_data in (False, True):
            with self.subTest(existing_new_data=existing_new_data):
                store.save_file({"watchlist": {"old": {}}}, CONFIG_FILE, directory=str(self.legacy))
                root_codes, stale, preferred = {"root": True}, {"stale": True}, {"preferred": True}
                status = {"run_date": "2026-10-10"}
                store.save_file(root_codes, CODE_LIST_FILES[0], directory=str(self.legacy))
                store.save_file(stale, CODE_LIST_FILES[1], directory=str(self.legacy))
                store.save_file(preferred, CODE_LIST_FILES[1], directory=str(self.legacy / "data"))
                store.save_file(status, CODES_STATUS_FILE, directory=str(self.legacy))
                if existing_new_data:
                    store.save_file({"watchlist": {"current": {}}}, CONFIG_FILE)
                    store.save_file({"new": True}, CODE_LIST_FILES[0], directory=store.data_cache_dir())
                store.initialize_data_directory()
                directory = self.current / "legacy_backup" if existing_new_data else self.current
                self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=str(directory / "data")), root_codes)
                self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=str(directory)), root_codes)
                self.assertEqual(store.load_file(CODE_LIST_FILES[1], directory=str(directory / "data")), preferred)
                self.assertEqual(store.load_file(CODE_LIST_FILES[1], directory=str(directory)), stale)
                self.assertEqual(store.load_file(CODES_STATUS_FILE, directory=str(directory / "data")), status)
                self.assertEqual(store.load_file(CODES_STATUS_FILE, directory=str(directory)), status)
                if existing_new_data:
                    self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=store.data_cache_dir()), {"new": True})
                self.assertFalse(self.legacy.exists())

    def test_code_file_links_keep_their_targets_without_writing_through_data_directory_links(self):
        snapshot = self.home / "market_snapshot.json"
        snapshot.write_text('{"codes": {"sh600519": {"name": "贵州茅台"}}}', encoding="utf-8")
        archive = self.home / "code_archive"
        archive.mkdir()
        payload = {"codes": {"sh600519": {"name": "贵州茅台"}}}
        for linked_data_directory in (False, True):
            with self.subTest(linked_data_directory=linked_data_directory):
                store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
                (self.legacy / CODE_LIST_FILES[0]).symlink_to("../market_snapshot.json")
                if linked_data_directory:
                    (self.legacy / "data").symlink_to("../code_archive", target_is_directory=True)
                store.initialize_data_directory()
                directory = self.current / "legacy_backup" if linked_data_directory else self.current
                self.assertEqual((directory / CODE_LIST_FILES[0]).resolve(), snapshot.resolve())
                self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=str(directory)), payload)
                data = directory / "data"
                if linked_data_directory:
                    self.assertTrue(data.is_symlink())
                    self.assertEqual(data.resolve(), archive.resolve())
                    self.assertEqual(list(archive.iterdir()), [])
                else:
                    self.assertFalse((data / CODE_LIST_FILES[0]).is_symlink())
                    self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=str(data)), payload)
                self.assertFalse(self.legacy.exists())

    def test_copy_failure_retains_old_directory_and_configuration_for_retry(self):
        old = {"watchlist": {"old": {}}}
        store.save_file(old, CONFIG_FILE, directory=str(self.legacy))
        with patch.object(store, "_stage_history", side_effect=OSError("read only")), self.assertRaises(OSError):
            store.initialize_data_directory()
        self.assertFalse((self.current / CONFIG_FILE).exists())
        self.assertEqual(store.load_file(CONFIG_FILE, directory=str(self.legacy)), old)
        self.assertEqual(store.initialize_data_directory(), old)
        self.assertFalse(self.legacy.exists())

    def test_non_directory_cache_entries_are_preserved_before_old_directory_is_deleted(self):
        for linked_cache in (False, True):
            with self.subTest(linked_cache=linked_cache):
                store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
                if linked_cache:
                    (self.legacy / "cache").symlink_to("../missing-cache")
                else:
                    (self.legacy / "cache").write_bytes(b"original cache entry")
                store.initialize_data_directory()
                directory = self.current / "legacy_backup" if linked_cache else self.current
                if linked_cache:
                    self.assertTrue((directory / "cache").is_symlink())
                    self.assertEqual((directory / "cache").resolve(), (self.home / "missing-cache").resolve())
                else:
                    self.assertEqual((directory / "cache").read_bytes(), b"original cache entry")
                self.assertFalse(self.legacy.exists())

    def test_install_failure_rolls_back_and_does_not_delete_old_directory(self):
        old, current = {"watchlist": {"old": {}}}, {"watchlist": {"current": {}}}
        store.save_file(old, CONFIG_FILE, directory=str(self.legacy))
        store.save_file(current, CONFIG_FILE, directory=str(self.current))
        original_replace = store.os.replace

        def fail_config(source, target):
            if Path(target) == self.current / CONFIG_FILE and "files" in Path(source).parts:
                raise OSError("no space")
            return original_replace(source, target)

        with patch.object(store.os, "replace", side_effect=fail_config), self.assertRaises(OSError):
            store.initialize_data_directory()
        self.assertTrue(self.legacy.exists())
        self.assertEqual(store.load_file(CONFIG_FILE), current)
        self.assertEqual(store.load_file(CONFIG_FILE, directory=str(self.legacy)), old)

    def test_target_directory_link_cannot_write_outside_program_data_directory(self):
        old = {"watchlist": {"old": {}}}
        store.save_file(old, CONFIG_FILE, directory=str(self.legacy))
        store.save_file({"old_codes": True}, CODE_LIST_FILES[0], directory=str(self.legacy / "data"))
        external = self.home / "external_codes"
        store.save_file({"external": True}, CODE_LIST_FILES[0], directory=str(external))
        self.current.mkdir(parents=True)
        (self.current / "data").symlink_to(external, target_is_directory=True)
        with self.assertRaises(OSError):
            store.initialize_data_directory()
        self.assertTrue(self.legacy.exists())
        self.assertEqual(store.load_file(CONFIG_FILE, directory=str(self.legacy)), old)
        self.assertEqual(store.load_file(CODE_LIST_FILES[0], directory=str(external)), {"external": True})
        self.assertFalse((self.current / CONFIG_FILE).exists())

    def test_cleanup_failure_keeps_migrated_configuration_and_can_retry_without_overwriting_it(self):
        store.save_file({"watchlist": {"old": {}}}, CONFIG_FILE, directory=str(self.legacy))
        original_remove = store.shutil.rmtree

        def refuse_legacy(path, *args, **kwargs):
            if Path(path) == self.legacy:
                raise PermissionError("read only")
            return original_remove(path, *args, **kwargs)

        with patch.object(store.shutil, "rmtree", side_effect=refuse_legacy), self.assertRaises(OSError):
            store.initialize_data_directory()
        self.assertTrue(self.legacy.exists())
        self.assertEqual(store.load_file(CONFIG_FILE)["watchlist"], {"old": {}})
        store.save_file({"watchlist": {"updated": {}}}, CONFIG_FILE)
        self.assertEqual(store.initialize_data_directory()["watchlist"], {"updated": {}})
        self.assertFalse(self.legacy.exists())

    def test_backup_retains_unknown_cache_version_and_does_not_replace_an_earlier_backup(self):
        store.save_file({"watchlist": {"old": {}}}, CONFIG_FILE, directory=str(self.legacy))
        store.save_file({"watchlist": {"current": {}}}, CONFIG_FILE)
        self.write_database(self.legacy / "cache", "history", version=999)
        earlier = self.current / "legacy_backup"
        earlier.mkdir()
        (earlier / "extra.txt").write_text("earlier backup")
        store.initialize_data_directory()
        self.assertEqual((earlier / "extra.txt").read_text(), "earlier backup")
        latest = self.current / "legacy_backup_1" / "cache"
        self.assertEqual(self.read_database(latest), "history")
        with closing(sqlite3.connect(latest / "history.sqlite3")) as db:
            self.assertEqual(db.execute("PRAGMA user_version").fetchone()[0], 999)
        self.assertFalse(self.legacy.exists())

    def test_custom_icon_in_old_directory_remains_available_after_cleanup(self):
        self.legacy.mkdir()
        icon = self.legacy / "icon.png"
        icon.write_bytes(b"icon data")
        store.save_file({"app_icon": "custom", "custom_icon_path": str(icon)}, CONFIG_FILE, directory=str(self.legacy))
        config = store.initialize_data_directory()
        self.assertEqual(config["custom_icon_path"], str(self.current / "icon.png"))
        self.assertEqual(Path(config["custom_icon_path"]).read_bytes(), b"icon data")
        self.assertFalse(self.legacy.exists())

    def test_absolute_internal_links_remain_valid_after_old_directory_is_deleted(self):
        store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
        (self.legacy / "extra.txt").write_text("linked data")
        (self.legacy / "link.txt").symlink_to(self.legacy / "extra.txt")
        store.initialize_data_directory()
        self.assertFalse(self.legacy.exists())
        self.assertTrue((self.current / "link.txt").is_symlink())
        self.assertEqual((self.current / "link.txt").read_text(), "linked data")

    def test_relative_internal_external_and_nested_links_keep_their_targets(self):
        pictures = self.home / "Pictures"
        pictures.mkdir()
        (pictures / "icon.png").write_bytes(b"external icon")
        for existing_new_data in (False, True):
            with self.subTest(existing_new_data=existing_new_data):
                (self.legacy / "nested" / "deep").mkdir(parents=True)
                (self.legacy / "assets").mkdir()
                (self.legacy / "assets" / "icon.png").write_bytes(b"internal icon")
                (self.legacy / "icon.png").symlink_to("../Pictures/icon.png")
                (self.legacy / "nested" / "deep" / "icon.png").symlink_to("../../../Pictures/icon.png")
                (self.legacy / "nested" / "deep" / "chain.png").symlink_to("../../icon.png")
                (self.legacy / "nested" / "internal.png").symlink_to("../assets/icon.png")
                (self.legacy / "nested" / "pictures").symlink_to("../../Pictures", target_is_directory=True)
                config = {"custom_icon_path": str(self.legacy / "icon.png")}
                store.save_file(config, CONFIG_FILE, directory=str(self.legacy))
                if existing_new_data:
                    store.save_file({**config, "new": True}, CONFIG_FILE)
                migrated = store.initialize_data_directory()
                directory = self.current / "legacy_backup" if existing_new_data else self.current
                self.assertEqual(Path(migrated["custom_icon_path"]).read_bytes(), b"external icon")
                for relative in ("icon.png", "nested/deep/icon.png", "nested/deep/chain.png", "nested/pictures/icon.png"):
                    with self.subTest(link=relative):
                        self.assertEqual((directory / relative).resolve(), (pictures / "icon.png").resolve())
                        self.assertEqual((directory / relative).read_bytes(), b"external icon")
                self.assertEqual((directory / "nested" / "internal.png").read_bytes(), b"internal icon")
                self.assertFalse(self.legacy.exists())

    def test_old_alias_to_new_directory_is_removed_without_deleting_new_data(self):
        current = {"watchlist": {"current": {}}}
        store.save_file(current, CONFIG_FILE)
        self.legacy.symlink_to(self.current, target_is_directory=True)
        self.assertEqual(store.initialize_data_directory(), current)
        self.assertFalse(os.path.lexists(self.legacy))
        self.assertEqual(store.load_file(CONFIG_FILE), current)

    def test_other_application_support_and_system_cache_directories_are_not_migrated(self):
        other = self.home / "Library" / "Application Support" / "com.sbr0574.StockWidget"
        cache = self.home / "Library" / "Caches" / "com.sbr0574.StockWidget"
        store.save_file({"watchlist": {"other": {}}}, CONFIG_FILE, directory=str(other))
        self.write_database(cache)
        self.assertEqual(store.initialize_data_directory(), {})
        self.assertFalse(self.current.exists())
        self.assertTrue((other / CONFIG_FILE).exists())
        self.assertEqual(self.read_database(cache), "history")

    def test_wal_snapshot_preserves_committed_records_before_cleanup(self):
        cache = self.legacy / "cache"
        cache.mkdir(parents=True)
        store.save_file({"watchlist": {"old": {}}}, CONFIG_FILE, directory=str(self.legacy))
        with closing(sqlite3.connect(cache / "history.sqlite3")) as db, db:
            db.execute("PRAGMA journal_mode=WAL")
            db.execute("CREATE TABLE sample (value TEXT)")
            db.execute("INSERT INTO sample VALUES ('in WAL')")
            db.commit()
            self.assertTrue((cache / "history.sqlite3-wal").exists())
            store.initialize_data_directory()
            self.assertEqual(self.read_database(self.current / "cache"), "in WAL")
            self.assertFalse((self.current / "cache" / "history.sqlite3-wal").exists())
            self.assertFalse((self.current / "cache" / "history.sqlite3-shm").exists())
            self.assertFalse(self.legacy.exists())

    def test_orphan_sidecars_are_preserved_before_cleanup(self):
        sidecars = {"history.sqlite3" + suffix: ("original" + suffix).encode()
                    for suffix in ("-wal", "-shm", "-journal")}
        store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
        cache = self.legacy / "cache"
        cache.mkdir()
        for name, contents in sidecars.items():
            (cache / name).write_bytes(contents)
        store.initialize_data_directory()
        for name, contents in sidecars.items():
            with self.subTest(file=name):
                self.assertEqual((self.current / "cache" / name).read_bytes(), contents)
        self.assertFalse(self.legacy.exists())

    def test_invalid_database_recovery_preserves_original_files_or_retains_old_directory(self):
        store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
        cache = self.legacy / "cache"
        cache.mkdir()
        original = {"history.sqlite3" + suffix: ("invalid" + suffix).encode()
                    for suffix in ("", "-wal", "-shm", "-journal")}
        for name, contents in original.items():
            (cache / name).write_bytes(contents)
        try:
            store.initialize_data_directory()
        except OSError:
            self.assertTrue(self.legacy.exists())
            self.assertFalse((self.current / CONFIG_FILE).exists())
            preserved = cache
        else:
            self.assertFalse(self.legacy.exists())
            preserved = self.current / "cache"
        for name, contents in original.items():
            with self.subTest(file=name):
                self.assertEqual((preserved / name).read_bytes(), contents)

    def test_damaged_database_fallback_preserves_main_database_and_all_sidecars(self):
        for error_code in (sqlite3.SQLITE_CORRUPT, sqlite3.SQLITE_NOTADB):
            with self.subTest(error_code=error_code):
                store.save_file({"watchlist": {}}, CONFIG_FILE, directory=str(self.legacy))
                cache = self.legacy / "cache"
                cache.mkdir()
                original = {"history.sqlite3" + suffix: ("damaged" + suffix).encode()
                            for suffix in ("", "-wal", "-shm", "-journal")}
                for name, contents in original.items():
                    (cache / name).write_bytes(contents)
                error = sqlite3.DatabaseError("damaged database")
                error.sqlite_errorcode = error_code
                with patch.object(store.sqlite3, "connect", side_effect=error):
                    store.initialize_data_directory()
                self.assertFalse(self.legacy.exists())
                directory = self.current if error_code == sqlite3.SQLITE_CORRUPT else self.current / "legacy_backup"
                for name, contents in original.items():
                    with self.subTest(file=name):
                        self.assertEqual((directory / "cache" / name).read_bytes(), contents)

    def test_invalid_configuration_is_not_replaced_or_deleted(self):
        self.legacy.mkdir()
        path = self.legacy / CONFIG_FILE
        for contents in ("{bad json", "[]"):
            with self.subTest(contents=contents):
                path.write_text(contents)
                with self.assertRaises(OSError):
                    store.initialize_data_directory()
                self.assertEqual(path.read_text(), contents)
                self.assertFalse((self.current / CONFIG_FILE).exists())

    def test_platform_defaults_use_appdata_only_on_windows(self):
        for platform, expected in (("darwin", self.current), ("linux", self.home / "StockWidget"),
                                   ("win32", self.home / "roaming" / "StockWidget")):
            with self.subTest(platform=platform), patch.object(store.sys, "platform", platform), \
                    patch.dict(os.environ, {"APPDATA": str(self.home / "roaming")}):
                self.assertEqual(store.config_paths(), str(expected))
