"""固定程序数据路径及缓存清理范围。"""

from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import Mock, patch

from stockwidget.constants import CONFIG_FILE, CODE_LIST_FILES
from stockwidget.core.config_store import data_cache_dir, load_file, save_file
from stockwidget.data.bars import BarResult
from tests.data.test_bars import A_SHARE, sample_bars
from tests.support import SettingsTestCase


class StorageSettingsTests(SettingsTestCase):
    def test_open_data_directory_targets_configuration_parent_and_reset_keeps_cache_location(self):
        dialog, window = self._make_dialog()
        root = window.history.cache.directory.parent
        with patch("stockwidget.ui.settings.dialog.QDesktopServices.openUrl", return_value=True) as open_url:
            dialog.ui.btn_open_cache_dir.click()
        self.assertEqual(Path(open_url.call_args.args[0].toLocalFile()), root)
        self.assertTrue(root.is_dir())
        window.reset_settings()
        self.assertEqual(window.history.cache.directory, root / "cache")

    def test_clear_cache_keeps_configuration_watchlist_and_market_codes(self):
        dialog, window = self._make_dialog(watchlist={"sh600519": {"checked": True, "cost": 42}})
        save_file(window.current_config(), CONFIG_FILE)
        payload = {"codes": {"sh600519": {"name": "贵州茅台"}}}
        save_file(payload, CODE_LIST_FILES[0], directory=data_cache_dir())
        history = window.history.cache
        history.clock = lambda: datetime(2026, 9, 30, 2, tzinfo=timezone.utc)
        history.fetcher = Mock(return_value=BarResult(sample_bars(), "sina"))
        history.get(A_SHARE, "daily", "sina")
        before = load_file(CONFIG_FILE)
        dialog.ui.btn_clear_cache.click()
        self.assertFalse((history.directory / "history.sqlite3").exists())
        self.assertEqual(load_file(CONFIG_FILE), before)
        self.assertEqual(load_file(CODE_LIST_FILES[0], directory=data_cache_dir()), payload)
        self.assertEqual(window.watchlist, before["watchlist"])
