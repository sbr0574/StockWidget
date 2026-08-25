# -*- coding: utf-8 -*-
"""应用层市场代码更新时间与重试定时器测试。"""

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from stockwidget import app as app_module
from stockwidget.data.code_lists import CODES_RETRY_SECONDS


class TestAppCodeRefresh(unittest.TestCase):
    def test_weekend_with_no_stale_files_stops_refresh_timer(self):
        app = SimpleNamespace(
            app_name="StockWidgetTest",
            _codes_retry_timer=Mock(),
            _start_codes_refresh=Mock(),
        )
        with patch.object(app_module, "stale_code_files", return_value=()):
            app_module.App._schedule_codes_refresh(app)

        app._codes_retry_timer.stop.assert_called_once_with()
        app._start_codes_refresh.assert_not_called()

    def test_stale_codes_before_nine_are_scheduled(self):
        app = SimpleNamespace(
            app_name="StockWidgetTest",
            _codes_retry_timer=Mock(),
            _start_codes_refresh=Mock(),
        )
        with (
            patch.object(app_module, "stale_code_files", return_value=("stock.json",)),
            patch.object(app_module, "code_refresh_delay_seconds", return_value=1800),
        ):
            app_module.App._schedule_codes_refresh(app)

        app._codes_retry_timer.start.assert_called_once_with(1800 * 1000)
        app._start_codes_refresh.assert_not_called()

    def test_stale_codes_after_nine_refresh_immediately(self):
        app = SimpleNamespace(
            app_name="StockWidgetTest",
            _codes_retry_timer=Mock(),
            _start_codes_refresh=Mock(),
        )
        with (
            patch.object(app_module, "stale_code_files", return_value=("stock.json",)),
            patch.object(app_module, "code_refresh_delay_seconds", return_value=0),
        ):
            app_module.App._schedule_codes_refresh(app)

        app._start_codes_refresh.assert_called_once_with()

    def test_missing_current_remote_data_retries_after_half_hour(self):
        app = SimpleNamespace(
            app_name="StockWidgetTest",
            _codes_refresh_running=True,
            _remote_source="github",
            _codes_retry_timer=Mock(),
            _schedule_codes_refresh=Mock(),
            settings_dlg=None,
            win=Mock(),
        )
        with (
            patch.object(app_module, "load_best_codes", return_value={"sh1": {}}),
            patch.object(app_module, "stale_code_files", return_value=("stock.json",)),
        ):
            app_module.App._on_codes_refresh_finished(
                app, {"source": "github", "updated": ()}
            )

        app._schedule_codes_refresh.assert_called_once_with(CODES_RETRY_SECONDS)
        self.assertEqual(CODES_RETRY_SECONDS, 30 * 60)


if __name__ == "__main__":
    unittest.main()
