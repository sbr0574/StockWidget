# -*- coding: utf-8 -*-
"""应用层代码列表初始化和重试调度测试。"""

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from stockwidget.app import App


class AppCodeRefreshTests(unittest.TestCase):
    def test_schedule_refresh_uses_milliseconds(self):
        app = SimpleNamespace(_codes_retry_timer=Mock())

        App._schedule_codes_refresh(app, 1800)

        app._codes_retry_timer.start.assert_called_once_with(1800 * 1000)

    def test_schedule_refresh_clamps_non_positive_delay(self):
        app = SimpleNamespace(_codes_retry_timer=Mock())

        App._schedule_codes_refresh(app, 0)

        app._codes_retry_timer.start.assert_called_once_with(1000)

    def test_codes_loaded_updates_window_before_remote_sync(self):
        settings = Mock()
        settings.isVisible.return_value = True
        app = SimpleNamespace(
            _codes_local_ready=False,
            win=Mock(),
            settings_dlg=settings,
            save_now=Mock(),
            _start_codes_refresh=Mock(),
        )
        codes = {"sh600000": {"code": "600000", "market": "sh"}}

        App._on_codes_loaded(app, codes)

        self.assertTrue(app._codes_local_ready)
        app.win.set_codes_list.assert_called_once_with(codes)
        app.save_now.assert_called_once_with()
        settings.refresh_data_state.assert_called_once_with()
        settings.refresh_code_search.assert_called_once_with()
        app._start_codes_refresh.assert_called_once_with()

    def test_refresh_marks_cached_before_starting_worker(self):
        settings = Mock()
        settings.isVisible.return_value = True
        app = SimpleNamespace(
            _codes_refresh_running=False,
            _codes_local_ready=True,
            code_manager=Mock(),
            settings_dlg=settings,
            codes_refresh_finished=Mock(),
        )

        with patch("stockwidget.app.threading.Thread") as thread:
            App._start_codes_refresh(app)

        self.assertTrue(app._codes_refresh_running)
        app.code_manager.begin_remote_check.assert_called_once_with()
        settings.refresh_data_state.assert_called_once_with()
        thread.assert_called_once()
        thread.return_value.start.assert_called_once_with()

    def test_refresh_finished_updates_window_and_schedules_retry(self):
        manager = Mock()
        manager.codes.return_value = {"sh600000": {}}
        settings = Mock()
        settings.isVisible.return_value = True
        app = SimpleNamespace(
            _codes_refresh_running=True,
            code_manager=manager,
            win=Mock(),
            settings_dlg=settings,
            _schedule_codes_refresh=Mock(),
        )

        App._on_codes_refresh_finished(app, {"retry_seconds": 1800})

        self.assertFalse(app._codes_refresh_running)
        app.win.set_codes_list.assert_called_once_with(manager.codes.return_value)
        settings.refresh_data_state.assert_called_once_with()
        settings.refresh_code_search.assert_called_once_with()
        app._schedule_codes_refresh.assert_called_once_with(1800)


if __name__ == "__main__":
    unittest.main()
