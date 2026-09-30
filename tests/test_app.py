"""应用装配、代码同步调度、图标、版本检查和资源完整性。"""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch
from xml.etree import ElementTree as ET
import os
import tempfile
import unittest

from PySide6.QtGui import QColor, QIcon, QPixmap
import requests

from stockwidget.app import App, _load_custom_icon
from stockwidget.constants import CONFIG_FILE, CODE_LIST_FILES
from stockwidget.data import update_check

from tests.support import QtTestCase


class AppIconTests(QtTestCase):

    def _write_icon(self, directory: str) -> str:
        path = os.path.join(directory, "custom.png")
        pixmap = QPixmap(24, 24)
        pixmap.fill(QColor("#3578e5"))
        self.assertTrue(pixmap.save(path))
        return path

    def test_load_custom_icon_accepts_image_and_rejects_invalid_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write_icon(temp_dir)
            normalized, icon = _load_custom_icon(path)

            self.assertEqual(normalized, os.path.abspath(path))
            self.assertFalse(icon.isNull())

        normalized, icon = _load_custom_icon(path)
        self.assertEqual(normalized, "")
        self.assertTrue(icon.isNull())

    def test_set_custom_icon_applies_it_immediately(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write_icon(temp_dir)
            app = SimpleNamespace(
                _icon_choice="default",
                _custom_icon_path="",
                setWindowIcon=Mock(),
                tray=Mock(),
            )

            self.assertTrue(App.set_custom_icon(app, path))

        self.assertEqual(app._icon_choice, "custom")
        self.assertEqual(app._custom_icon_path, os.path.abspath(path))
        app.setWindowIcon.assert_called_once()
        app.tray.setIcon.assert_called_once()

    def test_save_now_writes_custom_choice_and_path(self):
        app = SimpleNamespace(
            win=Mock(),
            _icon_choice="custom",
            _custom_icon_path=r"C:\icons\mine.png",
            _start_on_boot=False,
        )
        app.win.current_config.return_value = {"watchlist": {}}

        with patch("stockwidget.app.save_file") as save:
            App.save_now(app)

        saved_config, file_name = save.call_args.args
        self.assertEqual(file_name, CONFIG_FILE)
        self.assertEqual(saved_config["app_icon"], "custom")
        self.assertEqual(saved_config["custom_icon_path"], r"C:\icons\mine.png")

    def test_invalid_custom_path_falls_back_to_default(self):
        app = SimpleNamespace(
            _icon_choice="custom",
            _custom_icon_path=r"C:\missing\icon.png",
            find_icon=Mock(return_value=QIcon()),
            setWindowIcon=Mock(),
            tray=Mock(),
        )

        App.set_app_icon(app, "custom")

        self.assertEqual(app._icon_choice, "default")
        self.assertEqual(app._custom_icon_path, "")
        app.setWindowIcon.assert_called_once()
        app.tray.setIcon.assert_called_once()


class AppCodeRefreshTests(unittest.TestCase):
    def test_quit_uses_qt_event_loop_shutdown(self):
        app = SimpleNamespace(tray=Mock(), save_now=Mock(), quit=Mock())
        App.quit_app(app)
        app.tray.hide.assert_called_once_with()
        app.save_now.assert_called_once_with()
        app.quit.assert_called_once_with()

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


ROOT = Path(__file__).parents[1]


class ResourceTests(unittest.TestCase):
    def test_us_alias_cache_is_not_a_client_download(self):
        self.assertNotIn("cache_us_cn_aliases.json", CODE_LIST_FILES)

    def test_packaged_resources_exist_and_preserve_virtual_names(self):
        root = ET.parse(ROOT / "resources" / "resources.qrc").getroot()
        aliases = {}
        for node in root.iter("file"):
            path = ROOT / "resources" / node.text
            self.assertTrue(path.is_file(), path)
            self.assertIn(path.parent.name, ("icons", "data"))
            self.assertEqual(node.attrib["alias"], path.name)
            aliases[node.attrib["alias"]] = path
        self.assertTrue(set(CODE_LIST_FILES).issubset(aliases))
        self.assertIn("StockWidget.ico", aliases)


class UpdateCheckTests(unittest.TestCase):
    def test_failed_release_sources_return_specific_diagnostics(self):
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(
            requests.ConnectTimeout(), requests.exceptions.ProxyError()
        )):
            self.assertEqual(update_check.get_update_info("1.0.0", errors=errors), (False, None))
        self.assertEqual(errors, ["GitHub：连接超时", "Gitee：代理请求失败"])

    def test_release_fallback_succeeds_after_a_request_error(self):
        response = Mock(status_code=200)
        response.json.return_value = {"tag_name": "v1.5.0"}
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(requests.ReadTimeout(), response)):
            self.assertEqual(update_check.get_update_info("1.0.0", errors=errors), (True, "1.5.0"))
        self.assertEqual(errors, ["GitHub：读取超时"])

    def test_http_and_response_errors_are_reported_without_changing_default_return(self):
        response = requests.Response()
        response.status_code = 403
        invalid = Mock(status_code=200)
        invalid.json.side_effect = requests.exceptions.JSONDecodeError("bad", "", 0)
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(response, invalid)):
            self.assertIsNone(update_check.get_latest_release(errors=errors))
        self.assertEqual(errors, ["GitHub：HTTP 403：访问被拒绝", "Gitee：响应不是有效 JSON"])
        with patch.object(update_check.requests, "get", side_effect=requests.ConnectTimeout()):
            self.assertIsNone(update_check.get_latest_release())

    def test_release_check_falls_back_to_gitee(self):
        with patch.object(
            update_check,
            "_release_version",
            side_effect=(None, "1.5.0"),
        ) as fetch:
            result = update_check.get_latest_release()

        self.assertEqual(result, "1.5.0")
        self.assertIn("api.github.com", fetch.call_args_list[0].args[0])
        self.assertIn("gitee.com", fetch.call_args_list[1].args[0])

    def test_project_links_need_no_source_constants(self):
        self.assertIn("github.com", update_check.project_links()["project"])
        self.assertIn(
            "gitee.com",
            update_check.project_links(use_gitee=True)["project"],
        )
