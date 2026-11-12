"""程序数据目录选择及缓存清理范围。"""

from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import Mock, patch

from PySide6.QtCore import QObject, QPoint, Qt, Signal, QTimer
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication, QLabel, QMessageBox
from shiboken6 import delete

from stockwidget.app import App
from stockwidget.constants import CONFIG_FILE, CODE_LIST_FILES
from stockwidget.core.config_store import data_cache_dir, load_file, save_file
from stockwidget.data.bars import BarResult
from stockwidget.ui.history import HistoryDialog
from tests.data.test_bars import A_SHARE, sample_bars
from tests.support import SettingsTestCase


class StorageApp(QObject):
    save_state_changed = Signal()
    current_config = App.current_config
    save_now = App.save_now
    change_data_directory = App.change_data_directory
    clear_history_cache = App.clear_history_cache

    def __init__(self):
        super().__init__()
        self.pending_data_directory = None
        self.save_error = ""
        self._save_error_reported = False
        self._icon_choice = None
        self._custom_icon_path = ""
        self._start_on_boot = False
        self.app_version = "1.0.0"
        self._has_update, self._latest_version = False, None
        self.code_data_state = Mock(return_value=("cached", ""))
        self.code_data_error = Mock(return_value="")
        self.tray, self.quit = Mock(), Mock()


class StorageSettingsTests(SettingsTestCase):
    def test_opacity_limits_match_sliders_saved_values_and_applied_windows(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True, opacity_pct=0,
                                           bg={"r": 0, "g": 0, "b": 0, "a": 0},
                                           chart_sync_opacity=False, chart_bg_alpha_pct=0,
                                           chart_opacity_pct=0, taskbar_opacity_pct=0)
        controls = ((dialog.ui.slider_bg_alpha, lambda: round(window.bg.alpha() / 2.55), 1),
                    (dialog.ui.slider_all_alpha, lambda: window.opacity_pct, 10),
                    (dialog.ui.slider_chart_bg_alpha, lambda: window.view_options.chart_bg_alpha_pct, 1),
                    (dialog.ui.slider_chart_all_alpha, lambda: window.view_options.chart_opacity_pct, 10),
                    (dialog.ui.taskbar_opacity_pct, lambda: window.view_options.taskbar_opacity_pct, 10))
        for slider, read_value, minimum in controls:
            with self.subTest(control=slider.objectName()):
                self.assertEqual(read_value(), minimum)
                slider.setValue(55)
                slider.setSliderDown(True)
                slider.setSliderPosition(minimum)
                self.assertEqual(read_value(), 55)
                slider.setSliderDown(False)
                self.assertEqual(read_value(), minimum)
                slider.setValue(0)
                self.assertEqual(slider.value(), minimum)
        self.assertAlmostEqual(window.windowOpacity(), .1, delta=1 / 255)
        window.set_bg_alpha_percent(0)
        window.set_window_opacity_percent(0)
        self.assertEqual(round(window.bg.alpha() / 2.55), 1)
        self.assertEqual(window.current_config()["opacity_pct"], 10)
        popup = HistoryDialog(None, "popup")
        self.addCleanup(delete, popup)
        window.history.dialog = popup
        window.history._apply_theme()
        self.assertAlmostEqual(popup.windowOpacity(), .1, delta=1 / 255)
        self.assertEqual(round(popup.palette().color(QPalette.Base).alpha() / 2.55), 1)
        self.assertEqual(popup.ui.close_button.palette().color(QPalette.ButtonText), window.fg)

    def test_taskbar_font_is_editable_after_reset_and_new_defaults_are_applied(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_rows=4,
                                           taskbar_font_size=13, taskbar_opacity_pct=90,
                                           taskbar_auto_color=False)
        window.set_taskbar_geometry(88, 192)
        dialog.ui.btn_reset_appearance.click()
        dialog.ui.btn_reset_settings.click()
        options = window.view_options
        self.assertEqual((options.taskbar_rows, options.taskbar_font_size, options.taskbar_opacity_pct), (3, 8, 50))
        self.assertTrue(options.taskbar_auto_color)
        dialog.taskbar_settings.enabled.setChecked(True)
        self.assertTrue(dialog.ui.taskbar_font_size.isEnabled())
        original = window.font.pointSize()
        dialog.ui.taskbar_font_size.setSliderDown(True)
        dialog.ui.taskbar_font_size.setSliderPosition(10)
        self.assertEqual(options.taskbar_font_size, 8)
        dialog.ui.taskbar_font_size.setSliderDown(False)
        self.assertEqual(window.get_taskbar_appearance()[0].pointSize(), 10)
        self.assertEqual(window.font.pointSize(), original)
        self.assertEqual(window.current_config()["taskbar_font_size"], 10)

    def test_taskbar_font_limit_tracks_rows_and_dpi_without_losing_larger_preferences(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_rows=1)
        window.set_taskbar_geometry(88, 192)
        dialog.ui.taskbar_font_size.setValue(12)
        for rows, maximum in ((2, 15), (3, 10), (4, 7), (1, 15)):
            with self.subTest(rows=rows):
                dialog.ui.taskbar_rows.setValue(rows)
                self.assertEqual(dialog.ui.taskbar_font_size.maximum(), maximum)
                self.assertEqual(dialog.ui.taskbar_font_size.value(), min(12, maximum))
                self.assertEqual(dialog.ui.taskbar_font_size_label.text(), f"{min(12, maximum)} pt")
                self.assertEqual(window.current_config()["taskbar_font_size"], 12)
        dialog.ui.taskbar_rows.setValue(3)
        window.set_taskbar_geometry(44, 96)
        self.assertEqual(dialog.ui.taskbar_font_size.maximum(), 8)
        self.assertEqual(window.view_options.taskbar_font_size, 12)
        window.set_font_family("Arial")
        self.assertEqual(window.get_taskbar_appearance()[0].family(), window.font.family())

    def test_taskbar_size_control_and_saved_preferences_stay_within_five_to_fifteen_points(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_rows=1, taskbar_font_size=30)
        window.set_taskbar_geometry(88, 192)
        self.assertEqual(window.get_taskbar_appearance()[0].pointSize(), 15)
        for selected, expected in ((0, 5), (5, 5), (15, 15), (30, 15)):
            with self.subTest(selected=selected):
                dialog.ui.taskbar_font_size.setValue(selected)
                self.assertEqual(window.get_taskbar_appearance()[0].pointSize(), expected)
                self.assertEqual(window.current_config()["taskbar_font_size"], expected)

    def test_write_failure_does_not_block_settings_editing_or_exit(self):
        owner = StorageApp()
        self.addCleanup(delete, owner)
        with patch("stockwidget.app.save_file", side_effect=PermissionError()), \
                patch("stockwidget.ui.settings.dialog.QMessageBox.critical") as critical:
            dialog, window = self._make_dialog(app=owner)
            owner.win = window
            window.set_on_change(owner.save_now)
            self.assertEqual(owner.save_error, "")
            self.assertTrue(dialog.ui.btn_icon_default.isChecked())
            dialog.show()
            window.set_font_size(12)
            self.qt_app.processEvents()
            self.assertEqual(window.font.pointSize(), 12)
            self.assertIn("尚未保存", dialog.ui.data_directory_status.text())
            owner.tray.showMessage.assert_called_once()
            App.quit_app(owner)
            owner.quit.assert_called_once()
            critical.assert_not_called()

    def test_directory_chooser_shows_pending_restart_and_handles_copy_errors(self):
        owner = StorageApp()
        self.addCleanup(delete, owner)
        dialog, owner.win = self._make_dialog(app=owner)
        target = str(Path(owner.win.history.cache.directory).parent.parent / "chosen")

        def select(path):
            owner.pending_data_directory = path
            owner.save_state_changed.emit()

        with patch.object(owner, "change_data_directory", side_effect=select) as change, \
                patch("stockwidget.ui.settings.groups.QFileDialog.getExistingDirectory", return_value=target):
            dialog.ui.btn_data_directory.click()
            change.assert_called_once_with(target)
        self.assertEqual(dialog.ui.data_directory.text(), target)
        self.assertIn("重启", dialog.ui.data_directory_status.text())
        original = owner.win.history.cache.directory
        with patch.object(owner, "change_data_directory", side_effect=PermissionError()), \
                patch("stockwidget.ui.settings.groups.QFileDialog.getExistingDirectory", return_value="other"), \
                patch("stockwidget.ui.settings.groups.QMessageBox.warning") as warning:
            dialog.ui.btn_data_directory.click()
            warning.assert_called_once()
        self.assertEqual(dialog.ui.data_directory.text(), target)
        self.assertEqual(owner.win.history.cache.directory, original)

    def test_storage_actions_and_destructive_buttons_render_in_both_themes(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True)
        dialog.show()
        for mode in ("light", "dark"):
            window.set_view_options(color_mode=mode)
            dialog.ui.settings_pages.setCurrentIndex(5)
            self.qt_app.processEvents()
            dialog.ui.taskbar_scroll.ensureWidgetVisible(dialog.ui.taskbar_font_size)
            self.qt_app.processEvents()
            self.assertEqual(dialog.ui.taskbar_scroll.horizontalScrollBar().maximum(), 0)
            self.assertTrue(dialog.ui.taskbar_scroll.viewport().rect().contains(
                dialog.ui.taskbar_font_size.mapTo(dialog.ui.taskbar_scroll.viewport(),
                                                 dialog.ui.taskbar_font_size.rect().bottomRight())))
            self.assertEqual(dialog.ui.taskbar_font_size.width(), dialog.ui.taskbar_opacity_pct.width())
            self.assertEqual(dialog.ui.taskbar_font_size.x(), dialog.ui.taskbar_opacity_pct.x())
            dialog.ui.settings_pages.setCurrentWidget(dialog.ui.general)
            dialog.ui.general_scroll.ensureWidgetVisible(dialog.ui.btn_clear_cache)
            self.qt_app.processEvents()
            group = dialog.ui.data_directory_settings
            positions = [button.mapTo(group, QPoint()) for button in (dialog.ui.btn_open_cache_dir, dialog.ui.btn_clear_cache)]
            self.assertEqual(positions[0].y(), positions[1].y())
            self.assertGreaterEqual(positions[0].y(), 25)
            self.assertGreaterEqual(positions[0].x(), 10)
            self.assertEqual(dialog.ui.general_scroll.horizontalScrollBar().maximum(), 0)
            for page, button in ((dialog.ui.about, dialog.ui.btn_reset_settings),
                                 (dialog.ui.floating, dialog.ui.btn_reset_appearance),
                                 (dialog.ui.watchlist, dialog.ui.btn_clear_watchlist)):
                dialog.ui.settings_pages.setCurrentWidget(page)
                self.qt_app.processEvents()
                image = button.grab().toImage()
                red_pixels = sum(1 for y in range(image.height()) for x in range(image.width())
                                 if (color := image.pixelColor(x, y)).red() > 200
                                 and color.green() < 110 and color.blue() < 110)
                self.assertGreater(red_pixels, 4, (mode, button.objectName()))

    def test_update_prompt_renders_an_accessible_download_link(self):
        dialog, _ = self._make_dialog()
        observed = {}

        def inspect_prompt():
            box = QApplication.activeModalWidget()
            if not isinstance(box, QMessageBox):
                return
            label = box.findChild(QLabel, "qt_msgbox_label")
            observed.update(text=label.text(), external=label.openExternalLinks(), flags=label.textInteractionFlags())
            box.accept()

        QTimer.singleShot(0, inspect_prompt)
        dialog._on_update_check_finished((True, "9.9.9"))
        self.assertIn('href="https://github.com/sbr0574/StockWidget/releases"', observed["text"])
        self.assertTrue(observed["external"])
        self.assertTrue(observed["flags"] & Qt.LinksAccessibleByMouse)

    def test_open_data_directory_targets_configuration_parent_and_reset_keeps_cache_location(self):
        dialog, window = self._make_dialog()
        root = window.history.cache.directory.parent
        with patch("stockwidget.ui.settings.groups.QDesktopServices.openUrl", return_value=True) as open_url:
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
        history.get(A_SHARE, "daily")
        before = load_file(CONFIG_FILE)
        dialog.ui.btn_clear_cache.click()
        self.assertFalse((history.directory / "history.sqlite3").exists())
        self.assertEqual(load_file(CONFIG_FILE), before)
        self.assertEqual(load_file(CODE_LIST_FILES[0], directory=data_cache_dir()), payload)
        self.assertEqual(window.watchlist, before["watchlist"])

    def test_clear_cache_also_removes_the_copy_waiting_for_restart(self):
        owner = StorageApp()
        self.addCleanup(delete, owner)
        dialog, window = self._make_dialog(watchlist={"sh600519": {"checked": True}}, app=owner)
        owner.win = window
        history = window.history.cache
        history.clock = lambda: datetime(2026, 9, 30, 2, tzinfo=timezone.utc)
        history.fetcher = Mock(return_value=BarResult(sample_bars(), "sina"))
        history.get(A_SHARE, "daily")
        target = history.directory.parent.parent / "chosen"
        with patch("stockwidget.core.config_store._save_location"):
            owner.change_data_directory(str(target))
        self.assertTrue((target / "cache" / "history.sqlite3").exists())
        dialog.ui.btn_clear_cache.click()
        self.assertFalse(history.database.exists())
        self.assertFalse((target / "cache" / "history.sqlite3").exists())
        self.assertEqual(load_file(CONFIG_FILE, directory=str(target))["watchlist"], window.watchlist)
