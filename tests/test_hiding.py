"""Hiding must preserve location, honor daily schedules and use quote time."""

from datetime import datetime, timedelta
import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QTime
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QScrollArea
from shiboken6 import delete

from stockwidget.core.hide_rules import (
    QUOTE_TIMEZONE, all_quotes_stale, due_hide_times, normalize_hide_times, quote_timestamp,
)
from stockwidget.ui.settings_dialog import SettingsDialog
from stockwidget.ui.taskbar import taskbar_message, render_taskbar
from stockwidget.ui.taskbar import TaskbarController
from stockwidget.platform.taskbar import TaskbarArea
from stockwidget.ui.widget import FloatLabel


NOW = datetime(2026, 9, 30, 15, 0, 0, tzinfo=QUOTE_TIMEZONE)
CODES = {"sh600000": {"market": "sh", "code": "600000", "checked": True},
         "sz000001": {"market": "sz", "code": "000001", "checked": True}}


def entry(age=31, *, timestamp=False):
    updated = NOW - timedelta(seconds=age)
    result = {"name": "测试", "current_price": 10, "prev_close": 9,
              "opening_price": 9, "high_price": 11, "low_price": 8,
              "deals_vol": 100, "deals_amt": 1000,
              "purchaser_price": [9] * 5, "seller_price": [10] * 5,
              "purchaser_vol": [100] * 5, "seller_vol": [100] * 5}
    if timestamp:
        result["timestamp"] = updated.timestamp()
    else:
        result.update(date=updated.strftime("%Y/%m/%d"), time=updated.strftime("%H:%M:%S"))
    return result


class HideRuleTests(unittest.TestCase):
    def test_normalize_unique_minutes_and_limit_three(self):
        self.assertEqual(normalize_hide_times([None, "25:00", "9:05", "09:05", "15:00",
                                              "00:00", "20:00", "10:00:01"]),
                         ["00:00", "09:05", "15:00"])
        for invalid in (None, "15:00", 15, {}):
            self.assertEqual(normalize_hide_times(invalid), [])

    def test_sina_beijing_time_and_eastmoney_epoch_ignore_computer_timezone(self):
        self.assertEqual(quote_timestamp(entry()), NOW.timestamp() - 31)
        self.assertEqual(quote_timestamp(entry(timestamp=True)), NOW.timestamp() - 31)
        self.assertEqual(quote_timestamp({"date": "2026-09-30", "time": "07:00:00+00:00"}),
                         NOW.timestamp())
        for value in (None, 0, "-", "bad", float("nan"), float("inf"), 1e99):
            self.assertIsNone(quote_timestamp({"timestamp": value}))
        for value in ({}, {"date": "2026-09-30"}, {"date": "bad", "time": "15:00:00"}):
            self.assertIsNone(quote_timestamp(value))

    def test_strict_threshold_all_selected_missing_future_and_previous_day(self):
        for second in (30, 29, -10):
            self.assertFalse(all_quotes_stale({"a": entry(), "b": entry(second)}, ["a", "b"],
                                             NOW.timestamp()))
        self.assertTrue(all_quotes_stale({"a": entry(31), "b": entry(86400, timestamp=True)},
                                        ["a", "b"], NOW.timestamp()))
        for data in ({}, {"a": entry()}, {"a": entry(), "b": {}},
                     {"a": entry(), "b": {"date": "2026-09-30", "time": "invalid"}}):
            self.assertFalse(all_quotes_stale(data, ["a", "b"], NOW.timestamp()))
        self.assertFalse(all_quotes_stale({}, [], NOW.timestamp()))

    def test_schedule_minute_delayed_ticks_midnight_and_clock_backwards(self):
        self.assertEqual(due_hide_times(["15:00"], NOW - timedelta(seconds=1), NOW), ["15:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW, NOW + timedelta(seconds=59)), ["15:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW - timedelta(seconds=1),
                                       NOW + timedelta(minutes=2)), ["15:00"])
        midnight = NOW.replace(hour=0)
        self.assertEqual(due_hide_times(["00:00", "23:59"], midnight - timedelta(seconds=1),
                                       midnight), ["00:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW + timedelta(hours=1),
                                       NOW - timedelta(hours=1)), [])


class HidingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self.enterContext(patch.object(FloatLabel, "_refresh_from_function"))
        self.enterContext(patch("stockwidget.ui.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.widget.apply_click_through"))
        self.enterContext(patch("stockwidget.ui.hide_controller.time.time", return_value=NOW.timestamp()))
        self.monotonic = self.enterContext(patch("stockwidget.ui.hide_controller.time.monotonic", return_value=100))
        self.win = FloatLabel({"watchlist": CODES, "taskbar_enabled": True,
                               "scheduled_hide_times": ["15:00", "16:00"]}, CODES)
        self.win.show()
        self.app.processEvents()

    def tearDown(self):
        delete(self.win)
        self.app.processEvents()

    def refresh(self, data=None):
        self.win._process_data((True, data if data is not None else {code: entry() for code in CODES}, None))

    def restore_and_countdown(self):
        self.win.hide_widget()
        self.win.toggle_win()
        self.win.set_hide_options(auto_hide_enabled=True)
        self.refresh()

    def test_initial_automatic_hide_all_surfaces_restore_mode_and_position(self):
        for mode in ("float", "taskbar", "both"):
            for source in ("sina", "eastmoney"):
                with self.subTest(mode=mode, source=source):
                    self.win.hide_controller._manual_show = False
                    self.win.widget_visible = True
                    self.win.display_mode = mode
                    self.win.set_data_source(source)
                    self.win.move(120, 170)
                    self.win.set_hide_options(auto_hide_enabled=True)
                    notices = []
                    handler = lambda: notices.append(self.win.widget_visible)
                    self.win.widget_visibility_changed.connect(handler)
                    self.refresh({code: entry(timestamp=source == "eastmoney") for code in CODES})
                    self.assertFalse(self.win.widget_visible)
                    self.assertIn(False, notices)
                    self.assertTrue(self.win.isHidden())
                    self.assertEqual((self.win.pos(), self.win.display_mode), (QPoint(120, 170), mode))
                    self.win.toggle_win()
                    self.assertTrue(self.win.widget_visible)
                    self.assertEqual((self.win.pos(), self.win.display_mode), (QPoint(120, 170), mode))
                    self.win.widget_visibility_changed.disconnect(handler)

    def test_manual_restore_countdown_five_seconds_not_restarted_by_refresh(self):
        self.restore_and_countdown()
        self.assertTrue(self.win.widget_visible)
        self.assertIn("5秒", self.win.hide_notice.text())
        self.assertTrue(self.win.hide_controller.countdown_timer.isActive())
        self.monotonic.return_value = 102
        self.refresh()
        self.win.hide_controller.check_countdown()
        self.assertIn("3秒", self.win.hide_notice.text())
        self.assertEqual(self.win.model.rowCount(), 2)
        self.assertEqual(taskbar_message(self.win), self.win.hide_notice.text())
        for dpi in (96, 144, 192):
            self.assertFalse(render_taskbar(self.win, round(48 * dpi / 96), dpi=dpi).isNull())
        self.monotonic.return_value = 104.999
        self.win.hide_controller.check_countdown()
        self.assertTrue(self.win.widget_visible)
        self.monotonic.return_value = 105
        self.win.hide_controller.check_countdown()
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.assertFalse(self.win.hide_controller.countdown_timer.isActive())
        self.win.toggle_win()
        self.refresh()
        self.assertIn("5秒", self.win.hide_notice.text())

    def test_refresh_fresh_incomplete_missing_time_or_failure_cancels_countdown(self):
        missing_time = {key: value for key, value in entry().items() if key not in ("date", "time")}
        for payload in ((True, {"sh600000": entry(), "sz000001": entry(30)}, None),
                        (True, {"sh600000": entry()}, None),
                        (True, {"sh600000": entry(), "sz000001": missing_time}, None),
                        (True, {}, None), (False, None, "网络请求失败")):
            with self.subTest(payload=payload):
                self.restore_and_countdown()
                self.win._process_data(payload)
                self.assertTrue(self.win.widget_visible)
                self.assertEqual(self.win.hide_notice.text(), "")
                self.assertFalse(self.win.hide_controller.countdown_timer.isActive())

    def test_disabling_auto_switching_source_and_watchlist_cancel_pending_hide(self):
        self.restore_and_countdown()
        self.win.set_hide_options(auto_hide_enabled=False)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00"])
        self.restore_and_countdown()
        generation = self.win._quote_generation
        self.win.set_data_source("eastmoney")
        self.assertEqual(self.win.hide_notice.text(), "")
        self.win._process_data((True, {code: entry() for code in CODES}, None, generation))
        self.assertIsNone(self.win.hide_controller._deadline)
        self.refresh()
        self.assertIsNotNone(self.win.hide_controller._deadline)
        self.win.set_watchlist({})
        self.assertIsNone(self.win.hide_controller._deadline)
        self.win.reset_settings()
        self.assertFalse(self.win.auto_hide_enabled)
        self.assertFalse(self.win.scheduled_hide_enabled)
        self.assertEqual(self.win.scheduled_hide_times, [])

    def test_scheduled_hide_overrides_countdown_once_a_day_even_when_hidden(self):
        self.restore_and_countdown()
        self.win.set_hide_options(scheduled_hide_enabled=True)
        controller = self.win.hide_controller
        controller._last_check = NOW - timedelta(seconds=1)
        controller.check_schedule(NOW)
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.win.toggle_win()
        controller.check_schedule(NOW + timedelta(seconds=20))
        self.assertTrue(self.win.widget_visible)
        controller.check_schedule(NOW + timedelta(days=1))
        self.assertFalse(self.win.widget_visible)
        controller.check_schedule(NOW + timedelta(days=1, hours=1))
        self.win.toggle_win()
        controller.check_schedule(NOW + timedelta(days=1, hours=1, seconds=10))
        self.assertTrue(self.win.widget_visible)

    def test_schedule_timer_independent_of_refresh_and_preserves_config(self):
        self.win.set_hide_options(scheduled_hide_enabled=True, auto_hide_enabled=True)
        self.win.hide_widget()
        self.assertFalse(self.win.timer.isActive())
        self.assertTrue(self.win.hide_controller.schedule_timer.isActive())
        saved = self.win.current_config()
        other = FloatLabel(saved, CODES)
        try:
            self.assertTrue(other.scheduled_hide_enabled)
            self.assertTrue(other.auto_hide_enabled)
            self.assertEqual(other.scheduled_hide_times, ["15:00", "16:00"])
        finally:
            delete(other)
        self.win.set_hide_options(scheduled_hide_enabled=False)
        self.assertFalse(self.win.hide_controller.schedule_timer.isActive())
        self.assertTrue(self.win.auto_hide_enabled)
        self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00"])

    def test_display_mode_menu_restore_also_gets_countdown(self):
        self.win.set_hide_options(auto_hide_enabled=True)
        self.win.hide_widget()
        self.win.set_display_mode("both")
        self.refresh()
        self.assertTrue(self.win.widget_visible)
        self.assertIn("5秒", self.win.hide_notice.text())

    def test_native_taskbar_hides_and_restores_with_auto_and_scheduled_rules(self):
        self.enterContext(patch("stockwidget.ui.taskbar.QApplication.platformName", return_value="windows"))
        self.enterContext(patch("stockwidget.ui.taskbar.find_taskbar", return_value=TaskbarArea(1, 1280, 48, 900)))
        self.enterContext(patch("stockwidget.ui.widget.sys.platform", "win32"))
        native = Mock()
        self.enterContext(patch("stockwidget.ui.taskbar.NativeTaskbarWindow", return_value=native))
        controller = TaskbarController(self.win, Mock())
        try:
            self.win.set_hide_options(auto_hide_enabled=True, scheduled_hide_enabled=True)
            for mode in ("taskbar", "both"):
                self.win.set_display_mode(mode)
                self.win.hide_controller._manual_show = False
                controller.apply_mode()
                self.assertTrue(controller._active)
                self.refresh()
                self.assertFalse(controller._active)
                native.hide.assert_called()
                self.win.toggle_win()
                self.refresh()
                self.assertTrue(controller._active)
                self.assertIn("5秒", self.win.hide_notice.text())
                self.win.hide_controller._last_check = NOW - timedelta(seconds=1)
                self.win.hide_controller._fired.clear()
                self.win.hide_controller.check_schedule(NOW)
                self.assertFalse(controller._active)
                self.assertFalse(self.win.widget_visible)
                self.assertEqual(self.win.display_mode, mode)
                self.win.toggle_win()
                self.assertTrue(controller._active)
        finally:
            controller.close()
            delete(controller)

    def test_scheduled_hide_cancels_drag_and_restores_original_position(self):
        origin = self.win.pos()
        start = origin + QPoint(5, 5)
        self.win.begin_drag(start)
        self.win.move_drag(start + QPoint(40, 25))
        self.assertNotEqual(self.win.pos(), origin)
        self.win.set_hide_options(scheduled_hide_enabled=True)
        self.win.hide_controller._last_check = NOW - timedelta(seconds=1)
        self.win.hide_controller.check_schedule(NOW)
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.pos(), origin)
        self.assertIsNone(self.win._drag_pos)
        self.win.toggle_win()
        self.assertEqual(self.win.pos(), origin)

    def test_settings_add_delete_limit_duplicates_toggle_theme_and_layout(self):
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(self.win, self.win)
        original = self.app.palette()
        try:
            dialog.show()
            dialog.ui.tab_widget.setCurrentWidget(dialog.ui.functions)
            dialog.ui.gb_scheduled_hide.setChecked(True)
            for hour in (17, 18):
                dialog.ui.hide_time_edit.setTime(QTime(hour, 0))
                dialog.ui.btn_add_hide_time.click()
            self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00", "17:00"])
            self.assertFalse(dialog.ui.btn_add_hide_time.isEnabled())
            dialog.ui.list_hide_times.setCurrentRow(1)
            dialog.ui.btn_del_hide_time.click()
            self.assertEqual(self.win.scheduled_hide_times, ["15:00", "17:00"])
            self.assertTrue(dialog.ui.btn_add_hide_time.isEnabled())
            dialog.ui.hide_time_edit.setTime(QTime(15, 0))
            self.assertFalse(dialog.ui.btn_add_hide_time.isEnabled())
            dialog.ui.cb_auto_hide.setChecked(True)
            self.assertTrue(self.win.auto_hide_enabled)
            self.win.set_data_source("eastmoney")
            self.assertTrue(dialog.ui.cb_auto_hide.isEnabled())
            for enabled in (False, True):
                dialog.ui.gb_scheduled_hide.setChecked(enabled)
                for foreground, background in (("#eeeeee", "#222222"), ("#222222", "#eeeeee")):
                    palette = QPalette(original)
                    palette.setColor(QPalette.WindowText, QColor(foreground))
                    palette.setColor(QPalette.Window, QColor(background))
                    self.app.setPalette(palette)
                    self.app.processEvents()
                    self.assertEqual(dialog.ui.hide_time_edit.isEnabled(), enabled)
                    self.assertEqual(dialog.ui.list_hide_times.isEnabled(), enabled)
            page = dialog.ui.functions
            self.assertFalse(page.findChildren(QScrollArea))
            for control in (dialog.ui.gb_hotkeys, dialog.float_paging_settings,
                            dialog.ui.gb_scheduled_hide, dialog.ui.list_hide_times,
                            dialog.ui.btn_add_hide_time, dialog.ui.btn_del_hide_time):
                bounds = control.rect().translated(control.mapTo(page, QPoint()))
                self.assertTrue(page.rect().contains(bounds), (control.objectName(), bounds, page.rect()))
            groups = (dialog.ui.gb_hotkeys, dialog.float_paging_settings, dialog.ui.gb_scheduled_hide)
            for upper, lower in zip(groups, groups[1:]):
                self.assertLess(upper.geometry().bottom(), lower.geometry().top())
        finally:
            self.app.setPalette(original)
            delete(dialog)


if __name__ == "__main__":
    unittest.main()
