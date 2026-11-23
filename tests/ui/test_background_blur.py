"""跨平台背景模糊开关、整体不透明度锁定、鼠标穿透双向互斥与同步行为测试。"""

import sys
from unittest.mock import Mock, patch

from shiboken6 import delete

from stockwidget.core.view_options import ViewOptions
from stockwidget.data.bars import BarResult
from stockwidget.data.quotes import _new_entry
from stockwidget.platform import capabilities
from stockwidget.ui.menus import build_history_menu
from tests.data.test_bars import sample_bars
from tests.support import SettingsTestCase


def _menu_actions(menu):
    return {action.text(): action for action in menu.actions() if not action.isSeparator()}


class BackgroundBlurTests(SettingsTestCase):
    def test_blur_sync_locks_overall_opacity_and_preserves_saved_values(self):
        watchlist = {
            "sh600000": {"code": "600000", "market": "sh", "type": "沪", "name": "浦发银行", "checked": True},
        }
        quote = {
            "sh600000": _new_entry("浦发银行", 10.0, 9.9, 10.2, 10.3, 9.8,
                                   100000, 1000000, [1000] * 5, [10.1] * 5, [1000] * 5, [10.2] * 5,
                                   "2026-11-23", "14:45:04"),
        }
        with patch("stockwidget.ui.floating.widget.background_blur_supported", return_value=True), \
                patch("stockwidget.ui.settings.dialog.background_blur_supported", return_value=True), \
                patch("stockwidget.ui.floating.widget.apply_background_blur") as float_blur_call, \
                patch("stockwidget.ui.history.apply_background_blur") as chart_blur_call:
            dialog, window = self._make_dialog(watchlist=watchlist, opacity_pct=70, chart_sync_opacity=True, chart_enabled=True)
            window.history.cache.get = Mock(return_value=BarResult(sample_bars(), "sina"))
            window.quotes.accept_result((True, quote, None))
            window.show()
            window.history.open_row("float", 0)
            while window.history._busy:
                self.qt_app.processEvents()
            self.qt_app.processEvents()
            popup = window.history.dialog

            self.assertFalse(window.view_options.float_bg_blur)
            self.assertFalse(window.view_options.chart_bg_blur)
            self.assertAlmostEqual(window.windowOpacity(), 0.70, delta=1 / 255)
            self.assertAlmostEqual(popup.windowOpacity(), 0.70, delta=1 / 255)
            self.assertTrue(dialog.ui.slider_all_alpha.isEnabled())
            self.assertFalse(dialog.ui.cb_chart_blur.isEnabled())

            # Enabling main float blur locks both surfaces' overall opacity to 100% while synced.
            dialog.ui.cb_float_blur.click()
            self.qt_app.processEvents()
            self.assertTrue(window.view_options.float_bg_blur)
            self.assertTrue(window.view_options.chart_bg_blur)
            self.assertTrue(dialog.ui.cb_chart_blur.isChecked())
            self.assertFalse(dialog.ui.cb_chart_blur.isEnabled())
            self.assertEqual(window.opacity_pct, 70)
            self.assertEqual(window.view_options.chart_opacity_pct, 70)
            self.assertEqual(window.windowOpacity(), 1.0)
            self.assertEqual(popup.windowOpacity(), 1.0)
            self.assertFalse(dialog.ui.slider_all_alpha.isEnabled())
            self.assertFalse(dialog.ui.slider_chart_all_alpha.isEnabled())
            self.assertTrue(float_blur_call.call_args.args[1])
            self.assertTrue(chart_blur_call.call_args.args[1])

            # Context menus disable overall opacity presets while blur is active; taskbar stays enabled.
            float_menu = window.build_context_menu("float")
            taskbar_menu = window.build_context_menu("taskbar")
            chart_menu = build_history_menu(popup)
            try:
                self.assertFalse(_menu_actions(float_menu)["不透明度"].isEnabled())
                self.assertTrue(_menu_actions(taskbar_menu)["不透明度"].isEnabled())
                self.assertFalse(_menu_actions(chart_menu)["不透明度"].isEnabled())
            finally:
                for menu in (float_menu, taskbar_menu, chart_menu):
                    delete(menu)

            # Turning off sync keeps chart's synced blur state and lets both toggles work independently.
            dialog.ui.cb_chart_sync_opacity.click()
            self.qt_app.processEvents()
            self.assertTrue(window.view_options.chart_bg_blur)
            self.assertTrue(dialog.ui.cb_chart_blur.isEnabled())
            self.assertFalse(dialog.ui.slider_chart_all_alpha.isEnabled())

            dialog.ui.cb_float_blur.click()
            self.qt_app.processEvents()
            self.assertFalse(window.view_options.float_bg_blur)
            self.assertAlmostEqual(window.windowOpacity(), 0.70, delta=1 / 255)
            self.assertTrue(dialog.ui.slider_all_alpha.isEnabled())
            self.assertTrue(window.view_options.chart_bg_blur)
            self.assertEqual(popup.windowOpacity(), 1.0)
            self.assertFalse(dialog.ui.slider_chart_all_alpha.isEnabled())

            dialog.ui.cb_chart_blur.click()
            self.qt_app.processEvents()
            self.assertFalse(window.view_options.chart_bg_blur)
            self.assertAlmostEqual(popup.windowOpacity(), 0.70, delta=1 / 255)
            self.assertTrue(dialog.ui.slider_chart_all_alpha.isEnabled())
            chart_menu_after = build_history_menu(popup)
            try:
                self.assertTrue(_menu_actions(chart_menu_after)["不透明度"].isEnabled())
            finally:
                delete(chart_menu_after)

            # Detached window never requests blur even when popup blur is enabled.
            dialog.ui.cb_chart_blur.click()
            self.qt_app.processEvents()
            chart_blur_call.reset_mock()
            popup.pop_out_button.click()
            self.qt_app.processEvents()
            chart_blur_call.assert_not_called()

    def test_float_blur_and_click_through_are_mutually_exclusive_bidirectionally(self):
        with patch("stockwidget.ui.floating.widget.background_blur_supported", return_value=True), \
                patch("stockwidget.ui.settings.dialog.background_blur_supported", return_value=True), \
                patch("stockwidget.ui.floating.widget.click_through_supported", return_value=True), \
                patch("stockwidget.ui.settings.dialog.click_through_supported", return_value=True):
            dialog, window = self._make_dialog(opacity_pct=60, click_through=True, chart_sync_opacity=True)
            self.assertTrue(window.click_through)
            self.assertFalse(window.view_options.float_bg_blur)

            # Enabling main float blur disables click-through.
            dialog.ui.cb_float_blur.click()
            self.qt_app.processEvents()
            self.assertTrue(window.view_options.float_bg_blur)
            self.assertFalse(window.click_through)
            self.assertFalse(dialog.ui.cb_click_through.isChecked())
            self.assertTrue(dialog.ui.cb_click_through.isEnabled())
            self.assertEqual(window.windowOpacity(), 1.0)

            # Re-enabling click-through (e.g. via settings, menu, or hotkey) disables main float blur.
            window.toggle_click_through()
            self.qt_app.processEvents()
            self.assertTrue(window.click_through)
            self.assertTrue(dialog.ui.cb_click_through.isChecked())
            self.assertFalse(window.view_options.float_bg_blur)
            self.assertFalse(dialog.ui.cb_float_blur.isChecked())
            self.assertFalse(window.view_options.chart_bg_blur)
            self.assertAlmostEqual(window.windowOpacity(), 0.60, delta=1 / 255)
            self.assertTrue(dialog.ui.slider_all_alpha.isEnabled())

            # When chart style is unsynced, enabling click-through does not disable independent chart blur.
            dialog.ui.cb_chart_sync_opacity.click()
            dialog.ui.cb_chart_blur.click()
            dialog.ui.cb_float_blur.click()
            self.qt_app.processEvents()
            self.assertTrue(window.view_options.float_bg_blur)
            self.assertTrue(window.view_options.chart_bg_blur)
            dialog.ui.cb_click_through.click()
            self.qt_app.processEvents()
            self.assertTrue(window.click_through)
            self.assertFalse(window.view_options.float_bg_blur)
            self.assertTrue(window.view_options.chart_bg_blur)

            # Conflicting saved config resolves in favor of blur (disables click-through).
            _, conflict_window = self._make_dialog(float_bg_blur=True, click_through=True)
            self.assertTrue(conflict_window.view_options.float_bg_blur)
            self.assertFalse(conflict_window.click_through)

            # Reset appearance turns off both blur toggles and restores default opacity.
            dialog.ui.btn_reset_appearance.click()
            self.qt_app.processEvents()
            self.assertFalse(window.view_options.float_bg_blur)
            self.assertFalse(window.view_options.chart_bg_blur)
            self.assertTrue(window.view_options.chart_sync_opacity)

    def test_platform_blur_capabilities_and_disabled_tooltips(self):
        with patch("sys.platform", "win32"):
            self.assertTrue(capabilities.background_blur_supported())
        with patch("sys.platform", "darwin"), patch("stockwidget.platform.capabilities.is_cocoa", return_value=True):
            self.assertTrue(capabilities.background_blur_supported())
        with patch("sys.platform", "darwin"), patch("stockwidget.platform.capabilities.is_cocoa", return_value=False):
            self.assertFalse(capabilities.background_blur_supported())
        with patch("sys.platform", "linux"), patch("stockwidget.platform.capabilities.session_type", return_value="wayland"):
            self.assertFalse(capabilities.background_blur_supported())
            self.assertIn("Wayland", capabilities.background_blur_unsupported_tooltip())
            dialog, _ = self._make_dialog()
            self.assertFalse(dialog.ui.cb_float_blur.isEnabled())
            self.assertFalse(dialog.ui.cb_chart_blur.isEnabled())
            self.assertIn("KDE", dialog.ui.cb_float_blur.toolTip())
        with patch("sys.platform", "linux"), patch("stockwidget.platform.capabilities.session_type", return_value="x11"), \
                patch("stockwidget.platform.capabilities._x11_kwin_blur_supported", return_value=False):
            self.assertFalse(capabilities.background_blur_supported())
            self.assertIn("KDE", capabilities.background_blur_unsupported_tooltip())
            dialog, _ = self._make_dialog()
            self.assertFalse(dialog.ui.cb_float_blur.isEnabled())
            self.assertFalse(dialog.ui.cb_chart_blur.isEnabled())
        with patch("sys.platform", "linux"), patch("stockwidget.platform.capabilities.session_type", return_value="x11"), \
                patch("stockwidget.platform.capabilities._x11_kwin_blur_supported", return_value=True):
            self.assertTrue(capabilities.background_blur_supported())

    def test_view_options_blur_defaults_and_windows_border_radius(self):
        defaults = ViewOptions.from_config({})
        self.assertFalse(defaults.float_bg_blur)
        self.assertFalse(defaults.chart_bg_blur)
        loaded = ViewOptions.from_config({"float_bg_blur": 1, "chart_bg_blur": True})
        self.assertTrue(loaded.float_bg_blur)
        self.assertTrue(loaded.chart_bg_blur)
        if sys.platform == "win32":
            _, window = self._make_dialog()
            self.assertIn("border-radius: 5px;", window.panel.styleSheet())
            window.set_view_options(float_bg_blur=True)
            self.assertIn("border-radius: 0px;", window.panel.styleSheet())
            window.set_view_options(float_bg_blur=False)
            self.assertIn("border-radius: 5px;", window.panel.styleSheet())
