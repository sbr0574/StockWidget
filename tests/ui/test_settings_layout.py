"""设置按操作顺序展示常用内容，滚动后仍能完整访问次要选项。"""

from PySide6.QtCore import QPoint

from tests.support import SettingsTestCase


class SettingsLayoutTests(SettingsTestCase):
    def test_taskbar_offset_accepts_and_restores_negative_values(self):
        self.enable_windows_taskbar()
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_offset=-80)
        self.assertEqual(dialog.ui.taskbar_offset.value(), -80)
        dialog.ui.taskbar_offset.setValue(-120)
        self.assertEqual(window.current_config()["taskbar_offset"], -120)
        window.set_taskbar_offset(70)
        self.assertEqual(dialog.ui.taskbar_offset.value(), 70)

    def test_common_data_and_window_controls_are_reachable_without_scrolling(self):
        dialog, _window = self._make_dialog()
        dialog.show()
        for page, scroll, control in (
            (dialog.ui.data, dialog.ui.data_scroll, dialog.ui.float_metric_pool),
            (dialog.ui.floating, dialog.ui.floating_scroll, dialog.ui.gb_fcn),
        ):
            with self.subTest(page=page.objectName()):
                dialog.ui.settings_pages.setCurrentWidget(page)
                self.qt_app.processEvents()
                bounds = control.rect().translated(control.mapTo(scroll.viewport(), QPoint()))
                self.assertTrue(scroll.viewport().rect().intersects(bounds))
                self.assertEqual(scroll.verticalScrollBar().value(), 0)
                self.assertEqual(scroll.horizontalScrollBar().maximum(), 0)

    def test_preview_and_detailed_charts_remain_accessible_after_metric_options(self):
        dialog, _window = self._make_dialog()
        dialog.show()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        self.qt_app.processEvents()
        scroll = dialog.ui.data_scroll
        for control in (dialog.ui.gb_mini_chart, dialog.ui.chart_indicators):
            scroll.ensureWidgetVisible(control)
            self.qt_app.processEvents()
            bounds = control.rect().translated(control.mapTo(scroll.viewport(), QPoint()))
            self.assertTrue(scroll.viewport().rect().contains(bounds))
        self.assertEqual(scroll.horizontalScrollBar().maximum(), 0)
