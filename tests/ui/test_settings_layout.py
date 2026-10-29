"""设置按操作顺序展示常用内容，滚动后仍能完整访问次要选项。"""

from PySide6.QtCore import QPoint

from tests.support import SettingsTestCase


class SettingsLayoutTests(SettingsTestCase):
    def test_data_general_and_appearance_keep_controls_in_their_responsible_group(self):
        dialog, window = self._make_dialog(chart_enabled=True, name_length=2)
        dialog.show()
        for group, controls in (
            (dialog.ui.gb_data_setting, (dialog.ui.sb_interval, dialog.ui.cb_chart_enabled)),
            (dialog.ui.gb_data, (dialog.ui.float_metric_pool, dialog.ui.cb_head)),
            (dialog.ui.gb_metric_options, (dialog.ui.cmb_name_length, dialog.ui.cb_code_visible,
                                           dialog.ui.cb_type_visible, dialog.ui.cmb_unit_mode,
                                           dialog.ui.cmb_average_period, dialog.ui.mini_chart_options)),
            (dialog.ui.general, (dialog.ui.gb_fcn, dialog.ui.float_row_settings, dialog.ui.float_split_settings)),
            (dialog.ui.floating, (dialog.ui.gb_text, dialog.ui.gb_color, dialog.ui.gb_opacity,
                                  dialog.ui.gb_icon, dialog.ui.cmb_color_mode)),
        ):
            for control in controls:
                self.assertTrue(group.isAncestorOf(control))
        for page, scroll in ((dialog.ui.data, dialog.ui.data_scroll), (dialog.ui.general, dialog.ui.general_scroll),
                             (dialog.ui.floating, dialog.ui.floating_scroll)):
            dialog.ui.settings_pages.setCurrentWidget(page)
            self.qt_app.processEvents()
            self.assertEqual(scroll.horizontalScrollBar().maximum(), 0)
        dialog.ui.cmb_name_length.setCurrentIndex(dialog.ui.cmb_name_length.findData(0))
        self.assertEqual(window.current_config()["name_length"], 0)
        window.set_name_length(-1)
        self.assertEqual(dialog.ui.cmb_name_length.currentData(), -1)
        dialog.ui.cb_chart_enabled.setChecked(False)
        self.assertFalse(window.view_options.chart_enabled)

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
            (dialog.ui.general, dialog.ui.general_scroll, dialog.ui.gb_fcn),
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
        for control in (dialog.ui.average_period_row, dialog.ui.mini_chart_options, dialog.ui.chart_enabled_row):
            scroll.ensureWidgetVisible(control)
            self.qt_app.processEvents()
            bounds = control.rect().translated(control.mapTo(scroll.viewport(), QPoint()))
            self.assertTrue(scroll.viewport().rect().contains(bounds))
        self.assertEqual(scroll.horizontalScrollBar().maximum(), 0)
