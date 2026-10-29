"""迷你图选择、固定示例及主题同步的用户行为。"""

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtTest import QTest

from tests.support import SettingsTestCase


class MiniChartSettingsTests(SettingsTestCase):
    def test_refresh_settings_work_in_all_themes_and_do_not_save_old_provider(self):
        dialog, window = self._make_dialog(data_source="retired-provider")
        dialog.show()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        self.qt_app.processEvents()
        for value, mode in enumerate(("system", "light", "dark"), 3):
            with self.subTest(mode=mode):
                window.set_view_options(color_mode=mode)
                dialog.ui.data_scroll.ensureWidgetVisible(dialog.ui.sb_interval)
                self.qt_app.processEvents()
                dialog.ui.sb_interval.setValue(value)
                self.assertEqual(window.refresh_seconds, value)
                self.assertEqual(window.current_config()["refresh_seconds"], value)
                self.assertNotIn("data_source", window.current_config())
                self.assertFalse(dialog.ui.sb_interval.grab().isNull())

    def test_data_page_keeps_all_three_preview_rows_inside_the_group(self):
        dialog, _window = self._make_dialog()
        dialog.show()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        self.qt_app.processEvents()
        group = dialog.ui.mini_chart_options
        dialog.ui.data_scroll.ensureWidgetVisible(group)
        self.qt_app.processEvents()
        for kind in ("kline", "trend", "amplitude"):
            preview = getattr(dialog.ui, f"btn_mini_{kind}")
            self.assertGreaterEqual(preview.y(), 25)
            bounds = preview.rect().translated(preview.mapTo(dialog.ui.gb_metric_options, QPoint()))
            self.assertLessEqual(bounds.bottom(), dialog.ui.gb_metric_options.height() - 10)
            self.assertTrue(dialog.ui.data_scroll.viewport().rect().contains(
                preview.rect().translated(preview.mapTo(dialog.ui.data_scroll.viewport(), QPoint()))))
        self.assertEqual(dialog.ui.data_scroll.horizontalScrollBar().maximum(), 0)

    def test_whole_preview_selects_saved_style_and_keeps_all_examples_while_history_is_off(self):
        dialog, window = self._make_dialog(chart_enabled=False)
        dialog.show()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        dialog.ui.data_scroll.ensureWidgetVisible(dialog.ui.mini_chart_options)
        self.qt_app.processEvents()
        for kind in ("trend", "amplitude", "kline"):
            preview = getattr(dialog.ui, f"btn_mini_{kind}")
            QTest.mouseClick(preview, Qt.LeftButton, pos=QPoint(preview.width() - 10, preview.height() - 10))
            with self.subTest(kind=kind):
                self.assertEqual(window.current_config()["mini_chart_type"], kind)
                self.assertEqual(preview.kind, kind)
                self.assertEqual((preview.model.rowCount(), preview.model.columnCount()), (3, 3))
                self.assertEqual([preview.model.index(row, 2).data() for row in range(3)],
                                 ["+4.20%", "-2.60%", "+0.40%"])
                self.assertEqual([preview.model.index(row, 1).data(Qt.UserRole).kind for row in range(3)],
                                 [kind] * 3)
                self.assertFalse(window.view_options.chart_enabled)
                for other in ("kline", "trend", "amplitude"):
                    button = getattr(dialog.ui, f"btn_mini_{other}")
                    self.assertEqual(button.kind, other)
                    self.assertEqual(button.isChecked(), other == kind)
        window.set_view_options(mini_chart_type="amplitude")
        preview = dialog.ui.btn_mini_amplitude
        self.assertTrue(preview.isChecked())
        cells = [preview.model.index(row, 1).data(Qt.UserRole) for row in range(3)]
        self.assertEqual([cell.limits for cell in cells], [(-5.0, 7.0)] * 3)
        dialog.ui.btn_mini_trend.setFocus()
        QTest.keyClick(dialog.ui.btn_mini_trend, Qt.Key_Space)
        self.assertEqual(window.view_options.mini_chart_type, "trend")

    def test_preview_follows_direction_colors_and_keeps_theme_text_readable(self):
        dialog, window = self._make_dialog(mini_chart_type="trend")
        preview = dialog.ui.btn_mini_trend
        window.set_unicolor(False)
        window.set_up_color(QColor("#ad3412"))
        window.set_down_color(QColor("#24864a"))
        self.assertEqual(preview.model.index(0, 2).data(Qt.ForegroundRole), QColor("#ad3412"))
        self.assertEqual(preview.model.index(1, 2).data(Qt.ForegroundRole), QColor("#24864a"))
        window.set_unicolor(True)
        for mode in ("light", "dark"):
            window.set_view_options(color_mode=mode)
            foreground = preview.model.index(0, 2).data(Qt.ForegroundRole)
            self.assertEqual(foreground, preview.palette().color(QPalette.Text))
            self.assertGreater(abs(foreground.lightness() - preview.palette().color(QPalette.Window).lightness()), 100)
            image = preview.grab().toImage()
            self.assertFalse(image.isNull())
            # Unselected cards retain a visible background without hovering.
            idle = dialog.ui.btn_mini_amplitude
            image = idle.grab().toImage()
            ratio = image.devicePixelRatio()
            self.assertNotEqual(image.pixelColor(round(8 * ratio), round(20 * ratio)),
                                idle.palette().color(QPalette.Window))
