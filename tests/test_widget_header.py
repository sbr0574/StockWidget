"""通过实际鼠标事件和渲染验证表头排序、颜色及浮窗拖动。"""

import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, QPoint, QRect, Qt
from PySide6.QtGui import QColor, QMouseEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.ui.widget import FloatLabel


class WidgetHeaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        with patch.object(FloatLabel, "_refresh_from_function"):
            self.window = FloatLabel({
                "header_visible": True,
                "name_visible": False,
                "visible_metrics": ["price", "change_pct"],
                "fg": "#ed952a",
                "bg": {"r": 0, "g": 0, "b": 0, "a": 255},
            }, {})
        self.window.timer.stop()
        self.window._wayland_drag = False
        self.window._last_full_rows = [
            {"名称": name, "现价": str(price), "涨幅": str(change)}
            for name, price, change in (("A", 10, 3), ("B", 30, 1), ("C", 20, 2))
        ]
        self.window._last_color_roles = [
            dict.fromkeys(row, "text") for row in self.window._last_full_rows
        ]
        self.window._last_sort_values = [
            {"现价": 10, "涨幅": 3}, {"现价": 30, "涨幅": 1}, {"现价": 20, "涨幅": 2},
        ]
        self.window._clear_message()
        self.window._reproject_cached_data()
        self.window.show()
        self.app.processEvents()
        self.window.move(100, 100)
        self.header = self.window.table.horizontalHeader()
        self.saved = Mock()
        self.window.set_on_change(self.saved)

    def tearDown(self):
        delete(self.window)
        self.app.processEvents()

    def _header_pos(self, name="现价"):
        section = self.window.model._headers.index(name)
        return QPoint(
            self.header.sectionViewportPosition(section) + self.header.sectionSize(section) // 2,
            self.header.height() // 2,
        )

    def _click(self, name="现价"):
        QTest.mouseClick(self.header.viewport(), Qt.LeftButton, pos=self._header_pos(name))
        self.app.processEvents()

    def _send_mouse(self, target, kind, global_pos):
        button = Qt.NoButton if kind == QEvent.MouseMove else Qt.LeftButton
        buttons = Qt.NoButton if kind == QEvent.MouseButtonRelease else Qt.LeftButton
        event = QMouseEvent(
            kind, target.mapFromGlobal(global_pos), global_pos,
            button, buttons, Qt.NoModifier,
        )
        QApplication.sendEvent(target, event)

    def test_click_cycles_descending_ascending_original_order(self):
        for order, prices in (
            (Qt.DescendingOrder, ["30", "20", "10"]),
            (Qt.AscendingOrder, ["10", "20", "30"]),
            (None, ["10", "30", "20"]),
            (Qt.DescendingOrder, ["30", "20", "10"]),
        ):
            self._click()
            column = self.window.model._headers.index("现价")
            self.assertEqual([row[column] for row in self.window.model._rows], prices)
            self.assertEqual(self.window.sort_header, None if order is None else "现价")
            self.assertEqual(self.header.isSortIndicatorShown(), order is not None)
            self.assertEqual("名称" in self.window.model._headers, order is not None)
            if order is not None:
                self.assertEqual(self.window.sort_order, order)
                self.assertEqual(self.header.sortIndicatorOrder(), order)
                self.assertEqual(self.header.sortIndicatorSection(), column)
        self.assertEqual(self.window.visible_metrics, ["price", "change_pct"])
        self.saved.assert_not_called()

    def test_new_column_starts_descending_and_name_does_not_sort(self):
        self._click()
        self._click()
        self._click("涨幅")
        self.assertEqual(self.window.sort_header, "涨幅")
        self.assertEqual(self.window.sort_order, Qt.DescendingOrder)
        self._click("名称")
        self.assertEqual(self.window.sort_header, "涨幅")
        self.assertEqual(self.window.sort_order, Qt.DescendingOrder)

    def _column_widths(self):
        return {
            name: self.header.sectionSize(index)
            for index, name in enumerate(self.window.model._headers)
        }

    def _sort_arrow_bounds(self, baseline, snapshot):
        self.assertEqual(snapshot.size(), baseline.size())
        changed = [
            (x, y)
            for y in range(snapshot.height())
            for x in range(snapshot.width())
            if snapshot.pixelColor(x, y) != baseline.pixelColor(x, y)
        ]
        self.assertTrue(changed, "排序后应显示箭头")
        rect = QRect(
            QPoint(min(x for x, _ in changed), min(y for _, y in changed)),
            QPoint(max(x for x, _ in changed), max(y for _, y in changed)),
        )
        scale = snapshot.devicePixelRatio()
        self.assertLessEqual(rect.width(), 6 * scale)
        self.assertLessEqual(rect.height(), 5 * scale)
        # 变化只能出现在原有留白中，不能移动、截断或覆盖表头文字。
        for y in range(rect.top(), rect.bottom() + 1):
            for x in range(rect.left(), rect.right() + 1):
                self.assertEqual(baseline.pixelColor(x, y).name(), "#000000")
        return rect

    def test_sort_does_not_expand_columns_or_clip_header_text(self):
        self.window.visible_metrics = ["name", "price", "change_pct"]
        for font_size in (5, 10, 15):
            with self.subTest(font_size=font_size):
                self.window.clear_sort()
                self.window.set_font_size(font_size)
                self.window._reproject_cached_data()
                self.app.processEvents()
                widths = self._column_widths()
                window_width = self.window.width()
                baseline = self.header.viewport().grab().toImage()
                for name in ("现价", "现价", "现价", "涨幅", "涨幅", "涨幅"):
                    self._click(name)
                    self.assertEqual(self._column_widths(), widths)
                    self.assertEqual(self.window.width(), window_width)
                    snapshot = self.header.viewport().grab().toImage()
                    if self.header.isSortIndicatorShown():
                        self._sort_arrow_bounds(baseline, snapshot)
                    else:
                        self.assertEqual(snapshot, baseline)

    def test_sort_arrow_stays_next_to_label_in_wide_columns(self):
        self.window.visible_metrics = ["name", "price", "change_pct"]
        self.window._last_full_rows[0].update({
            "现价": "1234567890.12", "涨幅": "+1234567890.12%",
        })
        for font_size in (5, 10, 15):
            self.window.clear_sort()
            self.window.set_font_size(font_size)
            self.window._reproject_cached_data()
            self.app.processEvents()
            baseline = self.header.viewport().grab().toImage()
            widths = self._column_widths()
            scale = baseline.devicePixelRatio()
            for name in ("现价", "涨幅"):
                for order in (Qt.DescendingOrder, Qt.AscendingOrder):
                    with self.subTest(font_size=font_size, name=name, order=order):
                        self.window.set_sort(name, order)
                        self.app.processEvents()
                        snapshot = self.header.viewport().grab().toImage()
                        arrow = self._sort_arrow_bounds(baseline, snapshot)
                        section = self.header.sortIndicatorSection()
                        left = round(self.header.sectionViewportPosition(section) * scale)
                        right = left + round(self.header.sectionSize(section) * scale)
                        # 从未排序的截图提取实际文字边界，验证箭头与文字间距固定。
                        text_right = max(
                            x for y in range(baseline.height() - round(2 * scale))
                            for x in range(left, right)
                            if baseline.pixelColor(x, y).red() > 0
                        )
                        gap = (arrow.left() - text_right - 1) / scale
                        self.assertGreaterEqual(gap, 1)
                        self.assertLessEqual(gap, 3)
                        self.assertLess(arrow.right(), right)
                        self.assertEqual(self._column_widths(), widths)

    def test_sort_keeps_metric_widths_when_name_is_temporarily_added(self):
        widths = self._column_widths()
        for _ in range(3):
            self._click()
            actual = self._column_widths()
            self.assertEqual({name: actual[name] for name in widths}, widths)

    def test_sorted_columns_still_resize_for_new_quote_content(self):
        self._click()
        widths = self._column_widths()
        self.window._last_full_rows[0]["现价"] = "1234567890.12"
        self.window._reproject_cached_data()
        self.app.processEvents()
        actual = self._column_widths()
        self.assertGreater(actual["现价"], widths["现价"])
        self.assertEqual(actual["名称"], widths["名称"])
        self.assertEqual(actual["涨幅"], widths["涨幅"])

    def test_small_pointer_jitter_still_clicks_without_moving(self):
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        origin = self.window.pos()
        end = start + QPoint(1, 0)
        self._send_mouse(target, QEvent.MouseButtonPress, start)
        self._send_mouse(target, QEvent.MouseMove, end)
        self._send_mouse(target, QEvent.MouseButtonRelease, end)
        self.assertEqual(self.window.pos(), origin)
        self.assertEqual(self.window.sort_header, "现价")
        self.saved.assert_not_called()

    def test_header_and_table_drag_move_window_without_sorting(self):
        self._click()
        for target, pos in (
            (self.header.viewport(), self._header_pos()),
            (self.header.viewport(), self._header_pos("名称")),
            (self.window.table.viewport(), QPoint(5, 5)),
            (self.window, QPoint(1, 1)),
        ):
            with self.subTest(target=target, pos=pos):
                self.saved.reset_mock()
                origin = self.window.pos()
                start = target.mapToGlobal(pos)
                delta = QPoint(QApplication.startDragDistance() + 10, 10)
                self._send_mouse(target, QEvent.MouseButtonPress, start)
                self._send_mouse(target, QEvent.MouseMove, start + delta)
                self._send_mouse(target, QEvent.MouseButtonRelease, start + delta)
                self.assertEqual(self.window.pos(), origin + delta)
                self.assertEqual(self.window.sort_header, "现价")
                self.assertEqual(self.window.sort_order, Qt.DescendingOrder)
                self.saved.assert_called_once_with()

    def test_drag_returning_to_start_does_not_become_click(self):
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        self._send_mouse(target, QEvent.MouseButtonPress, start)
        self._send_mouse(target, QEvent.MouseMove, start + QPoint(50, 0))
        self._send_mouse(target, QEvent.MouseMove, start)
        self._send_mouse(target, QEvent.MouseButtonRelease, start)
        self.assertIsNone(self.window.sort_header)
        self._click()
        self.assertEqual(self.window.sort_order, Qt.DescendingOrder)

    def test_release_outside_pressed_section_does_not_sort(self):
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        self._send_mouse(target, QEvent.MouseButtonPress, start)
        # 即使没有收到中间的 MouseMove，也不能把远处的释放识别为单击。
        self._send_mouse(target, QEvent.MouseButtonRelease, start + QPoint(200, 0))
        self.assertIsNone(self.window.sort_header)

    def test_wayland_only_starts_system_move_after_drag_threshold(self):
        self.window._wayland_drag = True
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        handle = Mock()
        with patch.object(self.window, "windowHandle", return_value=handle):
            self._send_mouse(target, QEvent.MouseButtonPress, start)
            self._send_mouse(target, QEvent.MouseMove, start + QPoint(1, 0))
            handle.startSystemMove.assert_not_called()
            self._send_mouse(target, QEvent.MouseMove, start + QPoint(50, 0))
            self._send_mouse(target, QEvent.MouseMove, start + QPoint(60, 0))
            handle.startSystemMove.assert_called_once_with()
            self._send_mouse(target, QEvent.MouseButtonRelease, start + QPoint(60, 0))
        self.assertIsNone(self.window.sort_header)
        self._click()
        self.assertEqual(self.window.sort_header, "现价")

    def test_header_double_click_still_hides_window_and_clears_gesture(self):
        QTest.mouseDClick(self.header.viewport(), Qt.LeftButton, pos=self._header_pos())
        self.assertTrue(self.window.isHidden())
        self.assertIsNone(self.window._drag_pos)
        self.assertIsNone(self.window.sort_header)
        self.window.show()
        self._click()
        self.assertEqual(self.window.sort_order, Qt.DescendingOrder)

    def test_rendered_sort_arrow_matches_text_color_and_direction(self):
        self.window.visible_metrics = ["name", "price", "change_pct"]
        for color in ("#ed952a", "#ffffff", "#3768bc"):
            self.window.set_fg_color(QColor(color))
            self.window.clear_sort()
            self.window._reproject_cached_data()
            self.app.processEvents()
            baseline = self.header.viewport().grab().toImage()
            for order in (Qt.DescendingOrder, Qt.AscendingOrder):
                with self.subTest(color=color, order=order):
                    self.window.set_sort("现价", order)
                    self.app.processEvents()
                    snapshot = self.header.viewport().grab().toImage()
                    rect = self._sort_arrow_bounds(baseline, snapshot)
                    image = snapshot.copy(rect)
                    rows = [
                        sum(image.pixelColor(x, y).name() == color for x in range(image.width()))
                        for y in range(image.height())
                    ]
                    self.assertGreater(sum(rows), 0, "箭头应使用与文字相同的颜色")
                    top = sum(rows[:len(rows) // 2])
                    bottom = sum(rows[len(rows) // 2:])
                    if order == Qt.DescendingOrder:
                        self.assertGreater(top, bottom)
                    else:
                        self.assertLess(top, bottom)


if __name__ == "__main__":
    unittest.main()
