"""共用行情绘制：边框、颜色、表头排序与鼠标交互。"""

from unittest.mock import patch, Mock

from PySide6.QtCore import QPoint, QRect, Qt, QEvent
from PySide6.QtGui import QColor, QImage, QPainter, QMouseEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QStyleOptionViewItem
from shiboken6 import delete

from stockwidget.core.quote_presentation import BidAskCell, direction_color_role
from stockwidget.ui.controls.quote_view import (
    bid_ask_width,
    COLOR_ROLE_DOWN,
    COLOR_ROLE_NEUTRAL,
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
    KLineDelegate,
    SimpleTableModel,
)
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import render_taskbar
from stockwidget.ui.floating.widget import FloatLabel

from tests.support import QtTestCase


class QuoteTableTests(QtTestCase):

    def setUp(self):
        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.window = FloatLabel({
            "header_visible": True, "grid_visible": True, "fg": "#ffffff",
            "bg": {"r": 0, "g": 0, "b": 0, "a": 255}, "opacity_pct": 100,
            "visible_metrics": ["name", "price", "kline"],
        }, {})
        self.window._clear_message()
        self.window.quotes._last_full_rows = [{"名称": "", "现价": "", "K线": {"k": (10, 11, 12, 9, 10)}} for _ in range(3)]
        self.window.quotes._last_color_roles = [{} for _ in range(3)]
        self.window.quotes.reproject()
        self.window.show()
        self.app.processEvents()

    def tearDown(self):
        delete(self.window)
        self.app.processEvents()

    def test_grid_draws_header_and_candle_edges_once_without_double_outer_edges(self):
        table = self.window.table
        image = table.grab().toImage()
        ratio = image.devicePixelRatio()

        def brightness(x, y):
            return image.pixelColor(round(x * ratio), round(y * ratio)).red()

        origin = table.viewport().mapTo(table, QPoint())
        x = table.columnWidth(0) - 1
        y = origin.y() + table.rowHeight(0) - 1
        # Equal opacity at a crossing proves the translucent strokes do not stack.
        crossing = brightness(x, y)
        vertical = brightness(x, origin.y() + table.rowHeight(0) // 2)
        horizontal = brightness(table.columnWidth(0) // 2, y)
        self.assertLessEqual(abs(crossing - vertical), 2)
        self.assertLessEqual(abs(crossing - horizontal), 2)
        self.assertGreater(vertical, 0)
        self.assertGreater(brightness(x, origin.y() // 2), 0)
        self.assertGreater(brightness(table.columnWidth(0) // 2, origin.y() - 1), 0)
        candle_x = table.columnViewportPosition(2) + 3
        self.assertGreater(brightness(candle_x, y), 0)
        self.assertEqual(brightness(table.columnWidth(0) // 2, table.height() - 2), 0)
        self.assertGreater(brightness(table.columnWidth(0) // 2, table.height() - 1), 0)
        self.assertEqual(brightness(table.width() - 2, origin.y() + 3), 0)
        self.assertGreater(brightness(table.width() - 1, origin.y() + 3), 0)
        self.assertLessEqual(max(brightness(px, py) for px in range(3) for py in range(3)), vertical + 2)

    def test_grid_updates_after_color_header_and_page_changes_and_does_not_take_mouse_hits(self):
        w = self.window
        for split in (False, True):
            w.set_view_options(float_split_enabled=split)
            for table in (w.float_tables if split else (w.table,)):
                self.assertTrue(table.grid_layer.testAttribute(Qt.WA_TransparentForMouseEvents))
                point = table.viewport().mapTo(table, QPoint(3, 3))
                self.assertIs(table.childAt(point), table.viewport())
                self.assertEqual(table.grid_layer.geometry(), table.rect())
            w.set_grid_visible(False)
            self.assertTrue(all(table.grid_layer.isHidden() for table in w.float_tables))
            w.set_fg_color(QColor("#123456"))
            w.set_header_visible(False)
            w.set_grid_visible(True)
            self.app.processEvents()
            expected = QColor("#123456")
            expected.setAlpha(80)
            self.assertEqual(w.table.grid_layer.color, expected)
            self.assertEqual(w.table.grid_layer.geometry(), w.table.rect())
            w.set_view_options(float_paging_enabled=True, float_max_rows=1)
            w.quotes.change_page("float", 1)
            self.app.processEvents()
            self.assertEqual(w.table.grid_layer.geometry(), w.table.rect())
            w.set_view_options(float_paging_enabled=False)

    def test_one_bid_ask_column_keeps_colors_markers_and_a_fixed_axis_in_both_surfaces(self):
        w = self.window
        cells = [BidAskCell("1<", "234567", "up", "down"),
                 BidAskCell("123456", ">2", "up", "down"), BidAskCell("-", "-")]
        w.quotes._last_full_rows = [{"买一/卖一": cell} for cell in cells]
        w.quotes._last_color_roles = [{"买一/卖一": "text"} for _ in cells]
        w.set_unicolor(False)
        w.set_visible_metrics(["b1s1"])
        w.set_view_options(taskbar_sync_metrics=True, taskbar_rows=3)
        self.assertEqual(w.model.columnCount(), 1)
        self.assertEqual(w.taskbar_model.columnCount(), 1)
        self.assertEqual(w.current_config()["visible_metrics"], ["b1s1"])
        self.assertTrue(w.current_config()["b1s1_visible"])
        self.assertEqual(w.model.index(0, 0).data(Qt.UserRole), cells[0])
        self.assertEqual(w.model.index(0, 0).data(), "1< / 234567")
        calls = []

        class Recorder(QPainter):
            def drawText(self, *args):
                calls.append((args[0], args[-1], self.pen().color()))
                return super().drawText(*args)

        width = max(bid_ask_width(cell, w.font) for cell in cells)
        image = QImage(width, 100, QImage.Format_ARGB32_Premultiplied)
        image.fill(Qt.transparent)
        painter = Recorder(image)
        for row, cell in enumerate(cells):
            option = QStyleOptionViewItem()
            option.font, option.rect = w.font, QRect(0, row * 30, width, 30)
            w._default_item_delegate.paint(painter, option, w.model.index(row, 0))
            buy, sell, separator = calls[-3:]
            self.assertEqual((buy[1], sell[1], separator[1]), (cell.buy, cell.sell, "/"))
            self.assertAlmostEqual((buy[0].right() + sell[0].left()) / 2, width / 2)
            self.assertAlmostEqual(separator[0].center().x(), width / 2)
            self.assertEqual(buy[2], w.model.color_for_role(cell.buy_role))
            self.assertEqual(sell[2], w.model.color_for_role(cell.sell_role))
        painter.end()
        for dpi in (96, 192):
            calls.clear()
            with patch("stockwidget.ui.floating.taskbar.QPainter", Recorder):
                result = render_taskbar(w, 90 * dpi // 96, dpi, max_width=1000)
            self.assertFalse(result.isNull())
            separators = calls[2::3]
            self.assertEqual(len(separators), 3)
            self.assertEqual(len({rect.center().x() for rect, _, _ in separators}), 1)
        w.set_unicolor(True)
        self.assertEqual(w.model.color_for_role("up"), w.fg)
        self.assertEqual(w.taskbar_model.color_for_role("down"), w.fg)
        # Reordering the pair into the old candle column must restore this delegate.
        w.set_visible_metrics(["b1s1", "kline"])
        w.set_visible_metrics(["kline", "b1s1"])
        self.assertIs(w.table.itemDelegateForColumn(1), w._default_item_delegate)
        self.assertEqual(w.model.index(0, 1).data(Qt.UserRole), cells[0])

    def test_taskbar_keeps_its_independent_colors_for_both_sides(self):
        w = self.window
        w.quotes._last_full_rows = [{"买一/卖一": BidAskCell("23<", ">45", "up", "down")}]
        w.quotes._last_color_roles = [{"买一/卖一": "text"}]
        w.set_visible_metrics(["b1s1"])
        w.set_unicolor(True)
        w.set_view_options(taskbar_sync_appearance=False, taskbar_unicolor=False, taskbar_color="#112233")
        calls = []

        class Recorder(QPainter):
            def drawText(self, *args):
                calls.append((args[-1], self.pen().color()))
                return super().drawText(*args)

        with patch("stockwidget.ui.floating.taskbar.QPainter", Recorder):
            render_taskbar(w, 60)
        self.assertEqual(calls, [("23<", w.up_color), (">45", w.down_color), ("/", QColor("#112233"))])
        self.assertEqual(w.model.color_for_role("down"), w.fg)


class ColorSystemTests(QtTestCase):

    def test_direction_role_mapping(self):
        self.assertEqual(direction_color_role(1), COLOR_ROLE_UP)
        self.assertEqual(direction_color_role(-1), COLOR_ROLE_DOWN)
        self.assertEqual(direction_color_role(0), COLOR_ROLE_NEUTRAL)

    def test_table_uses_separate_roles_then_unifies_to_text_color(self):
        model = SimpleTableModel()
        model.set_rows_headers(
            [["普通", "+1", "-1", "0"]],
            ["普通", "上涨", "下跌", "中性"],
            [[COLOR_ROLE_TEXT, COLOR_ROLE_UP, COLOR_ROLE_DOWN, COLOR_ROLE_NEUTRAL]],
        )
        colors = (
            QColor("#112233"),
            QColor("#aa0000"),
            QColor("#00aa00"),
            QColor("#777777"),
        )
        model.set_colors(False, *colors)

        actual = [
            model.data(model.index(0, column), Qt.ItemDataRole.ForegroundRole).name()
            for column in range(4)
        ]
        self.assertEqual(actual, [color.name() for color in colors])
        self.assertEqual(
            model.headerData(0, Qt.Orientation.Horizontal, Qt.ItemDataRole.ForegroundRole).name(),
            colors[0].name(),
        )

        model.set_colors(True, *colors)
        unified = [
            model.data(model.index(0, column), Qt.ItemDataRole.ForegroundRole).name()
            for column in range(4)
        ]
        self.assertEqual(unified, [colors[0].name()] * 4)

    def test_kline_uses_direction_colors_and_unifies_when_enabled(self):
        delegate = KLineDelegate()
        colors = (
            QColor("#112233"),
            QColor("#aa0000"),
            QColor("#00aa00"),
            QColor("#777777"),
        )
        delegate.set_colors(False, *colors)

        self.assertEqual(delegate.candle_color(1, 2).name(), colors[1].name())
        self.assertEqual(delegate.candle_color(2, 1).name(), colors[2].name())
        self.assertEqual(delegate.candle_color(1, 1).name(), colors[3].name())
        self.assertEqual(delegate.reference_color().name(), colors[3].name())

        delegate.set_colors(True, *colors)
        self.assertEqual(delegate.candle_color(1, 2).name(), colors[0].name())
        self.assertEqual(delegate.reference_color().name(), colors[0].name())


class WidgetHeaderTests(QtTestCase):

    def setUp(self):
        # 这些测试使用手动填充的行情，显示窗口时也不能刷新并覆盖样本。
        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.window = FloatLabel({
            "header_visible": True,
            "name_visible": False,
            "visible_metrics": ["price", "change_pct"],
            "fg": "#ed952a",
            "bg": {"r": 0, "g": 0, "b": 0, "a": 255},
        }, {})
        self.window.timer.stop()
        self.window._wayland_drag = False
        self.window.quotes._last_full_rows = [
            {"名称": name, "现价": str(price), "涨幅": str(change)}
            for name, price, change in (("A", 10, 3), ("B", 30, 1), ("C", 20, 2))
        ]
        self.window.quotes._last_color_roles = [
            dict.fromkeys(row, "text") for row in self.window.quotes._last_full_rows
        ]
        self.window.quotes._last_sort_values = [
            {"现价": 10, "涨幅": 3}, {"现价": 30, "涨幅": 1}, {"现价": 20, "涨幅": 2},
        ]
        self.window._clear_message()
        self.window.quotes.reproject()
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
            self.assertEqual(self.window.quotes.sort_header, None if order is None else "现价")
            self.assertEqual(self.header.isSortIndicatorShown(), order is not None)
            self.assertEqual("名称" in self.window.model._headers, order is not None)
            if order is not None:
                self.assertEqual(self.window.quotes.sort_order, order)
                self.assertEqual(self.header.sortIndicatorOrder(), order)
                self.assertEqual(self.header.sortIndicatorSection(), column)
        self.assertEqual(self.window.visible_metrics, ["price", "change_pct"])
        self.saved.assert_not_called()

    def test_new_column_starts_descending_and_name_does_not_sort(self):
        self._click()
        self._click()
        self._click("涨幅")
        self.assertEqual(self.window.quotes.sort_header, "涨幅")
        self.assertEqual(self.window.quotes.sort_order, Qt.DescendingOrder)
        self._click("名称")
        self.assertEqual(self.window.quotes.sort_header, "涨幅")
        self.assertEqual(self.window.quotes.sort_order, Qt.DescendingOrder)

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
                self.window.quotes.clear_sort()
                self.window.set_font_size(font_size)
                self.window.quotes.reproject()
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
        self.window.quotes._last_full_rows[0].update({
            "现价": "1234567890.12", "涨幅": "+1234567890.12%",
        })
        for font_size in (5, 10, 15):
            self.window.quotes.clear_sort()
            self.window.set_font_size(font_size)
            self.window.quotes.reproject()
            self.app.processEvents()
            baseline = self.header.viewport().grab().toImage()
            widths = self._column_widths()
            scale = baseline.devicePixelRatio()
            for name in ("现价", "涨幅"):
                for order in (Qt.DescendingOrder, Qt.AscendingOrder):
                    with self.subTest(font_size=font_size, name=name, order=order):
                        self.window.quotes.set_sort(name, order)
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
        self.window.quotes._last_full_rows[0]["现价"] = "1234567890.12"
        self.window.quotes.reproject()
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
        self.assertEqual(self.window.quotes.sort_header, "现价")
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
                self.assertEqual(self.window.quotes.sort_header, "现价")
                self.assertEqual(self.window.quotes.sort_order, Qt.DescendingOrder)
                self.saved.assert_called_once_with()

    def test_drag_returning_to_start_does_not_become_click(self):
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        self._send_mouse(target, QEvent.MouseButtonPress, start)
        self._send_mouse(target, QEvent.MouseMove, start + QPoint(50, 0))
        self._send_mouse(target, QEvent.MouseMove, start)
        self._send_mouse(target, QEvent.MouseButtonRelease, start)
        self.assertIsNone(self.window.quotes.sort_header)
        self._click()
        self.assertEqual(self.window.quotes.sort_order, Qt.DescendingOrder)

    def test_release_outside_pressed_section_does_not_sort(self):
        target = self.header.viewport()
        start = target.mapToGlobal(self._header_pos())
        self._send_mouse(target, QEvent.MouseButtonPress, start)
        # 即使没有收到中间的 MouseMove，也不能把远处的释放识别为单击。
        self._send_mouse(target, QEvent.MouseButtonRelease, start + QPoint(200, 0))
        self.assertIsNone(self.window.quotes.sort_header)

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
        self.assertIsNone(self.window.quotes.sort_header)
        self._click()
        self.assertEqual(self.window.quotes.sort_header, "现价")

    def test_header_double_click_still_hides_window_and_clears_gesture(self):
        QTest.mouseDClick(self.header.viewport(), Qt.LeftButton, pos=self._header_pos())
        self.assertTrue(self.window.isHidden())
        self.assertIsNone(self.window._drag_pos)
        self.assertIsNone(self.window.quotes.sort_header)
        self.window.show()
        self._click()
        self.assertEqual(self.window.quotes.sort_order, Qt.DescendingOrder)

    def test_rendered_sort_arrow_matches_text_color_and_direction(self):
        self.window.visible_metrics = ["name", "price", "change_pct"]
        for color in ("#ed952a", "#ffffff", "#3768bc"):
            self.window.set_fg_color(QColor(color))
            self.window.quotes.clear_sort()
            self.window.quotes.reproject()
            self.app.processEvents()
            baseline = self.header.viewport().grab().toImage()
            for order in (Qt.DescendingOrder, Qt.AscendingOrder):
                with self.subTest(color=color, order=order):
                    self.window.quotes.set_sort("现价", order)
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
