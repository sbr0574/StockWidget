"""Render real tables to check single-stroke grids and the shared bid/ask axis."""

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QColor, QImage, QPainter
from PySide6.QtWidgets import QApplication, QStyleOptionViewItem
from shiboken6 import delete

from stockwidget.core.quote_presentation import BidAskCell
from stockwidget.ui.table_model import bid_ask_width
from stockwidget.ui.taskbar import render_taskbar
from stockwidget.ui.widget import FloatLabel


class QuoteTableTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.enterContext(patch.object(FloatLabel, "_refresh_from_function"))
        self.enterContext(patch("stockwidget.ui.widget.GlobalHotkeyManager"))
        self.window = FloatLabel({
            "header_visible": True, "grid_visible": True, "fg": "#ffffff",
            "bg": {"r": 0, "g": 0, "b": 0, "a": 255}, "opacity_pct": 100,
            "visible_metrics": ["name", "price", "kline"],
        }, {})
        self.window._clear_message()
        self.window._last_full_rows = [{"名称": "", "现价": "", "K线": {"k": (10, 11, 12, 9, 10)}} for _ in range(3)]
        self.window._last_color_roles = [{} for _ in range(3)]
        self.window._reproject_cached_data()
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
            w.change_page("float", 1)
            self.app.processEvents()
            self.assertEqual(w.table.grid_layer.geometry(), w.table.rect())
            w.set_view_options(float_paging_enabled=False)

    def test_one_bid_ask_column_keeps_colors_markers_and_a_fixed_axis_in_both_surfaces(self):
        w = self.window
        cells = [BidAskCell("1<", "234567", "up", "down"),
                 BidAskCell("123456", ">2", "up", "down"), BidAskCell("-", "-")]
        w._last_full_rows = [{"买一/卖一": cell} for cell in cells]
        w._last_color_roles = [{"买一/卖一": "text"} for _ in cells]
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
            with patch("stockwidget.ui.taskbar.QPainter", Recorder):
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
        w._last_full_rows = [{"买一/卖一": BidAskCell("23<", ">45", "up", "down")}]
        w._last_color_roles = [{"买一/卖一": "text"}]
        w.set_visible_metrics(["b1s1"])
        w.set_unicolor(True)
        w.set_view_options(taskbar_sync_appearance=False, taskbar_unicolor=False, taskbar_color="#112233")
        calls = []

        class Recorder(QPainter):
            def drawText(self, *args):
                calls.append((args[-1], self.pen().color()))
                return super().drawText(*args)

        with patch("stockwidget.ui.taskbar.QPainter", Recorder):
            render_taskbar(w, 60)
        self.assertEqual(calls, [("23<", w.up_color), (">45", w.down_color), ("/", QColor("#112233"))])
        self.assertEqual(w.model.color_for_role("down"), w.fg)
