"""两处独立分页、分栏、自动翻页和全区域拖动。"""

from unittest.mock import patch
import os

from PySide6.QtCore import QEvent, QPoint, Qt
from PySide6.QtGui import QColor, QKeyEvent, QPainter
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from stockwidget.core.view_options import ViewOptions
from stockwidget.ui.floating.taskbar import render_taskbar

from tests.support import PagingTestCase


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


class PagingInteractionTests(PagingTestCase):
    def test_independent_pages_and_incomplete_last_page(self):
        self.win.set_view_options(float_max_rows=2, float_page_mode="manual",
                                  taskbar_rows=3, taskbar_page_mode="manual")
        self.win.quotes.change_page("taskbar", 1)
        self.assertEqual(self.win.taskbar_model.rowCount(), 2)
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(), "标的4")
        self.assertEqual(self.win.model.index(0, 0).data(), "标的1")
        self.win.quotes.change_page("float", 1)
        self.assertEqual(self.win.model.index(0, 0).data(), "标的3")
        self.assertEqual(self.win.quotes.taskbar_page, 1)
        self.populate()
        self.assertEqual((self.win.quotes.float_page, self.win.quotes.taskbar_page), (1, 1))
        self.populate(2)
        self.assertEqual((self.win.quotes.float_page, self.win.quotes.taskbar_page), (0, 0))

    def test_split_pages_limit_each_column_and_balance_the_last_page(self):
        w = self.win
        self.populate(7)
        w.set_view_options(float_split_enabled=True, float_max_rows=2,
                           taskbar_sync_split=True, taskbar_rows=3, taskbar_page_mode="manual")
        self.assertEqual(w.quotes.get_page("float").count, 2)
        self.assertEqual(w.model._rows[0][0], "标的1")
        self.assertEqual(w.right_model._rows[0][0], "标的3")
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (2, 2))
        self.assertEqual(w.taskbar_model.rowCount(), 6)
        w.quotes.change_page("float", 1)
        self.assertEqual([row[0] for row in w.model._rows], ["标的5", "标的6"])
        self.assertEqual([row[0] for row in w.right_model._rows], ["标的7"])
        w.quotes.change_page("taskbar", 1)
        self.assertEqual(w.taskbar_model._rows[0][0], "标的7")
        self.assertEqual(w.taskbar_model.rowCount(), 1)
        w.set_view_options(float_split_enabled=False)
        self.assertEqual((w.quotes.float_page, w.quotes.taskbar_page), (0, 0))
        self.assertTrue(w.right_table.isHidden())
        self.assertEqual(w.model.rowCount(), 2)

    def test_unlimited_split_first_only_and_auto_paging(self):
        w = self.win
        w.show()
        w.set_view_options(float_split_enabled=True, float_max_rows=2,
                           float_page_mode="auto", float_page_interval=60)
        self.assertTrue(w.quotes.page_timers["float"].isActive())
        w.quotes.change_page("float", 1, automatic=True)
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (1, 0))
        self.assertEqual(w.model._rows[0][0], "标的5")
        w.set_view_options(float_page_mode="first")
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (2, 2))
        self.assertFalse(w.quotes.page_timers["float"].isActive())
        self.assertFalse(w.quotes.get_page("float").controls)
        w.set_view_options(float_paging_enabled=False)
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (3, 2))
        self.assertTrue(w.pager.isHidden())
        self.assertEqual(w.view_options.float_max_rows, 2)
        self.populate(0)
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (0, 0))
        self.populate(1)
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (1, 0))

    def test_taskbar_split_sync_keeps_independent_preferences_and_page_numbers(self):
        w = self.win
        self.populate(9)
        w.set_view_options(taskbar_sync_split=False, taskbar_split_enabled=True,
                           taskbar_split_separator=False, taskbar_rows=2, taskbar_page_mode="manual")
        w.quotes.change_page("taskbar", 1)
        w.set_view_options(taskbar_sync_split=True)
        self.assertEqual(w.view_options.split_settings("taskbar"), (False, True))
        self.assertEqual(w.quotes.taskbar_page, 0)
        w.set_view_options(float_split_enabled=True, float_split_separator=False)
        w.quotes.change_page("taskbar", 1)
        w.set_view_options(float_split_separator=True)
        self.assertEqual(w.quotes.taskbar_page, 1)
        self.assertEqual(w.quotes.float_page, 0)
        self.assertEqual(w.view_options.split_settings("taskbar"), (True, True))
        w.set_view_options(taskbar_sync_split=False)
        self.assertEqual(w.view_options.split_settings("taskbar"), (True, False))
        self.assertEqual(w.quotes.taskbar_page, 1)
        restored = ViewOptions.from_config(w.current_config())
        self.assertTrue(restored.taskbar_split_enabled)
        self.assertFalse(restored.taskbar_split_separator)

    def test_taskbar_renderer_places_halves_side_by_side_at_both_dpis(self):
        w = self.win
        w.set_view_options(taskbar_sync_split=False, taskbar_split_enabled=True,
                           taskbar_rows=3, taskbar_metrics=["name"])
        texts, lines = [], []

        class Recorder(QPainter):
            def drawText(self, *args):
                texts.append((args[0], args[-1]))
                return super().drawText(*args)

            def drawLine(self, *args):
                lines.append(args)
                return super().drawLine(*args)

        for dpi in (96, 192):
            for separator in (False, True):
                with self.subTest(dpi=dpi, separator=separator):
                    w.set_view_options(taskbar_split_separator=separator)
                    texts.clear()
                    lines.clear()
                    with patch("stockwidget.ui.floating.taskbar.QPainter", Recorder):
                        frame = render_taskbar(w, 72 * dpi // 96, dpi, max_width=480)
                    self.assertEqual([text for _, text in texts], [f"标的{i}" for i in range(1, 6)])
                    self.assertEqual(len({rect.x() for rect, _ in texts[:3]}), 1)
                    self.assertEqual(len({rect.x() for rect, _ in texts[3:]}), 1)
                    self.assertGreater(texts[3][0].x(), texts[0][0].right())
                    self.assertEqual(texts[0][0].y(), texts[3][0].y())
                    self.assertEqual(texts[1][0].y(), texts[4][0].y())
                    self.assertEqual(bool(lines), separator)
                    self.assertLessEqual(frame.width(), 480)

    def test_split_right_header_sorts_both_columns_and_retains_color_and_kline(self):
        w = self.win
        w.set_visible_metrics(["name", "price", "kline"])
        w.set_unicolor(False)
        w.set_view_options(float_split_enabled=True, float_paging_enabled=False)
        w.set_header_visible(True)
        w.show()
        self.app.processEvents()
        header = w.right_table.horizontalHeader()
        price = w.right_model._headers.index("现价")
        point = QPoint(header.sectionViewportPosition(price) + 5, header.height() // 2)
        QTest.mouseClick(header.viewport(), Qt.LeftButton, pos=point)
        self.assertEqual(w.quotes.sort_header, "现价")
        self.assertEqual([row[0] for row in w.model._rows], ["标的5", "标的4", "标的3"])
        self.assertEqual([row[0] for row in w.right_model._rows], ["标的2", "标的1"])
        self.assertEqual(w.right_model.index(0, price).data(Qt.ForegroundRole), w.up_color)
        self.assertIs(w.right_table.itemDelegateForColumn(2), w.right_k_delegate)
        self.assertTrue(w.table.horizontalHeader().isSortIndicatorShown())
        self.assertTrue(header.isSortIndicatorShown())
        self.assertEqual(w.table.columnWidth(price), w.right_table.columnWidth(price))
        self.assertEqual(w.table.height(), w.right_table.height())

    def test_split_columns_align_when_the_right_half_contains_longer_values(self):
        w = self.win
        w.quotes._last_full_rows[-1]["名称"] = "very long name in the right column"
        w.quotes._last_full_rows[-1]["现价"] = "1234567890.12"
        w.set_view_options(float_split_enabled=True, float_paging_enabled=False)
        w.quotes.reproject()
        w.show()
        self.app.processEvents()
        for column in range(w.model.columnCount()):
            self.assertEqual(w.table.columnWidth(column), w.right_table.columnWidth(column))
        self.assertEqual(w.table.width(), w.right_table.width())

    def test_hidden_separator_gap_and_empty_right_half_remain_draggable(self):
        w = self.win
        w.set_view_options(float_split_enabled=True, float_split_separator=False, float_max_rows=2)
        w.quotes.change_page("float", 1)
        w.show()
        self.app.processEvents()
        self.assertEqual(w.right_model.rowCount(), 0)
        self.assertTrue(w.split_separator.isHidden())
        gap = QPoint(w.table.x() + w.table.width() + 1, w.table.y() + 4)
        for target, point in ((w.panel, gap), (w.right_table.viewport(), QPoint(5, 5))):
            origin = w.pos()
            start = target.mapToGlobal(point)
            self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
            self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
            self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(40, 25), Qt.LeftButton, Qt.NoButton)
            self.assertEqual(w.pos(), origin + QPoint(40, 25))
            self.assertEqual(w.quotes.float_page, 1)

    def test_all_split_regions_share_click_drag_cancel_double_click_and_restore(self):
        w = self.win
        self.populate(9)
        controller = self.make_controller()
        w.set_header_visible(True)
        w.set_grid_visible(True)
        w.set_view_options(float_split_enabled=True, float_max_rows=2, taskbar_sync_split=True)
        for display in ("float", "both"):
            for mode in ("manual", "auto"):
                w.set_display_mode(display)
                w.set_view_options(float_page_mode=mode, float_page_interval=60)
                w.show()
                self.app.processEvents()
                regions = [(w.panel, QPoint(1, 1)), (w, QPoint(w.width() - 1, w.height() - 1)),
                           (w.split_separator, w.split_separator.rect().center())]
                regions += [(table.viewport(), QPoint(5, 5)) for table in w.float_tables]
                regions += [(table.horizontalHeader().viewport(), QPoint(5, 5)) for table in w.float_tables]
                regions += [(w.pager, QPoint(15, y)) for y in (3, w.pager.height() // 2, w.pager.height() - 3)]
                regions += [(w.hide_notice, QPoint(5, 5))]
                for target, point in regions:
                    with self.subTest(display=display, mode=mode, region=target.objectName(), point=point):
                        w.hide_notice.setText("行情已超过30秒未更新，5秒后隐藏")
                        w.hide_notice.show()
                        w._fit_to_contents()
                        self.app.processEvents()
                        origin = w.pos()
                        start = target.mapToGlobal(point)
                        page, sort = w.quotes.float_page, w.quotes.sort_header
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(1, 0), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(1, 0), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin)
                        page = w.quotes.float_page
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        QApplication.sendEvent(w, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                        self.assertEqual(w.pos(), origin)
                        self.mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.quotes.float_page, page)
                        self.assertEqual(w.quotes.sort_header, sort)
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(40, 25), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual(w.quotes.float_page, page)
                        self.assertEqual(w.quotes.sort_header, sort)
                        position = w.pos()
                        QTest.mouseDClick(target, Qt.LeftButton, pos=point)
                        self.assertFalse(w.widget_visible)
                        self.assertFalse(controller._active)
                        w.toggle_win()
                        self.app.processEvents()
                        self.assertEqual(w.pos(), position)
                        self.assertEqual(w.display_mode, display)
                        self.assertEqual(controller._active, display == "both")

    def test_boundary_and_edge_hide_preserve_all_region_interactions(self):
        w = self.win
        self.populate(9)
        controller = self.make_controller()
        w.set_header_visible(True)
        w.set_grid_visible(True)
        w.set_view_options(float_split_enabled=True, float_max_rows=2)
        w.move(80, 80)
        w.set_position_options(boundary_check_enabled=True, edge_hide_enabled=True)
        for display in ("float", "both"):
            for mode in ("manual", "auto"):
                w.set_display_mode(display)
                w.set_view_options(float_page_mode=mode, float_page_interval=60)
                self.app.processEvents()
                regions = [(w, QPoint(1, 1)), (w.panel, QPoint(1, 1)),
                           (w.split_separator, w.split_separator.rect().center())]
                regions += [(table.viewport(), QPoint(5, 5)) for table in w.float_tables]
                regions += [(table.horizontalHeader().viewport(), QPoint(5, 5)) for table in w.float_tables]
                regions += [(w.pager, QPoint(15, y)) for y in (3, w.pager.height() // 2, w.pager.height() - 3)]
                regions.append((w.hide_notice, QPoint(5, 5)))
                regions = [(target, point, None) for target, point in regions]
                regions += [(w.message_label, QPoint(5, 5), message)
                            for message in ("加载中…", "未选择标的", "网络请求失败")]
                for target, point, message in regions:
                    with self.subTest(display=display, mode=mode, region=target.objectName()):
                        if message:
                            w._show_message(message, is_error="失败" in message)
                        w.hide_notice.setText("隐藏倒计时")
                        w.hide_notice.show()
                        w.move(80, 80)
                        w._fit_to_contents()
                        origin = w.pos()
                        start = target.mapToGlobal(point)
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(1, 0), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(1, 0), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin)
                        page, sort = w.quotes.float_page, w.quotes.sort_header
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        QApplication.sendEvent(w, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                        self.assertEqual(w.pos(), origin)
                        self.mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(40, 25), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual((w.quotes.float_page, w.quotes.sort_header), (page, sort))
                        QTest.mouseDClick(target, Qt.LeftButton, pos=point)
                        self.assertFalse(w.widget_visible)
                        w.toggle_win()
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual(w.display_mode, display)
                        self.assertEqual(controller._active, display == "both")

    def test_first_only_ignores_manual_navigation(self):
        self.win.set_view_options(float_max_rows=2, float_page_mode="first", taskbar_rows=1)
        for surface in ("float", "taskbar"):
            self.win.quotes.change_page(surface, 1)
            self.assertEqual(getattr(self.win.quotes, f"{surface}_page"), 0)
        self.assertEqual(self.win.model.rowCount(), 2)
        self.assertEqual(self.win.taskbar_model.rowCount(), 1)
        self.assertTrue(self.win.pager.isHidden())

    def test_synced_paging_uses_float_mode_and_interval_with_own_row_limit(self):
        w = self.win
        w.show()
        with patch("sys.platform", "win32"):
            w.set_display_mode("both")
        w.set_view_options(taskbar_rows=4, taskbar_sync_paging=True,
                           taskbar_page_mode="manual", taskbar_page_interval=17,
                           float_max_rows=2, float_page_mode="auto", float_page_interval=11)
        self.assertEqual(w.view_options.page_settings("taskbar"), ("auto", 11))
        self.assertEqual(w.taskbar_model.rowCount(), 4)
        self.assertEqual(w.quotes.page_timers["taskbar"].interval(), 11000)
        self.assertTrue(w.quotes.page_timers["taskbar"].isActive())
        w.quotes.change_page("taskbar", 1)
        self.assertEqual(w.taskbar_model.index(0, 0).data(), "标的5")
        self.assertEqual(w.quotes.float_page, 0)
        w.set_view_options(float_paging_enabled=False)
        self.assertEqual(w.view_options.page_settings("taskbar"), ("first", 11))
        self.assertEqual(w.model.rowCount(), 5)
        self.assertEqual(w.taskbar_model.rowCount(), 4)
        self.assertFalse(w.quotes.get_page("taskbar").controls)
        self.assertFalse(w.quotes.page_timers["taskbar"].isActive())
        w.set_view_options(taskbar_sync_paging=False)
        self.assertEqual(w.view_options.page_settings("taskbar"), ("manual", 17))
        self.assertTrue(w.quotes.get_page("taskbar").controls)
        self.assertEqual(w.current_config()["taskbar_rows"], 4)

    def test_taskbar_opacity_applies_to_data_pager_and_message(self):
        w = self.win
        w.set_view_options(taskbar_sync_appearance=False, taskbar_page_mode="manual",
                           taskbar_opacity_pct=100)
        for message in (False, True):
            if message:
                w._show_message("暂无行情")
            def alpha(opacity):
                w.set_view_options(taskbar_opacity_pct=opacity)
                frame = render_taskbar(w, 64)
                return max(frame.pixelColor(x, y).alpha()
                           for y in range(frame.height()) for x in range(frame.width()))
            opaque, half, transparent = alpha(100), alpha(50), alpha(0)
            self.assertGreater(opaque, 150)
            self.assertGreater(half, 60)
            self.assertLessEqual(half, 129)
            self.assertEqual(transparent, 1)

    def test_appearance_sync_restores_all_independent_preferences(self):
        w = self.win
        w.set_view_options(taskbar_sync_appearance=False, taskbar_font_family="Arial",
                           taskbar_font_size=7, taskbar_color="#55bbdd",
                           taskbar_opacity_pct=42, taskbar_unicolor=False)
        independent = w.get_taskbar_appearance()
        self.assertEqual((independent[0].family(), independent[0].pointSize()), ("Arial", 7))
        self.assertEqual((independent[1].name(), *independent[2:]), ("#55bbdd", 42, False))
        self.assertEqual(w.taskbar_model.index(0, 1).data(Qt.ForegroundRole), w.up_color)
        w.set_view_options(taskbar_sync_appearance=True)
        w.set_fg_color(QColor("#ffaa11"))
        w.set_font_size(12)
        w.set_window_opacity_percent(65)
        self.assertEqual(w.get_taskbar_appearance(), (w.font, w.fg, 65, w.unicolor))
        w.set_view_options(taskbar_sync_appearance=False)
        self.assertEqual(w.get_taskbar_appearance(), independent)

    def test_disabling_float_paging_displays_all_rows_stops_timer_and_remembers_preferences(self):
        self.win.show()
        self.win.set_view_options(float_max_rows=2, float_page_mode="auto", float_page_interval=20)
        self.win.quotes.change_page("float", 1)
        self.assertTrue(self.win.quotes.page_timers["float"].isActive())
        self.win.set_view_options(float_paging_enabled=False)
        self.assertEqual(self.win.model.rowCount(), 5)
        self.assertFalse(self.win.quotes.get_page("float").controls)
        self.assertTrue(self.win.pager.isHidden())
        self.assertFalse(self.win.quotes.page_timers["float"].isActive())
        self.win.quotes.change_page("float", 1)
        self.assertEqual(self.win.quotes.float_page, 0)
        self.assertEqual(self.win.view_options.float_max_rows, 2)
        self.assertEqual(self.win.view_options.float_page_mode, "auto")
        self.assertEqual(self.win.view_options.float_page_interval, 20)
        self.win.set_view_options(float_paging_enabled=True)
        self.assertEqual(self.win.model.rowCount(), 2)
        self.assertTrue(self.win.quotes.page_timers["float"].isActive())

    def test_all_taskbar_metrics_and_live_sync_keep_independent_order(self):
        from stockwidget.core.quote_presentation import METRIC_SPECS
        all_metrics = [spec.metric_id for spec in METRIC_SPECS]
        self.win.set_view_options(taskbar_metrics=all_metrics)
        self.assertEqual(self.win.view_options.taskbar_metrics, all_metrics)
        self.assertEqual(self.win.taskbar_model.columnCount(), 11)
        self.assertFalse(render_taskbar(self.win, 44).isNull())
        independent = ["volume", "price", "name", "change"]
        self.win.set_view_options(taskbar_metrics=independent, taskbar_sync_metrics=True)
        self.win.set_visible_metrics(["change_pct", "name", "price", "volume"])
        self.assertEqual(self.win.taskbar_model._headers, self.win.model._headers)
        self.win.set_visible_metrics(["price", "change_pct"])
        self.assertEqual(self.win.get_surface_metrics("taskbar"), ["price", "change_pct"])
        self.assertEqual(self.win.taskbar_model._headers, ["现价", "涨幅"])
        self.win.set_view_options(taskbar_sync_metrics=False)
        self.assertEqual(self.win.get_surface_metrics("taskbar"), independent)
        self.assertEqual(self.win.taskbar_model._headers, ["成交量", "现价", "名称", "涨跌"])
        self.assertFalse(self.win.current_config()["taskbar_sync_metrics"])

    def test_sort_resets_pages_and_shared_order(self):
        self.win.set_view_options(float_max_rows=2, taskbar_page_mode="manual")
        self.win.quotes.change_page("float", 1)
        self.win.quotes.change_page("taskbar", 1)
        self.win.quotes.set_sort("现价", Qt.DescendingOrder)
        self.assertEqual((self.win.quotes.float_page, self.win.quotes.taskbar_page), (0, 0))
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(), "标的5")
        self.win.quotes.change_page("float", 1)
        self.win.quotes.change_page("taskbar", 1)
        self.win.set_watchlist({"sh600000": {"checked": True}})
        self.assertEqual((self.win.quotes.float_page, self.win.quotes.taskbar_page), (0, 0))

    def test_independent_metrics_colors_and_font_render(self):
        before = render_taskbar(self.win, 44)
        self.win.set_view_options(taskbar_metrics=["change", "price", "name", "kline"],
            taskbar_sync_appearance=False, taskbar_color="#55bbdd",
            taskbar_font_family="Arial", taskbar_font_size=6)
        self.assertEqual(self.win.taskbar_model._headers, ["涨跌", "现价", "名称", "K线"])
        self.assertEqual(self.win.model._headers, ["名称", "现价", "涨幅"])
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(Qt.ForegroundRole), QColor("#55bbdd"))
        self.assertNotEqual(before, render_taskbar(self.win, 44))
        self.win.set_fg_color(QColor("#ffaa11"))
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(Qt.ForegroundRole), QColor("#55bbdd"))
        self.win.set_view_options(taskbar_sync_appearance=True)
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(Qt.ForegroundRole), QColor("#ffaa11"))
        self.win.set_visible_metrics([])
        self.assertEqual(self.win.taskbar_model.columnCount(), 4)
        self.assertFalse(render_taskbar(self.win, 44).isNull())

    def test_auto_timers_do_not_restart_on_quotes_and_stop_when_hidden(self):
        self.win.set_view_options(float_max_rows=2, float_page_mode="auto", float_page_interval=1,
                                  taskbar_page_mode="auto", taskbar_page_interval=2)
        self.win.show()
        with patch("sys.platform", "win32"):
            self.win.set_display_mode("both")
        self.assertTrue(self.win.quotes.page_timers["float"].isActive())
        self.assertTrue(self.win.quotes.page_timers["taskbar"].isActive())
        for _ in range(4):
            QTest.qWait(300)
            self.populate()
        self.assertEqual(self.win.quotes.float_page, 1)
        self.assertEqual(self.win.quotes.taskbar_page, 0)
        self.win.hide_widget()
        self.assertFalse(self.win.quotes.page_timers["float"].isActive())
        self.assertFalse(self.win.quotes.page_timers["taskbar"].isActive())
        self.win.toggle_win()
        self.assertTrue(self.win.quotes.page_timers["taskbar"].isActive())
        self.win.set_display_mode("float")
        self.assertFalse(self.win.quotes.page_timers["taskbar"].isActive())

    def test_float_page_buttons_click_to_page_and_double_click_to_hide(self):
        self.win.set_view_options(float_max_rows=2)
        self.win.show()
        self.app.processEvents()
        position = self.win.pos()
        QTest.mouseClick(self.win.pager, Qt.LeftButton, pos=QPoint(15, self.win.pager.height() - 4))
        self.assertEqual(self.win.quotes.float_page, 1)
        self.assertEqual(self.win.pos(), position)
        self.assertIsNone(self.win._drag_pos)
        self.assertTrue(self.win.isVisible())
        QTest.mouseDClick(self.win.pager, Qt.LeftButton)
        self.assertFalse(self.win.widget_visible)

    def test_every_pager_region_drags_without_paging_in_manual_and_auto_modes(self):
        self.win.show()
        for mode in ("manual", "auto"):
            self.win.set_view_options(float_max_rows=2, float_page_mode=mode, float_page_interval=60)
            self.app.processEvents()
            pager = self.win.pager
            for y in (3, pager.height() // 2, pager.height() - 3):
                with self.subTest(mode=mode, y=y):
                    origin, page = self.win.pos(), self.win.quotes.float_page
                    start = pager.mapToGlobal(QPoint(15, y))
                    end = start + QPoint(45, 30)
                    self.mouse(pager, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                    self.mouse(pager, QEvent.MouseMove, end, Qt.NoButton, Qt.LeftButton)
                    self.mouse(pager, QEvent.MouseButtonRelease, end, Qt.LeftButton, Qt.NoButton)
                    self.assertEqual(self.win.pos(), origin + QPoint(45, 30))
                    self.assertEqual(self.win.quotes.float_page, page)
                    self.assertIsNone(self.win._drag_pos)

    def test_message_regions_drag_and_cancel_without_a_taskbar_controller(self):
        self.win.set_view_options(float_split_enabled=True)
        self.win.set_grid_visible(True)
        self.win.show()
        for text, kind in (("加载中…", "loading"), ("自选列表为空", "empty"), ("网络请求失败", "error")):
            with self.subTest(kind=kind):
                self.win._show_message(text, kind=kind)
                self.app.processEvents()
                target, origin = self.win.message_label, self.win.pos()
                start = target.mapToGlobal(QPoint(6, 6))
                self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                self.assertEqual(self.win.pos(), origin + QPoint(40, 25))
                QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                self.assertEqual(self.win.pos(), origin)
                self.mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
