import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, QPoint, QPointF, Qt
from PySide6.QtGui import QColor, QKeyEvent, QMouseEvent, QPainter
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.core.view_options import ViewOptions, column_ranges, page_slice
from stockwidget.platform.taskbar import NativeTaskbarWindow, TaskbarArea
from stockwidget.ui.taskbar import TaskbarController, render_taskbar
from stockwidget.ui.widget import FloatLabel


class PageMathTests(unittest.TestCase):
    def test_split_defaults_roundtrip_and_odd_item_distribution(self):
        defaults = ViewOptions.from_config({})
        self.assertFalse(defaults.float_split_enabled)
        self.assertTrue(defaults.taskbar_sync_split)
        options = ViewOptions.from_config({"float_split_enabled": True, "float_split_separator": False,
                                           "taskbar_sync_split": False, "taskbar_split_enabled": True})
        self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        for total in (0, 1, 2, 5, 6):
            ranges = column_ranges(total, True)
            left, right = (stop - start for start, stop in ranges)
            self.assertIn(left - right, (0, 1))
            self.assertEqual([i for start, stop in ranges for i in range(start, stop)], list(range(total)))
            self.assertEqual(column_ranges(total, False), ((0, total),))

    def test_unified_sync_preferences_migrate_legacy_appearance_and_paging(self):
        defaults = ViewOptions.from_config({})
        self.assertTrue(defaults.taskbar_sync_appearance)
        self.assertTrue(defaults.taskbar_sync_paging)
        legacy = ViewOptions.from_config({"taskbar_sync_font": True, "taskbar_sync_color": False,
                                          "taskbar_page_mode": "auto"})
        self.assertFalse(legacy.taskbar_sync_appearance)
        self.assertFalse(legacy.taskbar_sync_paging)
        explicit = ViewOptions.from_config({**legacy.to_config(), "taskbar_sync_appearance": True,
                                            "taskbar_sync_paging": True, "taskbar_opacity_pct": 150})
        self.assertTrue(explicit.taskbar_sync_appearance)
        self.assertTrue(explicit.taskbar_sync_paging)
        self.assertEqual(explicit.taskbar_opacity_pct, 100)
        self.assertEqual(ViewOptions.from_config({"taskbar_opacity_pct": -10}).taskbar_opacity_pct, 0)

    def test_legacy_enabled_modes_and_paging_migrate_without_overriding_explicit_switches(self):
        self.assertFalse(ViewOptions.from_config({}).taskbar_enabled)
        self.assertFalse(ViewOptions.from_config({}).float_paging_enabled)
        old = ViewOptions.from_config({"float_max_rows": 4, "display_mode": "taskbar"})
        self.assertTrue(old.float_paging_enabled)
        self.assertTrue(old.taskbar_enabled)
        self.assertTrue(old.taskbar_sync_metrics)
        explicit = ViewOptions.from_config({"float_max_rows": 4, "display_mode": "both",
                                            "float_paging_enabled": False, "taskbar_enabled": False})
        self.assertFalse(explicit.float_paging_enabled)
        self.assertFalse(explicit.taskbar_enabled)

    def test_last_page_and_clamping(self):
        page = page_slice(5, 3, "manual", 1)
        self.assertEqual((page.start, page.stop, page.count, page.index), (3, 5, 2, 1))
        self.assertTrue(page.controls)
        self.assertEqual(page_slice(2, 3, "manual", 9).index, 0)
        self.assertFalse(page_slice(0, 3, "auto").controls)

    def test_first_rows_and_unlimited(self):
        page = page_slice(5, 3, "first", 8)
        self.assertEqual((page.start, page.stop, page.index), (0, 3, 0))
        self.assertFalse(page.controls)
        self.assertEqual(page_slice(5, 0, "auto", 1).stop, 5)
        self.assertFalse(page_slice(5, 0, "auto").controls)

    def test_options_normalize_and_migrate(self):
        opts = ViewOptions.from_config({"taskbar_rows": 99, "float_max_rows": -1,
            "taskbar_page_interval": "bad", "taskbar_metrics": ["name", "price", "name", "kline", "volume"],
            "display_mode": "both"})
        self.assertEqual(opts.taskbar_rows, 4)
        self.assertEqual(opts.float_max_rows, 1)
        self.assertFalse(opts.float_paging_enabled)
        self.assertEqual(opts.taskbar_page_interval, 5)
        self.assertEqual(opts.taskbar_metrics, ["name", "price", "kline", "volume"])
        self.assertTrue(opts.taskbar_dual_open)
        self.assertTrue(opts.taskbar_enabled)
        self.assertFalse(opts.taskbar_sync_metrics)
        self.assertEqual(ViewOptions.from_config(opts.to_config()).to_config(), opts.to_config())


class PagingInteractionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.enterContext(patch.object(FloatLabel, "_refresh_from_function"))
        self.enterContext(patch("stockwidget.ui.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.widget.apply_click_through"))
        self.win = FloatLabel({"taskbar_enabled": True, "taskbar_sync_metrics": False,
                               "float_paging_enabled": True, "taskbar_sync_paging": False}, {})
        self.win._clear_message()
        self.controller = None
        self.populate()

    def populate(self, total=5):
        self.win._last_full_rows = [dict(zip(("名称", "现价", "涨幅", "涨跌"),
                                                   (f"标的{i}", str(i * 10), f"+{i}%", str(i)))) for i in range(1, total + 1)]
        self.win._last_color_roles = [{h: "up" for h in row} for row in self.win._last_full_rows]
        self.win._last_sort_values = [{"现价": i * 10} for i in range(1, total + 1)]
        self.win._reproject_cached_data()

    def tearDown(self):
        if self.controller:
            self.controller.close()
            delete(self.controller)
        delete(self.win)
        self.app.processEvents()

    def test_independent_pages_and_incomplete_last_page(self):
        self.win.set_view_options(float_max_rows=2, float_page_mode="manual",
                                  taskbar_rows=3, taskbar_page_mode="manual")
        self.win.change_page("taskbar", 1)
        self.assertEqual(self.win.taskbar_model.rowCount(), 2)
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(), "标的4")
        self.assertEqual(self.win.model.index(0, 0).data(), "标的1")
        self.win.change_page("float", 1)
        self.assertEqual(self.win.model.index(0, 0).data(), "标的3")
        self.assertEqual(self.win.taskbar_page, 1)
        self.populate()
        self.assertEqual((self.win.float_page, self.win.taskbar_page), (1, 1))
        self.populate(2)
        self.assertEqual((self.win.float_page, self.win.taskbar_page), (0, 0))

    def test_split_pages_limit_each_column_and_balance_the_last_page(self):
        w = self.win
        self.populate(7)
        w.set_view_options(float_split_enabled=True, float_max_rows=2,
                           taskbar_sync_split=True, taskbar_rows=3, taskbar_page_mode="manual")
        self.assertEqual(w.get_page("float").count, 2)
        self.assertEqual(w.model._rows[0][0], "标的1")
        self.assertEqual(w.right_model._rows[0][0], "标的3")
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (2, 2))
        self.assertEqual(w.taskbar_model.rowCount(), 6)
        w.change_page("float", 1)
        self.assertEqual([row[0] for row in w.model._rows], ["标的5", "标的6"])
        self.assertEqual([row[0] for row in w.right_model._rows], ["标的7"])
        w.change_page("taskbar", 1)
        self.assertEqual(w.taskbar_model._rows[0][0], "标的7")
        self.assertEqual(w.taskbar_model.rowCount(), 1)
        w.set_view_options(float_split_enabled=False)
        self.assertEqual((w.float_page, w.taskbar_page), (0, 0))
        self.assertTrue(w.right_table.isHidden())
        self.assertEqual(w.model.rowCount(), 2)

    def test_unlimited_split_first_only_and_auto_paging(self):
        w = self.win
        w.show()
        w.set_view_options(float_split_enabled=True, float_max_rows=2,
                           float_page_mode="auto", float_page_interval=60)
        self.assertTrue(w.page_timers["float"].isActive())
        w.change_page("float", 1, automatic=True)
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (1, 0))
        self.assertEqual(w.model._rows[0][0], "标的5")
        w.set_view_options(float_page_mode="first")
        self.assertEqual((w.model.rowCount(), w.right_model.rowCount()), (2, 2))
        self.assertFalse(w.page_timers["float"].isActive())
        self.assertFalse(w.get_page("float").controls)
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
        w.change_page("taskbar", 1)
        w.set_view_options(taskbar_sync_split=True)
        self.assertEqual(w.get_split_settings("taskbar"), (False, True))
        self.assertEqual(w.taskbar_page, 0)
        w.set_view_options(float_split_enabled=True, float_split_separator=False)
        w.change_page("taskbar", 1)
        w.set_view_options(float_split_separator=True)
        self.assertEqual(w.taskbar_page, 1)
        self.assertEqual(w.float_page, 0)
        self.assertEqual(w.get_split_settings("taskbar"), (True, True))
        w.set_view_options(taskbar_sync_split=False)
        self.assertEqual(w.get_split_settings("taskbar"), (True, False))
        self.assertEqual(w.taskbar_page, 1)
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
                    with patch("stockwidget.ui.taskbar.QPainter", Recorder):
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
        self.assertEqual(w.sort_header, "现价")
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
        w._last_full_rows[-1]["名称"] = "very long name in the right column"
        w._last_full_rows[-1]["现价"] = "1234567890.12"
        w.set_view_options(float_split_enabled=True, float_paging_enabled=False)
        w._reproject_cached_data()
        w.show()
        self.app.processEvents()
        for column in range(w.model.columnCount()):
            self.assertEqual(w.table.columnWidth(column), w.right_table.columnWidth(column))
        self.assertEqual(w.table.width(), w.right_table.width())

    def test_hidden_separator_gap_and_empty_right_half_remain_draggable(self):
        w = self.win
        w.set_view_options(float_split_enabled=True, float_split_separator=False, float_max_rows=2)
        w.change_page("float", 1)
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
            self.assertEqual(w.float_page, 1)

    def test_all_split_regions_share_click_drag_cancel_double_click_and_restore(self):
        w = self.win
        self.populate(9)
        controller = self.make_controller()
        w.set_header_visible(True)
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
                        page, sort = w.float_page, w.sort_header
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(1, 0), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(1, 0), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin)
                        page = w.float_page
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        QApplication.sendEvent(w, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                        self.assertEqual(w.pos(), origin)
                        self.mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.float_page, page)
                        self.assertEqual(w.sort_header, sort)
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(40, 25), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual(w.float_page, page)
                        self.assertEqual(w.sort_header, sort)
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
                        page, sort = w.float_page, w.sort_header
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        QApplication.sendEvent(w, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                        self.assertEqual(w.pos(), origin)
                        self.mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
                        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseMove, start + QPoint(40, 25), Qt.NoButton, Qt.LeftButton)
                        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(40, 25), Qt.LeftButton, Qt.NoButton)
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual((w.float_page, w.sort_header), (page, sort))
                        QTest.mouseDClick(target, Qt.LeftButton, pos=point)
                        self.assertFalse(w.widget_visible)
                        w.toggle_win()
                        self.assertEqual(w.pos(), origin + QPoint(40, 25))
                        self.assertEqual(w.display_mode, display)
                        self.assertEqual(controller._active, display == "both")

    def test_native_taskbar_drag_uses_full_size_when_float_is_a_strip(self):
        w = self.win
        controller = self.make_controller()
        w.move(80, 80)
        w.set_position_options(boundary_check_enabled=True, edge_hide_enabled=True)
        w.set_display_mode("both")
        self.app.processEvents()
        screen = self.app.primaryScreen().geometry()
        with patch("stockwidget.ui.position_controller.QCursor.pos", return_value=QPoint(99999, 99999)):
            w.move(screen.right() - w.width() + 1, 100)
            self.assertEqual(w.width(), 5)
            full = w.position_controller.full_geometry()
            controller._dpi = 192
            start = QPoint(400, screen.bottom())
            controller._dispatch_pointer("press", 80, 10, start, True)
            self.assertEqual(w._drag_pos.x(), 40)
            self.assertEqual(w.geometry(), full)
            controller._dispatch_pointer("move", 80, 10, start - QPoint(50, 50), False)
            controller._dispatch_pointer("cancel", 80, 10, start, False)
            self.assertEqual(w.position_controller.full_geometry(), full)
            self.assertEqual(w.width(), 5)
            self.assertEqual(w.display_mode, "both")
            self.assertTrue(controller._active)

    def test_first_only_ignores_manual_navigation(self):
        self.win.set_view_options(float_max_rows=2, float_page_mode="first", taskbar_rows=1)
        for surface in ("float", "taskbar"):
            self.win.change_page(surface, 1)
            self.assertEqual(getattr(self.win, f"{surface}_page"), 0)
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
        self.assertEqual(w.get_page_settings("taskbar"), ("auto", 11))
        self.assertEqual(w.taskbar_model.rowCount(), 4)
        self.assertEqual(w.page_timers["taskbar"].interval(), 11000)
        self.assertTrue(w.page_timers["taskbar"].isActive())
        w.change_page("taskbar", 1)
        self.assertEqual(w.taskbar_model.index(0, 0).data(), "标的5")
        self.assertEqual(w.float_page, 0)
        w.set_view_options(float_paging_enabled=False)
        self.assertEqual(w.get_page_settings("taskbar"), ("first", 11))
        self.assertEqual(w.model.rowCount(), 5)
        self.assertEqual(w.taskbar_model.rowCount(), 4)
        self.assertFalse(w.get_page("taskbar").controls)
        self.assertFalse(w.page_timers["taskbar"].isActive())
        w.set_view_options(taskbar_sync_paging=False)
        self.assertEqual(w.get_page_settings("taskbar"), ("manual", 17))
        self.assertTrue(w.get_page("taskbar").controls)
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
        self.win.change_page("float", 1)
        self.assertTrue(self.win.page_timers["float"].isActive())
        self.win.set_view_options(float_paging_enabled=False)
        self.assertEqual(self.win.model.rowCount(), 5)
        self.assertFalse(self.win.get_page("float").controls)
        self.assertTrue(self.win.pager.isHidden())
        self.assertFalse(self.win.page_timers["float"].isActive())
        self.win.change_page("float", 1)
        self.assertEqual(self.win.float_page, 0)
        self.assertEqual(self.win.view_options.float_max_rows, 2)
        self.assertEqual(self.win.view_options.float_page_mode, "auto")
        self.assertEqual(self.win.view_options.float_page_interval, 20)
        self.win.set_view_options(float_paging_enabled=True)
        self.assertEqual(self.win.model.rowCount(), 2)
        self.assertTrue(self.win.page_timers["float"].isActive())

    def test_all_taskbar_metrics_and_live_sync_keep_independent_order(self):
        from stockwidget.core.metric_layout import METRIC_SPECS
        all_metrics = [spec.metric_id for spec in METRIC_SPECS]
        self.win.set_view_options(taskbar_metrics=all_metrics)
        self.assertEqual(self.win.view_options.taskbar_metrics, all_metrics)
        self.assertEqual(self.win.taskbar_model.columnCount(), 12)
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
        self.win.change_page("float", 1)
        self.win.change_page("taskbar", 1)
        self.win.set_sort("现价", Qt.DescendingOrder)
        self.assertEqual((self.win.float_page, self.win.taskbar_page), (0, 0))
        self.assertEqual(self.win.taskbar_model.index(0, 0).data(), "标的5")

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
        self.assertTrue(self.win.page_timers["float"].isActive())
        self.assertTrue(self.win.page_timers["taskbar"].isActive())
        for _ in range(4):
            QTest.qWait(300)
            self.populate()
        self.assertEqual(self.win.float_page, 1)
        self.assertEqual(self.win.taskbar_page, 0)
        self.win.hide_widget()
        self.assertFalse(self.win.page_timers["float"].isActive())
        self.assertFalse(self.win.page_timers["taskbar"].isActive())
        self.win.toggle_win()
        self.assertTrue(self.win.page_timers["taskbar"].isActive())
        self.win.set_display_mode("float")
        self.assertFalse(self.win.page_timers["taskbar"].isActive())

    def test_float_page_buttons_click_to_page_and_double_click_to_hide(self):
        self.win.set_view_options(float_max_rows=2)
        self.win.show()
        self.app.processEvents()
        position = self.win.pos()
        QTest.mouseClick(self.win.pager, Qt.LeftButton, pos=QPoint(15, self.win.pager.height() - 4))
        self.assertEqual(self.win.float_page, 1)
        self.assertEqual(self.win.pos(), position)
        self.assertIsNone(self.win._drag_pos)
        self.assertTrue(self.win.isVisible())
        QTest.mouseDClick(self.win.pager, Qt.LeftButton)
        self.assertFalse(self.win.widget_visible)

    def mouse(self, target, kind, global_pos, button, buttons):
        local = target.mapFromGlobal(global_pos)
        event = QMouseEvent(kind, QPointF(local), QPointF(global_pos), button, buttons, Qt.NoModifier)
        QApplication.sendEvent(target, event)

    def test_every_pager_region_drags_without_paging_in_manual_and_auto_modes(self):
        self.win.show()
        for mode in ("manual", "auto"):
            self.win.set_view_options(float_max_rows=2, float_page_mode=mode, float_page_interval=60)
            self.app.processEvents()
            pager = self.win.pager
            for y in (3, pager.height() // 2, pager.height() - 3):
                with self.subTest(mode=mode, y=y):
                    origin, page = self.win.pos(), self.win.float_page
                    start = pager.mapToGlobal(QPoint(15, y))
                    end = start + QPoint(45, 30)
                    self.mouse(pager, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                    self.mouse(pager, QEvent.MouseMove, end, Qt.NoButton, Qt.LeftButton)
                    self.mouse(pager, QEvent.MouseButtonRelease, end, Qt.LeftButton, Qt.NoButton)
                    self.assertEqual(self.win.pos(), origin + QPoint(45, 30))
                    self.assertEqual(self.win.float_page, page)
                    self.assertIsNone(self.win._drag_pos)

    def test_message_regions_drag_and_cancel_without_a_taskbar_controller(self):
        self.win.set_view_options(float_split_enabled=True)
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

    def make_controller(self):
        self.enterContext(patch("stockwidget.ui.taskbar.QApplication.platformName", return_value="windows"))
        self.enterContext(patch("stockwidget.ui.widget.sys.platform", "win32"))
        self.enterContext(patch("stockwidget.ui.taskbar.find_taskbar", return_value=TaskbarArea(1, 1280, 48, 900)))
        self.native = Mock()
        self.native.poll_pointer.return_value = None
        self.enterContext(patch("stockwidget.ui.taskbar.NativeTaskbarWindow", return_value=self.native))
        self.over = self.enterContext(patch("stockwidget.ui.taskbar.cursor_over_taskbar", return_value=False))
        self.controller = TaskbarController(self.win, Mock())
        self.controller.apply_mode()
        return self.controller

    def test_taskbar_double_click_hides_and_toggle_restores_the_same_location(self):
        controller = self.make_controller()
        self.win.set_display_mode("taskbar")
        controller._click(100, 20)
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

        controller._double_click(100, 20)
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.populate()
        self.win.set_view_options(taskbar_rows=3)
        controller.refresh()
        self.assertFalse(controller._active)
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.timer.isActive())
        self.win.toggle_win()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())

    def test_disabled_taskbar_cannot_be_started_by_drag_or_display_position_menu(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_enabled=False)
        self.win.move(100, 100)
        target, origin = self.win.table.viewport(), self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.over.reset_mock()
        self.over.return_value = True
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseMove, start + QPoint(0, 600), Qt.NoButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(0, 600), Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.pos(), origin + QPoint(0, 600))
        self.assertFalse(controller._active)
        self.assertEqual(self.win.display_mode, "float")
        self.assertFalse(self.win.taskbar_preview_active)
        self.over.assert_not_called()
        self.win.set_display_mode("taskbar")
        self.assertEqual(self.win.display_mode, "float")
        menu = self.win.build_context_menu()
        try:
            actions = {action.text(): action for action in menu.actions()}
            modes = actions["显示位置"].menu()
            self.assertEqual([action.isEnabled() for action in modes.actions()], [True, False, False])
        finally:
            delete(menu)
        self.win.set_view_options(taskbar_enabled=True)
        self.win.set_display_mode("taskbar")
        self.assertTrue(controller._active)
        self.win.set_view_options(taskbar_enabled=False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)

    def test_disabling_taskbar_during_docking_cancels_preview_and_keeps_float(self):
        controller = self.make_controller()
        origin = self.win.pos()
        start = origin + QPoint(5, 5)
        self.over.return_value = True
        self.win.begin_drag(start)
        self.win.move_drag(start + QPoint(0, 200))
        self.assertTrue(controller._active)
        self.win.set_view_options(taskbar_enabled=False)
        self.win.finish_drag()
        self.assertFalse(controller._active)
        self.assertFalse(controller._dragging)
        self.assertFalse(self.win.taskbar_preview_active)
        self.assertEqual(self.win.display_mode, "float")
        self.assertEqual(self.win.pos(), origin)
        self.assertFalse(controller.timer.isActive())

    def test_taskbar_page_click_only_changes_taskbar_page(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        rect = controller._pager_rect
        controller._click(rect.center().x(), rect.bottom() - 2)
        self.app.processEvents()
        self.assertEqual(self.win.taskbar_page, 1)
        controller._double_click(rect.center().x(), rect.center().y())
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.isVisible())

    def test_dual_mode_drag_never_checks_docking_or_hides_taskbar(self):
        controller = self.make_controller()
        self.win.set_view_options(float_split_enabled=True)
        self.win.set_display_mode("both")
        self.win.move(100, 100)
        target = self.win.right_table.viewport()
        origin = self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.over.reset_mock()
        self.native.hide.reset_mock()
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        for delta, over in ((QPoint(90, 30), False), (QPoint(90, 650), True), (QPoint(200, 90), False)):
            self.over.return_value = over
            self.mouse(target, QEvent.MouseMove, start + delta, Qt.NoButton, Qt.LeftButton)
            self.assertEqual(self.win.pos(), origin + delta)
            self.assertEqual(self.win.display_mode, "both")
            self.assertTrue(self.win.isVisible())
            self.assertTrue(controller._active)
            self.assertFalse(self.win.taskbar_preview_active)
        self.mouse(target, QEvent.MouseButtonRelease, start + delta, Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "both")
        self.over.assert_not_called()
        self.native.hide.assert_not_called()
        # Esc cancellation also uses the same free-drag path.
        origin = self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseMove, start + QPoint(50, 30), Qt.NoButton, Qt.LeftButton)
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertTrue(controller._active)
        self.native.hide.assert_not_called()

    def test_taskbar_pager_drag_uses_same_transition_without_flipping_page(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        rect = controller._pager_rect
        press = QPoint(700, 760)
        outside = press - QPoint(0, 150)
        controller._dispatch_pointer("press", rect.center().x(), rect.bottom() - 2, press, True)
        controller._dispatch_pointer("move", rect.center().x(), -130, outside, False)
        controller._dispatch_pointer("release", rect.center().x(), -130, outside, False)
        self.assertEqual(self.win.taskbar_page, 0)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())

    def test_native_dual_drag_preserves_taskbar_and_cancel_cleans_capture(self):
        controller = self.make_controller()
        self.win.set_display_mode("both")
        origin = self.win.pos()
        press = QPoint(700, 760)
        self.native.hide.reset_mock()
        controller._dispatch_pointer("press", 100, 20, press, True)
        controller._dispatch_pointer("move", 110, -130, press - QPoint(0, 150), False)
        self.assertTrue(controller._active)
        self.assertEqual(self.win.display_mode, "both")
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertFalse(controller.input_timer.isActive())
        self.assertIsNone(controller._native_press)
        self.native.release_pointer.assert_called()
        self.native.hide.assert_not_called()

    def test_dual_mode_double_click_hides_both_and_toggle_restores_both(self):
        controller = self.make_controller()
        self.win.set_display_mode("both")
        controller._double_click(100, 20)
        self.app.processEvents()
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertFalse(self.win.widget_visible)
        self.win.toggle_win()
        self.assertTrue(self.win.isVisible())
        self.assertTrue(controller._active)
        self.win.mouseDoubleClickEvent(QMouseEvent(QEvent.MouseButtonDblClick, QPointF(5, 5), QPointF(5, 5),
                                                   Qt.LeftButton, Qt.LeftButton, Qt.NoModifier))
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.win.toggle_win()
        self.assertEqual(self.win.display_mode, "both")

    def test_native_drag_out_uses_shared_drag_and_can_redock_or_cancel(self):
        controller = self.make_controller()
        self.win.set_display_mode("taskbar")
        original = self.win.pos()
        press, outside = QPoint(700, 760), QPoint(710, 500)
        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertEqual(self.win.pos(), outside - self.win._drag_pos)
        controller._dispatch_pointer("cancel", 0, 0, outside, False)
        self.assertEqual(self.win.pos(), original)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())

        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        controller._dispatch_pointer("move", 100, 15, press, True)
        self.assertTrue(controller._active)
        controller._dispatch_pointer("release", 100, 15, press, True)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        controller._dispatch_pointer("release", 110, -245, outside, False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertIsNone(self.win._drag_pos)

    def test_native_small_movement_and_stationary_page_click_never_undock(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        original = self.win.pos()
        press = QPoint(700, 760)
        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 101, 15, press + QPoint(1, 0), True)
        controller._dispatch_pointer("release", 101, 15, press + QPoint(1, 0), True)
        self.assertEqual(self.win.pos(), original)
        self.assertEqual(self.win.display_mode, "taskbar")
        rect = controller._pager_rect
        controller._dispatch_pointer("press", rect.center().x(), rect.bottom() - 2, press, True)
        controller._dispatch_pointer("move", rect.center().x() + 1, rect.bottom() - 2, press + QPoint(1, 0), True)
        controller._dispatch_pointer("release", rect.center().x() + 1, rect.bottom() - 2, press + QPoint(1, 0), True)
        self.assertEqual(self.win.taskbar_page, 1)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

    def test_shared_menu_has_independent_metrics_and_common_sort_and_hide(self):
        self.make_controller()
        self.win.set_display_mode("taskbar")
        menu = self.win.build_context_menu("taskbar")
        try:
            actions = {action.text(): action for action in menu.actions()}
            self.assertIn("排序", actions)
            self.assertIn("设置…", actions)
            metric_actions = {action.text(): action for action in actions["显示指标"].menu().actions()}
            self.assertTrue(metric_actions["成交量"].isEnabled())
            metric_actions["涨幅"].trigger()
            self.assertEqual(self.win.view_options.taskbar_metrics, ["name", "price"])
            self.assertEqual(self.win.visible_metrics, ["name", "price", "change_pct"])
            actions["隐藏"].trigger()
            self.assertFalse(self.win.widget_visible)
        finally:
            delete(menu)

    def test_drag_preview_leave_and_commit(self):
        controller = self.make_controller()
        original = self.win.pos()
        self.win._drag_start_window_pos = original
        controller.drag_started()
        self.assertFalse(controller._active)
        self.over.return_value = True
        controller.drag_moved()
        self.assertTrue(controller._active)
        self.assertTrue(self.win.isVisible())
        self.assertEqual(self.win.display_mode, "float")
        self.over.return_value = False
        controller.drag_moved()
        self.assertFalse(controller._active)
        self.over.return_value = True
        controller.drag_moved()
        self.win.move(40, 750)
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())
        self.assertEqual(self.win.pos(), original)

    def test_drag_cancel_and_dual_commit(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_dual_open=True)
        origin = self.win.pos()
        self.win._drag_start_window_pos = origin
        controller.drag_started()
        self.over.return_value = True
        controller.drag_moved()
        self.win.move(0, 700)
        controller.drag_finished(False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertEqual(self.win.pos(), origin)
        controller.drag_started()
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "both")
        self.assertTrue(self.win.isVisible())
        self.assertTrue(controller._active)

    def test_drag_events_preview_dock_and_escape_without_a_controller_shortcut(self):
        controller = self.make_controller()
        self.win.move(100, 100)
        origin = self.win.pos()
        target = self.win.table.viewport()

        def mouse(kind, global_pos, button, buttons):
            local = target.mapFromGlobal(global_pos)
            event = QMouseEvent(kind, QPointF(local), QPointF(global_pos), button, buttons, Qt.NoModifier)
            QApplication.sendEvent(target, event)

        start = target.mapToGlobal(QPoint(3, 3))
        mouse(QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.over.return_value = True
        mouse(QEvent.MouseMove, start + QPoint(0, 100), Qt.NoButton, Qt.LeftButton)
        self.assertTrue(controller._active)
        self.assertTrue(self.win.isVisible())
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertFalse(controller._active)
        mouse(QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "float")

        mouse(QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        mouse(QEvent.MouseMove, start + QPoint(0, 100), Qt.NoButton, Qt.LeftButton)
        mouse(QEvent.MouseButtonRelease, start + QPoint(0, 100), Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())
        self.assertEqual(self.win.pos(), origin)

    def test_failed_drag_attachment_never_hides_float(self):
        controller = self.make_controller()
        self.native.present.side_effect = OSError("模拟失败")
        self.win._drag_start_window_pos = self.win.pos()
        controller.drag_started()
        self.over.return_value = True
        controller.drag_moved()
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())


class NativeInputTests(unittest.TestCase):
    def test_polling_finishes_background_drag_once_and_supports_escape(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native.hwnd, native._pointer_down, native._ignore_release = 1, True, False
        native._last_pointer_position = (100, 20)
        native.user = Mock()
        native.user.GetCapture.return_value = 1
        native.user.GetCursorPos.side_effect = lambda pointer: (
            setattr(pointer._obj, "x", 110), setattr(pointer._obj, "y", -150), True)[-1]
        native.user.ScreenToClient.return_value = True
        native.user.GetAsyncKeyState.side_effect = lambda key: 0x8000 if key == 1 else 0
        self.assertEqual(native.poll_pointer(), ("move", 110, -150))
        self.assertIsNone(native.poll_pointer())
        native.user.GetAsyncKeyState.return_value = 0
        native.user.GetAsyncKeyState.side_effect = None
        self.assertEqual(native.poll_pointer(), ("release", 110, -150))
        self.assertIsNone(native.poll_pointer())
        self.assertTrue(native._ignore_release)
        native._pointer_down = True
        native.user.GetAsyncKeyState.side_effect = lambda key: 0x8000 if key == 0x1B else 0
        self.assertEqual(native.poll_pointer(), ("cancel", 110, -150))

    def test_double_click_sequence_does_not_dispatch_two_single_clicks(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native._ignore_release = False
        native._pointer_down = False
        native.hwnd = 1
        native._on_pointer = Mock()
        native.user = Mock()
        xy = (20 << 16) | 100
        for message in (0x201, 0x202, 0x203, 0x202):
            native._wndproc(1, message, 0, xy)
        self.assertEqual([call.args for call in native._on_pointer.call_args_list],
                         [("press", 100, 20), ("release", 100, 20), ("double_click", 100, 20)])

    def test_native_capture_and_cancel_emit_pointer_events_only(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native.hwnd, native._pointer_down, native._ignore_release = 1, False, False
        native._last_frame = None
        native._on_pointer, native.user = Mock(), Mock()
        native.user.GetCapture.return_value = 1
        native._wndproc(1, 0x201, 1, (20 << 16) | 100)
        native.user.SetCapture.assert_called_once_with(1)
        native._wndproc(1, 0x200, 1, ((-100 & 0xffff) << 16) | 120)
        native.hide()
        native.user.DestroyWindow.assert_not_called()
        native._wndproc(1, 0x215, 0, 0)
        self.assertFalse(native._pointer_down)
        self.assertEqual([call.args[0] for call in native._on_pointer.call_args_list], ["press", "move", "cancel"])
