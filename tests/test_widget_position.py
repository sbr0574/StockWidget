import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, QPoint, QPointF, QRect, Qt
from PySide6.QtGui import QEnterEvent, QKeyEvent, QMouseEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.ui.settings_dialog import SettingsDialog
from stockwidget.ui.widget import FloatLabel


class WidgetPositionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self.refresh = self.enterContext(patch.object(FloatLabel, "_refresh_from_function"))
        self.enterContext(patch("stockwidget.ui.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.widget.apply_click_through"))
        self.pointer = self.enterContext(patch("stockwidget.ui.position_controller.QCursor.pos",
                                             return_value=QPoint(99999, 99999)))
        self.windows = []
        self.screen = self.app.primaryScreen().geometry()

    def tearDown(self):
        for window in reversed(self.windows):
            delete(window)
        self.app.processEvents()

    def window(self, **config):
        window = FloatLabel(config, {})
        self.windows.append(window)
        window.show()
        self.app.processEvents()
        self.refresh.reset_mock()
        return window

    def edge_position(self, window, edge):
        x, y = self.screen.left() + 80, self.screen.top() + 100
        if edge == "left":
            x = self.screen.left()
        elif edge == "right":
            x = self.screen.right() - window.width() + 1
        elif edge == "top":
            y = self.screen.top()
        else:
            y = self.screen.bottom() - window.height() + 1
        return QPoint(x, y)

    def test_default_off_restores_screen_overflow_verbatim_and_drag_allows_it(self):
        for point in (QPoint(self.screen.left() - 30, 100),
                      QPoint(self.screen.right() - 10, self.screen.bottom() - 10),
                      QPoint(-5000, -5000)):
            with self.subTest(point=point):
                window = self.window(pos={"x": point.x(), "y": point.y()})
                self.assertFalse(window.boundary_check_enabled)
                self.assertFalse(window.edge_hide_enabled)
                self.assertEqual(window.pos(), point)
                start = point + QPoint(10, 10)
                window.begin_drag(start)
                window.move_drag(start - QPoint(100, 100))
                window.finish_drag()
                self.assertEqual(window.pos(), point - QPoint(100, 100))

    def test_boundary_on_clamps_startup_and_contents_growth(self):
        window = self.window(boundary_check_enabled=True,
                             pos={"x": self.screen.right() - 10, "y": self.screen.bottom() - 10})
        self.assertTrue(self.screen.contains(window.geometry()))
        window._show_message("窗口内容变宽" * 8)
        self.app.processEvents()
        self.assertTrue(self.screen.contains(window.geometry()))
        window.set_font_size(15)
        self.app.processEvents()
        self.assertTrue(self.screen.contains(window.geometry()))

    def test_quotes_resize_immediately_for_longer_shorter_content_and_row_counts(self):
        for split in (False, True):
            with self.subTest(split=split):
                window = self.window(visible_metrics=["name", "price"], name_visible=True,
                                     float_split_enabled=split)
                window.watchlist = {str(i): {"checked": True} for i in range(6)}
                def quotes(count, name, price):
                    rows = {str(i): {"name": name, "price": price} for i in range(count)}
                    with patch("stockwidget.ui.widget.format_quote", side_effect=lambda data, *args, **kwargs: (
                            {"名称": data["name"], "现价": data["price"]}, {}, {})):
                        window._process_data((True, rows, None))
                    for table in window.float_tables if split else (window.table,):
                        bounds = table.rect().translated(table.mapTo(window, QPoint()))
                        self.assertTrue(window.rect().contains(bounds))
                    return window.size()

                # 不处理后续事件或等待下一次刷新，直接检查本轮内容对应的外框。
                compact = quotes(1, "A", "1.23")
                expanded = quotes(6, "A longer stock name", "123456.78")
                self.assertGreater(expanded.width(), compact.width())
                self.assertGreater(expanded.height(), compact.height())
                self.assertEqual(quotes(1, "A", "1.23"), compact)

    def test_error_and_recovery_resize_in_the_same_data_update(self):
        window = self.window(visible_metrics=["price"], name_visible=False)
        window.watchlist = {"test": {"checked": True}}
        with patch("stockwidget.ui.widget.format_quote", return_value=({"现价": "1.23"}, {}, {})):
            window._process_data((True, {"test": {}}, None))
            compact = window.size()
            window._process_data((False, None, "Quote request failed; retrying the data source"))
            self.assertGreater(window.width(), compact.width())
            self.assertGreater(window.height(), compact.height())
            window._process_data((True, {"test": {}}, None))
            self.assertEqual(window.size(), compact)

    def test_right_edge_stays_attached_while_quotes_grow_and_shrink(self):
        window = self.window(visible_metrics=["name", "price"], name_visible=True,
                             boundary_check_enabled=True, edge_hide_enabled=True)
        window.watchlist = {"test": {"checked": True}}
        window.move(80, 100)
        def quotes(name, price):
            with patch("stockwidget.ui.widget.format_quote", return_value=(
                    {"名称": name, "现价": price}, {}, {})):
                window._process_data((True, {"test": {}}, None))

        quotes("A", "1.23")
        compact = window.size()
        window.move(self.edge_position(window, "right"))
        controller = window.position_controller
        self.pointer.return_value = QPoint(self.screen.right(), controller.full_geometry().center().y())
        controller.check_pointer()
        quotes("Longer name", "1234.56")
        self.assertFalse(controller.collapsed)
        self.assertGreater(window.width(), compact.width())
        self.assertEqual(window.geometry().right(), self.screen.right())
        quotes("A", "1.23")
        self.assertFalse(controller.collapsed)
        self.assertEqual(window.size(), compact)
        self.assertEqual(window.geometry().right(), self.screen.right())
        self.pointer.return_value = QPoint(99999, 99999)
        controller.check_pointer()
        self.assertTrue(controller.collapsed)
        self.assertEqual(window.width(), 5)
        self.assertEqual(window.geometry().right(), self.screen.right())

    def test_boundary_drag_clamps_and_escape_restores_position(self):
        window = self.window(boundary_check_enabled=True)
        origin = QPoint(100, 100)
        window.move(origin)
        start = origin + QPoint(10, 10)
        window.begin_drag(start)
        window.move_drag(QPoint(self.screen.right() + 100, self.screen.bottom() + 100))
        self.assertTrue(self.screen.contains(window.geometry()))
        QApplication.sendEvent(window, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(window.pos(), origin)
        self.assertIsNone(window._drag_pos)

    def test_three_edges_shrink_pause_hover_restore_and_save_full_position(self):
        for edge in ("left", "top", "right"):
            with self.subTest(edge=edge):
                window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
                window.move(self.edge_position(window, edge))
                controller = window.position_controller
                full = controller.full_geometry()
                self.assertTrue(controller.collapsed)
                self.assertEqual(window.height() if edge == "top" else window.width(), 5)
                self.assertTrue(window.widget_visible)
                self.assertTrue(self.screen.contains(window.geometry()))
                self.assertFalse(window.timer.isActive())
                self.assertEqual(window.current_config()["pos"], {"x": full.x(), "y": full.y()})
                self.pointer.return_value = window.geometry().center()
                controller.check_pointer()
                self.assertFalse(controller.collapsed)
                self.assertEqual(window.geometry(), full)
                self.assertTrue(window.timer.isActive())
                self.pointer.return_value = QPoint(99999, 99999)
                controller.check_pointer()
                self.assertTrue(controller.collapsed)
                self.assertFalse(window.timer.isActive())

    def test_bottom_does_not_hide_and_disabling_boundary_expands_and_disables_edge(self):
        window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
        window.move(self.edge_position(window, "bottom"))
        self.assertFalse(window.position_controller.collapsed)
        window.move(self.edge_position(window, "right"))
        full = window.position_controller.full_geometry()
        window.set_position_options(boundary_check_enabled=False)
        self.assertFalse(window.edge_hide_enabled)
        self.assertFalse(window.position_controller.collapsed)
        self.assertEqual(window.geometry(), full)
        self.assertTrue(window.timer.isActive())
        window.move(-30, -20)
        self.assertEqual(window.pos(), QPoint(-30, -20))
        window.set_position_options(edge_hide_enabled=True)
        self.assertFalse(window.edge_hide_enabled)

    def test_collapsed_image_exposes_opposite_edge_and_expansion_restores_content(self):
        for edge in ("left", "top", "right"):
            with self.subTest(edge=edge):
                window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
                window.panel.setStyleSheet(
                    "QWidget#panel { border-radius: 5px; background: #123456; "
                    "border-left: 5px solid #ff0000; border-right: 5px solid #0000ff; "
                    "border-top: 5px solid #ffff00; border-bottom: 5px solid #00ff00; }")
                window.move(80, 100)
                origin = window.panel.pos()
                complete = window.grab().toImage()
                ratio = window.devicePixelRatioF()
                strip = round(5 * ratio)
                if edge == "left":
                    crop = QRect(complete.width() - strip, 0, strip, complete.height())
                elif edge == "top":
                    crop = QRect(0, complete.height() - strip, complete.width(), strip)
                else:
                    crop = QRect(0, 0, strip, complete.height())
                expected = complete.copy(crop)
                window.move(self.edge_position(window, edge))
                self.assertTrue(window.position_controller.collapsed)
                self.assertEqual(window.grab().toImage(), expected)
                self.pointer.return_value = window.geometry().center()
                window.position_controller.check_pointer()
                self.assertEqual(window.panel.pos(), origin)
                self.assertEqual(window.grab().toImage(), complete)
                self.pointer.return_value = QPoint(99999, 99999)

    def test_screen_extreme_hover_stays_expanded_until_pointer_leaves_full_window(self):
        for edge in ("left", "top", "right"):
            with self.subTest(edge=edge):
                window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
                window.move(self.edge_position(window, edge))
                controller = window.position_controller
                full = controller.full_geometry()
                extreme = full.center()
                if edge == "left":
                    extreme.setX(full.left() - 1)
                elif edge == "top":
                    extreme.setY(full.top() - 1)
                else:
                    extreme.setX(full.right() + 1)
                self.pointer.return_value = extreme
                controller.check_pointer()
                for _ in range(3):
                    QApplication.sendEvent(window, QEvent(QEvent.Leave))
                    window._defer_fit()
                    self.app.processEvents()
                    controller.check_pointer()
                    self.assertFalse(controller.collapsed)
                    self.assertEqual(window.geometry(), full)
                    self.assertTrue(window.timer.isActive())
                self.pointer.return_value = full.center()
                controller.check_pointer()
                self.assertFalse(controller.collapsed)
                self.pointer.return_value = QPoint(99999, 99999)
                controller.check_pointer()
                self.assertTrue(controller.collapsed)
                self.assertFalse(window.timer.isActive())

    def test_visible_strip_shares_drag_cancel_double_click_and_restore(self):
        def mouse(target, kind, point, button, buttons):
            event = QMouseEvent(kind, QPointF(target.mapFromGlobal(point)), QPointF(point),
                                button, buttons, Qt.NoModifier)
            QApplication.sendEvent(target, event)

        for edge in ("left", "top", "right"):
            with self.subTest(edge=edge):
                window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
                window.move(self.edge_position(window, edge))
                controller = window.position_controller
                full, strip = controller.full_geometry(), window.geometry()
                start = strip.center()
                target = window.childAt(window.rect().center())
                self.assertIs(target, window.panel)
                mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                mouse(target, QEvent.MouseMove, start + QPoint(1, 0), Qt.NoButton, Qt.LeftButton)
                mouse(target, QEvent.MouseButtonRelease, start + QPoint(1, 0), Qt.LeftButton, Qt.NoButton)
                self.assertEqual(window.geometry(), strip)
                delta = QPoint(-40 if edge == "right" else 40, 25)
                mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
                mouse(target, QEvent.MouseMove, start + delta, Qt.NoButton, Qt.LeftButton)
                self.assertEqual(window.pos(), full.topLeft() + delta)
                QApplication.sendEvent(window, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
                mouse(target, QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
                self.assertEqual(window.geometry(), strip)
                mouse(target, QEvent.MouseButtonDblClick, start, Qt.LeftButton, Qt.LeftButton)
                self.assertFalse(window.widget_visible)
                window.toggle_win()
                self.app.processEvents()
                self.assertEqual(window.geometry(), strip)
                self.assertEqual(controller.full_geometry(), full)

    def test_shrunk_geometry_survives_deferred_fit_topmost_hide_and_restore(self):
        for mode in ("float", "both"):
            with self.subTest(mode=mode):
                window = self.window(boundary_check_enabled=True, edge_hide_enabled=True,
                                     taskbar_enabled=True, display_mode=mode)
                window.move(self.edge_position(window, "right"))
                full = window.position_controller.full_geometry()
                window._defer_fit()
                self.app.processEvents()
                self.assertEqual(window.width(), 5)
                self.assertEqual(window.position_controller.full_geometry(), full)
                window.set_float_on_top(False)
                self.app.processEvents()
                self.assertEqual(window.width(), 5)
                window.hide_widget()
                self.assertFalse(window.widget_visible)
                window.toggle_win()
                self.app.processEvents()
                self.assertEqual(window.width(), 5)
                self.assertEqual(window.display_mode, mode)
                self.assertFalse(window.timer.isActive())

    def test_dragging_strip_expands_without_jumping_and_cancel_restores_strip(self):
        window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
        window.move(self.edge_position(window, "right"))
        controller = window.position_controller
        full = controller.full_geometry()
        strip = window.geometry()
        start = strip.center()
        window.begin_drag(start)
        self.assertEqual(window.geometry(), full)
        window.move_drag(start - QPoint(60, 0))
        self.assertEqual(window.pos(), full.topLeft() - QPoint(60, 0))
        self.assertFalse(controller.collapsed)
        window.finish_drag(False)
        self.assertTrue(controller.collapsed)
        self.assertEqual(window.geometry(), strip)
        window.begin_drag(start)
        window.move_drag(start - QPoint(60, 0))
        window.finish_drag()
        self.assertFalse(controller.collapsed)
        self.assertEqual(window.pos(), full.topLeft() - QPoint(60, 0))

    def test_collapsed_blocks_direct_fetch_and_late_data_and_resumes_on_enter(self):
        window = self.window(boundary_check_enabled=True, edge_hide_enabled=True)
        window.move(self.edge_position(window, "left"))
        with patch("stockwidget.ui.widget.threading.Thread") as thread:
            window.watchlist = {"sh600519": {"checked": True}}
            # 调用未被 setUp 替换的原入口，验证直接刷新也不会绕过暂停。
            ORIGINAL_REFRESH(window)
            thread.assert_not_called()
        before = window.message_label.text()
        window._process_data((False, None, "不应显示的迟到错误"))
        self.assertEqual(window.message_label.text(), before)
        self.pointer.return_value = window.geometry().center()
        event = QEnterEvent(QPointF(2, 2), QPointF(2, 2), QPointF(self.pointer.return_value))
        QApplication.sendEvent(window, event)
        self.assertFalse(window.position_controller.collapsed)
        self.assertTrue(window.timer.isActive())

    def test_function_controls_defaults_dependency_reset_and_theme(self):
        window = self.window()
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(window, window)
        self.windows.append(dialog)
        dialog.show()
        dialog.ui.tab_widget.setCurrentWidget(dialog.ui.functions)
        self.app.processEvents()
        self.assertFalse(dialog.ui.cb_boundary_check.isChecked())
        self.assertFalse(dialog.ui.cb_edge_hide.isEnabled())
        QTest.mouseClick(dialog.ui.cb_boundary_check, Qt.LeftButton)
        self.assertTrue(window.boundary_check_enabled)
        self.assertTrue(dialog.ui.cb_edge_hide.isEnabled())
        QTest.mouseClick(dialog.ui.cb_edge_hide, Qt.LeftButton)
        self.assertTrue(window.edge_hide_enabled)
        dialog._apply_theme_stylesheet()
        self.assertTrue(dialog.ui.cb_edge_hide.isEnabled())
        QTest.mouseClick(dialog.ui.cb_boundary_check, Qt.LeftButton)
        self.assertFalse(dialog.ui.cb_edge_hide.isChecked())
        self.assertFalse(dialog.ui.cb_edge_hide.isEnabled())
        window.set_position_options(boundary_check_enabled=True, edge_hide_enabled=True)
        self.assertTrue(dialog.ui.cb_edge_hide.isChecked())
        window.reset_settings()
        self.assertFalse(dialog.ui.cb_boundary_check.isChecked())
        self.assertFalse(dialog.ui.cb_edge_hide.isChecked())
        self.assertFalse(dialog.ui.cb_edge_hide.isEnabled())
        for control in (dialog.ui.cb_boundary_check, dialog.ui.cb_edge_hide):
            self.assertTrue(dialog.ui.functions.isAncestorOf(control))
            self.assertTrue(dialog.ui.gb_fcn.rect().contains(control.geometry()))


ORIGINAL_REFRESH = FloatLabel._refresh_from_function
