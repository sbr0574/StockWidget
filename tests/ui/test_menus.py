"""Menu actions follow the visible surface and retain shared state ownership."""

from unittest.mock import Mock, patch

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from shiboken6 import delete

from stockwidget.ui.menus import TrayIcon
from tests.support import PagingTestCase


def actions(menu):
    return {action.text(): action for action in menu.actions() if not action.isSeparator()}


class MenuBehaviorTests(PagingTestCase):
    def setUp(self):
        super().setUp()
        self.menus = []

    def tearDown(self):
        for menu in self.menus:
            delete(menu)
        super().tearDown()

    def menu(self, surface="float"):
        menu = self.win.build_context_menu(surface)
        self.menus.append(menu)
        return menu

    def test_sort_menu_follows_visible_column_order_and_omits_unavailable_fields(self):
        self.win.set_visible_metrics(["change_pct", "kline", "price"])
        sort = actions(self.menu())["排序"].menu()
        self.assertEqual(list(actions(sort)), ["涨幅", "现价", "恢复自选顺序"])
        self.assertFalse(actions(sort)["恢复自选顺序"].isEnabled())

    def test_independent_taskbar_sort_changes_shared_rows_and_can_be_cleared_from_float(self):
        self.win.set_visible_metrics(["name"])
        self.win.set_view_options(taskbar_sync_metrics=False, taskbar_metrics=["price"])
        sort = actions(self.menu("taskbar"))["排序"].menu()
        actions(actions(sort)["现价"].menu())["降序"].trigger()
        self.assertEqual(self.win.quotes.sort_header, "现价")
        self.assertEqual(self.win.quotes._ordered_rows[0]["现价"], "50")
        self.assertEqual(self.win.model._rows[0][0], "标的5")
        self.assertEqual(self.win.taskbar_model._rows[0][0], "50")

        float_sort = actions(actions(self.menu())["排序"].menu())
        current = float_sort["现价（当前排序，未显示）"].menu()
        self.assertTrue(actions(current)["降序"].isChecked())
        actions(current)["升序"].trigger()
        self.assertEqual(self.win.quotes._ordered_rows[0]["现价"], "10")
        float_sort["恢复自选顺序"].trigger()
        self.assertIsNone(self.win.quotes.sort_header)
        self.assertEqual(self.win.model._rows[0][0], "标的1")

    def test_chart_only_surface_disables_sort_until_a_shared_sort_needs_clearing(self):
        self.win.set_visible_metrics(["kline"])
        self.assertFalse(actions(self.menu())["排序"].isEnabled())
        self.win.quotes.set_sort("现价", Qt.DescendingOrder)
        sort = actions(self.menu())["排序"]
        self.assertTrue(sort.isEnabled())
        actions(sort.menu())["恢复自选顺序"].trigger()
        self.assertIsNone(self.win.quotes.sort_header)
        self.assertFalse(actions(self.menu())["排序"].isEnabled())

    def test_removing_sort_metric_from_one_surface_preserves_it_until_neither_uses_it(self):
        self.win.set_visible_metrics(["name", "price"])
        self.win.set_view_options(taskbar_metrics=["price"])
        self.win.quotes.set_sort("现价", Qt.DescendingOrder)
        self.win.set_visible_metrics(["name"])
        self.assertEqual(self.win.quotes.sort_header, "现价")
        self.assertEqual(self.win.model._rows[0][0], "标的5")
        self.win.quotes.float_page = self.win.quotes.taskbar_page = 1
        self.win.set_view_options(taskbar_metrics=["name"])
        self.assertIsNone(self.win.quotes.sort_header)
        self.assertEqual((self.win.quotes.float_page, self.win.quotes.taskbar_page), (0, 0))
        self.assertEqual(self.win.model._rows[0][0], "标的1")

    def test_inactive_independent_metrics_cannot_keep_or_start_a_sort(self):
        for changes in ({"taskbar_enabled": False}, {"taskbar_sync_metrics": True}):
            with self.subTest(changes=changes):
                self.win.set_visible_metrics(["name"])
                self.win.set_view_options(taskbar_enabled=True, taskbar_sync_metrics=False,
                                          taskbar_metrics=["price"])
                self.win.quotes.set_sort("现价", Qt.DescendingOrder)
                self.assertEqual(self.win.quotes.sort_header, "现价")
                self.win.set_view_options(**changes)
                self.assertIsNone(self.win.quotes.sort_header)
                self.win.quotes.set_sort("现价", Qt.DescendingOrder)
                self.assertIsNone(self.win.quotes.sort_header)
                self.assertEqual(self.win.view_options.taskbar_metrics, ["price"])

    def test_taskbar_actions_respect_sync_and_preserve_independent_preferences(self):
        self.win.set_visible_metrics(["name"])
        self.win.set_unicolor(False)
        self.win.set_view_options(taskbar_sync_metrics=True, taskbar_metrics=["volume"],
                                  taskbar_sync_split=True, float_split_enabled=False,
                                  taskbar_split_enabled=False, taskbar_sync_appearance=True,
                                  taskbar_unicolor=False)
        menu = actions(self.menu("taskbar"))
        actions(menu["显示指标"].menu())["现价"].trigger()
        menu["分栏"].trigger()
        menu["统一颜色"].trigger()
        self.assertEqual(self.win.visible_metrics, ["name", "price"])
        self.assertTrue(self.win.view_options.float_split_enabled)
        self.assertTrue(self.win.unicolor)
        self.assertEqual(self.win.view_options.taskbar_metrics, ["volume"])
        self.assertFalse(self.win.view_options.taskbar_split_enabled)
        self.assertFalse(self.win.view_options.taskbar_unicolor)

        self.win.set_view_options(taskbar_sync_metrics=False, taskbar_sync_split=False,
                                  taskbar_sync_appearance=False)
        menu = actions(self.menu("taskbar"))
        actions(menu["显示指标"].menu())["现价"].trigger()
        menu["分栏"].trigger()
        menu["统一颜色"].trigger()
        self.assertEqual(self.win.view_options.taskbar_metrics, ["volume", "price"])
        self.assertTrue(self.win.view_options.taskbar_split_enabled)
        self.assertTrue(self.win.view_options.taskbar_unicolor)
        self.assertEqual(self.win.visible_metrics, ["name", "price"])
        self.assertTrue(self.win.view_options.float_split_enabled)
        self.assertTrue(self.win.unicolor)

    def test_taskbar_auto_color_updates_saved_preference_and_respects_appearance_sync(self):
        self.win.set_view_options(taskbar_sync_appearance=False, taskbar_auto_color=False)
        self.assertNotIn("自动配色", actions(self.menu("float")))
        action = actions(self.menu("taskbar"))["自动配色"]
        self.assertTrue(action.isEnabled())
        self.assertFalse(action.isChecked())
        action.trigger()
        self.assertTrue(self.win.current_config()["taskbar_auto_color"])
        self.win.set_view_options(taskbar_sync_appearance=True)
        action = actions(self.menu("taskbar"))["自动配色"]
        self.assertFalse(action.isEnabled())
        self.assertTrue(action.isChecked())
        self.win.set_view_options(taskbar_sync_appearance=False)
        actions(self.menu("taskbar"))["自动配色"].trigger()
        self.assertFalse(self.win.current_config()["taskbar_auto_color"])

    def test_tray_visibility_label_and_split_follow_current_surface_each_time_it_opens(self):
        self.win.set_view_options(taskbar_sync_split=False, float_split_enabled=False,
                                  taskbar_split_enabled=True)
        with patch("sys.platform", "win32"):
            tray = TrayIcon(QIcon(), "Test", on_toggle=self.win.toggle_win,
                            on_open_settings=Mock(), on_quit=Mock(), source=self.win)
            try:
                self.win.set_widget_visible(True)
                tray.contextMenu().aboutToShow.emit()
                actions(tray.contextMenu())["隐藏行情"].trigger()
                self.assertFalse(self.win.widget_visible)
                tray.contextMenu().aboutToShow.emit()
                actions(tray.contextMenu())["显示行情"].trigger()
                self.assertTrue(self.win.widget_visible)
                for mode, expected in (("float", False), ("taskbar", True)):
                    self.win.set_display_mode(mode)
                    tray.contextMenu().aboutToShow.emit()
                    split = actions(tray.contextMenu())["分栏"]
                    self.assertEqual(split.isChecked(), expected)
                    split.trigger()
                    option = f"{mode}_split_enabled"
                    self.assertEqual(getattr(self.win.view_options, option), not expected)
            finally:
                delete(tray.contextMenu())
                delete(tray)


if __name__ == "__main__":
    import unittest
    unittest.main()
