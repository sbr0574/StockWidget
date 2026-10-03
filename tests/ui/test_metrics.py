"""固定指标目录、可排序显示池及共用行情列投影。"""

from unittest.mock import Mock, patch

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtTest import QTest

from stockwidget.core.quote_presentation import (
    METRIC_SPECS,
    visible_metrics_from_config,
    METRIC_IDS,
    metric_headers,
)
from stockwidget.ui.controls.metrics import MetricListWidget, MetricPoolWidget
from stockwidget.ui.controls.quote_view import COLOR_ROLE_TEXT
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.widget import FloatLabel

from tests.support import QtTestCase, SettingsTestCase


class MetricPoolWidgetTests(QtTestCase):

    def setUp(self):
        self.pool = MetricPoolWidget()

    def tearDown(self):
        self.pool.close()
        self.pool.deleteLater()
        self.qt_app.processEvents()

    def test_move_between_pools_and_reorder_visible_metrics(self):
        self.pool.set_visible_metrics(["price", "kline", "amount"])
        changes = []
        self.pool.visible_metrics_changed.connect(changes.append)

        self.pool.move_metric("displayed", "displayed", "amount", 0)
        self.pool.move_metric("displayed", "available", "price", 0)
        self.pool.move_metric("available", "displayed", "b1s1", 1)

        self.assertEqual(
            self.pool.visible_metrics,
            ["amount", "b1s1", "kline"],
        )
        self.assertEqual(changes[-1], ["amount", "b1s1", "kline"])
        self.assertEqual(self.pool.available_pool.count(), len(METRIC_IDS))
        for row in range(self.pool.available_pool.count()):
            item = self.pool.available_pool.item(row)
            self.assertEqual(bool(item.flags() & Qt.ItemIsDragEnabled),
                             item.data(Qt.UserRole) not in self.pool.visible_metrics)

    def test_catalog_keeps_all_entries_in_place_and_disables_already_displayed_entries(self):
        self.pool.resize(480, 300)
        self.pool.set_visible_metrics(["price"])
        self.pool.show()
        self.qt_app.processEvents()
        catalog = self.pool.available_pool
        identifiers = [catalog.item(i).data(Qt.UserRole) for i in range(catalog.count())]
        price = next(catalog.item(i) for i in range(catalog.count()) if catalog.item(i).data(Qt.UserRole) == "price")
        name = next(catalog.item(i) for i in range(catalog.count()) if catalog.item(i).data(Qt.UserRole) == "name")
        position, height = catalog.visualItemRect(name), catalog.height()
        QTest.mouseClick(catalog.viewport(), Qt.LeftButton, pos=position.center())
        QTest.mouseDClick(catalog.viewport(), Qt.LeftButton, pos=position.center())
        self.assertEqual(self.pool.visible_metrics, ["price", "name"])
        self.assertFalse(name.flags() & Qt.ItemIsEnabled)
        self.assertFalse(price.flags() & Qt.ItemIsDragEnabled)
        catalog.setCurrentItem(price)
        with patch("stockwidget.ui.controls.metrics.QDrag") as drag:
            catalog.startDrag(Qt.MoveAction)
        drag.assert_not_called()
        self.pool.move_metric("available", "displayed", "price", 0)
        self.assertEqual(self.pool.visible_metrics, ["price", "name"])
        self.pool.move_metric("displayed", "available", "name", 0)
        self.assertTrue(name.flags() & Qt.ItemIsEnabled)
        self.assertEqual([catalog.item(i).data(Qt.UserRole) for i in range(catalog.count())], identifiers)
        self.assertEqual(catalog.visualItemRect(name), position)
        self.assertEqual(catalog.height(), height)
        displayed_height = self.pool.displayed_pool.height()
        self.pool.set_visible_metrics(list(METRIC_IDS))
        self.qt_app.processEvents()
        self.assertEqual((catalog.height(), self.pool.displayed_pool.height()), (56, displayed_height))
        self.assertTrue(all(self.pool.displayed_pool.viewport().rect().contains(
            self.pool.displayed_pool.visualItemRect(self.pool.displayed_pool.item(i)))
            for i in range(self.pool.displayed_pool.count())))

    def test_disabled_catalog_text_remains_readable_on_a_dark_palette(self):
        palette = QPalette(self.pool.palette())
        palette.setColor(QPalette.Window, QColor("#222222"))
        palette.setColor(QPalette.Base, QColor("#222222"))
        palette.setColor(QPalette.Text, QColor("#eeeeee"))
        palette.setColor(QPalette.Mid, QColor("#222222"))
        self.pool.setPalette(palette)
        self.pool.resize(480, 260)
        self.pool.set_visible_metrics(["kline"])
        self.pool.set_theme(True)
        self.pool.show()
        self.qt_app.processEvents()
        catalog = self.pool.available_pool
        item = next(catalog.item(i) for i in range(catalog.count())
                    if catalog.item(i).data(Qt.UserRole) == "kline")
        image = catalog.viewport().grab(catalog.visualItemRect(item)).toImage()
        self.assertGreater(max(image.pixelColor(x, y).lightness()
                               for x in range(image.width()) for y in range(image.height())), 120)

    def test_available_pool_does_not_have_its_own_persisted_order(self):
        self.pool.set_visible_metrics(["price"])
        changes = []
        self.pool.visible_metrics_changed.connect(changes.append)

        self.pool.move_metric("available", "available", "kline", 0)

        self.assertEqual(self.pool.visible_metrics, ["price"])
        self.assertEqual(changes, [])

    def test_drop_insert_index_handles_blank_gap_in_wrapped_rows(self):
        pool = MetricListWidget("displayed", "空")
        pool.set_metric_ids(["price", "kline", "amount", "b1s1"])
        # 模拟两行布局：第一行两个块，第二行两个块
        rects = {
            0: QRect(0, 0, 50, 22),
            1: QRect(52, 0, 50, 22),
            2: QRect(0, 24, 60, 22),
            3: QRect(62, 24, 70, 22),
        }
        pool.visualRect = lambda index: rects[index.row()]

        # 第一行右侧空白处：应插在第一行之后（索引 2），而不是队尾
        self.assertEqual(pool._drop_insert_index(QPoint(120, 10)), 2)
        # 第一个块左侧
        self.assertEqual(pool._drop_insert_index(QPoint(-10, 5)), 0)
        # 最后一个块右侧
        self.assertEqual(pool._drop_insert_index(QPoint(200, 30)), 4)
        # 第二行两个块之间
        self.assertEqual(pool._drop_insert_index(QPoint(61, 30)), 3)

        pool.close()
        pool.deleteLater()


class FloatLabelMetricLayoutTests(QtTestCase):

    def setUp(self):
        self._refresh_patcher = patch.object(QuotePresenter, "refresh")
        self.refresh = self._refresh_patcher.start()
        self.windows = []

    def tearDown(self):
        for window in self.windows:
            window.close()
            window.deleteLater()
        self.qt_app.processEvents()
        self._refresh_patcher.stop()

    def _window(self, config):
        window = FloatLabel(config, {})
        self.windows.append(window)
        self.refresh.reset_mock()
        return window

    @staticmethod
    def _row_and_roles():
        headers = metric_headers([spec.metric_id for spec in METRIC_SPECS])
        row = {header: header for header in headers}
        roles = {header: COLOR_ROLE_TEXT for header in headers}
        return row, roles

    def test_name_metric_migrates_first_and_projects_in_order(self):
        window = self._window(
            {
                "name_visible": True,
                "visible_metrics": ["kline", "price", "b1s1"],
            }
        )
        row, roles = self._row_and_roles()

        window.quotes.project_rows([row], [roles])

        self.assertEqual(
            window.model._headers,
            ["名称", "K线", "现价", "买一/卖一"],
        )
        self.assertEqual(
            window.visible_metrics,
            ["name", "kline", "price", "b1s1"],
        )

    def test_startup_is_compact_and_keeps_saved_position_near_right_edge(self):
        screen = self.qt_app.primaryScreen().availableGeometry()
        saved = QPoint(screen.right() - 25, screen.top() + 100)
        window = self._window({
            "watchlist": {"sh600519": {"checked": True}},
            "pos": {"x": saved.x(), "y": saved.y()},
            "visible_metrics": ["price"],
            "name_visible": False,
        })
        window.show()
        self.qt_app.processEvents()
        self.assertLess(window.width(), 150)
        self.assertEqual(window.pos(), saved)
        self.assertTrue(window.table.isHidden())
        row, roles = self._row_and_roles()
        with patch("stockwidget.ui.floating.presenter.format_quote", return_value=(row, roles, {})):
            window.quotes.accept_result((True, {"sh600519": {}}, None))
        self.qt_app.processEvents()
        self.assertEqual(window.pos(), saved)
        self.assertFalse(window.table.isHidden())
        self.assertTrue(window.message_label.isHidden())
        self.assertEqual(window.model._headers, ["现价"])

    def test_name_metric_can_be_reordered_like_other_metrics(self):
        window = self._window(
            {
                "name_visible": True,
                "visible_metrics": ["kline", "price", "name"],
            }
        )
        row, roles = self._row_and_roles()

        window.quotes.project_rows([row], [roles])

        self.assertEqual(
            window.model._headers,
            ["K线", "现价", "名称"],
        )
        self.assertEqual(
            window.model._headers,
            metric_headers(window.visible_metrics),
        )

    def test_setting_order_serializes_legacy_flags_and_saves_once(self):
        window = self._window({})
        save = Mock()
        window.set_on_change(save)
        changes = []
        window.display_flags_changed.connect(lambda: changes.append(True))

        window.set_visible_metrics(["amount", "price"])

        self.assertEqual(window.visible_metrics, ["amount", "price"])
        self.assertIn("amount", window.visible_metrics)
        self.assertIn("price", window.visible_metrics)
        self.assertNotIn("change_pct", window.visible_metrics)
        config = window.current_config()
        self.assertEqual(config["visible_metrics"], ["amount", "price"])
        self.assertTrue(config["amount_visible"])
        self.assertTrue(config["price_visible"])
        self.assertFalse(config["name_visible"])
        self.assertEqual(visible_metrics_from_config(config), window.visible_metrics)
        save.assert_called_once_with()
        self.refresh.assert_called_once_with()
        self.assertEqual(changes, [True])

    def test_reenabling_metric_appends_it_and_bid_ask_toggle_together(self):
        window = self._window({"visible_metrics": ["price"]})

        window.set_metric_visible("amount", True)
        window.set_metric_visible("b1s1", True)
        window.set_metric_visible("b1s1", False)

        self.assertEqual(window.visible_metrics, ["name", "price", "amount"])
        self.assertNotIn("b1s1", window.visible_metrics)

    def test_kline_delegate_moves_to_new_column(self):
        window = self._window(
            {"name_visible": False, "visible_metrics": ["kline", "price"]}
        )
        row, roles = self._row_and_roles()
        window.quotes.project_rows([row], [roles])
        self.assertIs(window.table.itemDelegateForColumn(0), window.k_delegate)

        window.set_visible_metrics(["price", "kline"])
        window.quotes.project_rows([row], [roles])

        self.assertIs(
            window.table.itemDelegateForColumn(0),
            window._default_item_delegate,
        )
        self.assertIs(window.table.itemDelegateForColumn(1), window.k_delegate)

    def test_empty_layout_prompts_until_name_is_enabled(self):
        window = self._window(
            {"name_visible": False, "visible_metrics": []}
        )

        window.quotes.project_rows([], [])
        self.assertFalse(window.message_label.isHidden())
        self.assertIn("至少一个显示指标", window.message_label.text())

        window.set_metric_visible("name", True)
        window.quotes.project_rows([], [])

        self.assertEqual(window.model._headers, ["名称"])
        self.assertTrue(window.message_label.isHidden())

    def test_sort_forces_name_temporarily_and_restores_watchlist_order(self):
        window = self._window(
            {"name_visible": False, "visible_metrics": ["price", "change_pct"]}
        )
        row1, roles1 = self._row_and_roles()
        row2, roles2 = self._row_and_roles()
        row1["名称"], row1["现价"], row1["涨幅"] = "A", "10.00", "+1.00%"
        row2["名称"], row2["现价"], row2["涨幅"] = "B", "20.00", "+2.00%"
        window.quotes._last_full_rows = [row1, row2]
        window.quotes._last_color_roles = [roles1, roles2]
        window.quotes._last_sort_values = [
            {"现价": 10.0, "涨幅": 1.0},
            {"现价": 20.0, "涨幅": 2.0},
        ]

        window.quotes.set_sort("现价", Qt.SortOrder.DescendingOrder)

        self.assertEqual(window.model._headers, ["名称", "现价", "涨幅"])
        self.assertEqual(window.model._rows[0][0], "B")
        self.assertNotIn("name", window.visible_metrics)

        window.quotes.clear_sort()

        self.assertEqual(window.model._headers, ["现价", "涨幅"])
        self.assertEqual(window.model._rows[0][0], "10.00")

    def test_missing_sort_values_stay_at_bottom_in_both_directions(self):
        window = self._window(
            {"name_visible": True, "visible_metrics": ["name", "profit"]}
        )
        base_row, base_roles = self._row_and_roles()
        rows = []
        roles = []
        for name, profit in (("A", "+3.00%"), ("B", "-"), ("C", "-1.00%")):
            row = dict(base_row)
            row["名称"] = name
            row["浮盈"] = profit
            rows.append(row)
            roles.append(dict(base_roles))
        window.quotes._last_full_rows = rows
        window.quotes._last_color_roles = roles
        window.quotes._last_sort_values = [
            {"浮盈": 3.0}, {"浮盈": None}, {"浮盈": -1.0},
        ]

        window.quotes.set_sort("浮盈", Qt.SortOrder.AscendingOrder)
        self.assertEqual([r[0] for r in window.model._rows], ["C", "A", "B"])

        window.quotes.set_sort("浮盈", Qt.SortOrder.DescendingOrder)
        self.assertEqual([r[0] for r in window.model._rows], ["A", "C", "B"])

    def test_header_click_only_sorts_supported_metrics(self):
        window = self._window(
            {"name_visible": True, "visible_metrics": ["name", "price", "kline"]}
        )
        row, roles = self._row_and_roles()
        window.quotes._last_full_rows = [row]
        window.quotes._last_color_roles = [roles]
        window.quotes._last_sort_values = [{"现价": 1.0}]
        window.quotes.project_rows([row], [roles], [{"现价": 1.0}])

        window.quotes.header_clicked(window.model._headers.index("K线"))
        self.assertIsNone(window.quotes.sort_header)

        window.quotes.header_clicked(window.model._headers.index("现价"))
        self.assertEqual(window.quotes.sort_header, "现价")
        self.assertEqual(window.quotes.sort_order, Qt.SortOrder.DescendingOrder)
        window.quotes.header_clicked(window.model._headers.index("现价"))
        self.assertEqual(window.quotes.sort_order, Qt.SortOrder.AscendingOrder)


class MetricSettingsTests(SettingsTestCase):
    def test_buttons_and_metric_pools_follow_palette_changes(self):
        dialog, _window = self._make_dialog()
        original = self.qt_app.palette()
        try:
            palette = QPalette(original)
            palette.setColor(QPalette.Accent, QColor("#a23b61"))
            self.qt_app.setPalette(palette)
            self.qt_app.processEvents()
            self.assertIn("#a23b61", dialog.styleSheet())
            self.assertIn("rgba(162, 59, 97, 0.08)", dialog.styleSheet())
            for pool in (dialog.metric_pool.displayed_pool, dialog.metric_pool.available_pool):
                self.assertEqual(pool._drop_color.name(), "#a23b61")
                self.assertIn("rgba(162, 59, 97, 0.08)", pool.styleSheet())
        finally:
            self.qt_app.setPalette(original)

    def test_taskbar_metric_pool_reuses_selection_order_and_preserves_independent_metrics(self):
        from stockwidget.core.quote_presentation import METRIC_SPECS
        from stockwidget.ui.controls.metrics import MetricPoolWidget
        dialog, window = self._make_dialog()
        taskbar = dialog.taskbar_settings
        self.assertFalse(taskbar.enabled.isChecked())
        self.assertFalse(taskbar.body.isEnabled())
        self.assertFalse(taskbar.metrics.isEnabled())
        taskbar.enabled.setChecked(True)
        self.assertTrue(window.view_options.taskbar_enabled)
        self.assertTrue(taskbar.metrics.sync_toggle.isChecked())
        self.assertFalse(taskbar.metric_pool.isEnabled())
        self.assertIsInstance(taskbar.metric_pool, MetricPoolWidget)
        window.set_visible_metrics(["volume", "name", "price"])
        self.assertEqual(taskbar.metric_pool.visible_metrics, ["volume", "name", "price"])
        taskbar.metrics.sync_toggle.setChecked(False)
        self.assertTrue(taskbar.metric_pool.isEnabled())
        pool = taskbar.metric_pool
        for spec in METRIC_SPECS:
            if spec.metric_id not in pool.visible_metrics:
                pool.move_metric("available", "displayed", spec.metric_id, len(pool.visible_metrics))
        self.assertEqual(len(window.view_options.taskbar_metrics), len(METRIC_SPECS))
        pool.move_metric("displayed", "displayed", "kline", 0)
        independent = pool.visible_metrics
        self.assertEqual(independent[0], "kline")
        self.assertEqual(window.view_options.taskbar_metrics, independent)
        self.assertEqual(window.visible_metrics, ["volume", "name", "price"])
        taskbar.metrics.sync_toggle.setChecked(True)
        self.assertEqual(pool.visible_metrics, window.visible_metrics)
        window.set_visible_metrics(["price", "change"])
        self.assertEqual(pool.visible_metrics, ["price", "change"])
        taskbar.metrics.sync_toggle.setChecked(False)
        self.assertEqual(pool.visible_metrics, independent)
        taskbar.enabled.setChecked(False)
        self.assertFalse(taskbar.body.isEnabled())
        self.assertFalse(pool.isEnabled())
        self.assertEqual(window.display_mode, "float")

    def test_metric_pool_controls_order_and_name_visibility(self):
        dialog, window = self._make_dialog()
        save = Mock()
        window.set_on_change(save)

        self.assertEqual(
            dialog.metric_pool.visible_metrics,
            ["name", "price", "change_pct"],
        )
        self.assertIn("name", window.visible_metrics)
        displayed_ids = [
            dialog.metric_pool.displayed_pool.item(row).data(
                Qt.ItemDataRole.UserRole
            )
            for row in range(dialog.metric_pool.displayed_pool.count())
        ]
        self.assertEqual(displayed_ids, ["name", "price", "change_pct"])

        with patch.object(window.quotes, "refresh") as refresh:
            dialog.metric_pool.move_metric(
                "available", "displayed", "kline", 0
            )

        self.assertEqual(
            window.visible_metrics,
            ["kline", "name", "price", "change_pct"],
        )
        self.assertEqual(
            dialog.metric_pool.visible_metrics,
            window.visible_metrics,
        )
        save.assert_called_once_with()
        refresh.assert_called_once_with()

        # 名称与其他指标一样可从显示池移出，从而关闭名称列
        with patch.object(window.quotes, "refresh") as refresh:
            dialog.metric_pool.move_metric(
                "displayed", "available", "name", 0
            )
        self.assertNotIn("name", window.visible_metrics)
        self.assertEqual(
            window.visible_metrics,
            ["kline", "price", "change_pct"],
        )



    def test_outside_press_clears_watchlist_and_metric_selections(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        displayed = dialog.metric_pool.displayed_pool

        dialog.watchlist_editor.list_codes.setCurrentCell(1, 1)
        displayed.setCurrentItem(displayed.item(1))
        self.qt_app.processEvents()
        self.assertEqual(dialog.watchlist_editor.list_codes.currentRow(), 1)
        self.assertTrue(dialog.watchlist_editor.btn_del.isEnabled())
        self.assertIsNotNone(displayed.currentItem())

        # 点击设置页其他区域（空白处）→ 自选列表与指标池选中全部清除
        dialog._handle_outside_press(dialog.ui.gb_color)
        self.qt_app.processEvents()

        self.assertEqual(dialog.watchlist_editor.list_codes.currentRow(), -1)
        self.assertFalse(dialog.watchlist_editor.btn_del.isEnabled())
        self.assertFalse(dialog.watchlist_editor.btn_top.isEnabled())
        self.assertIsNone(displayed.currentItem())

    def test_press_inside_pool_clears_watchlist_and_other_pool(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        displayed = dialog.metric_pool.displayed_pool
        available = dialog.metric_pool.available_pool

        dialog.watchlist_editor.list_codes.setCurrentCell(0, 1)
        displayed.setCurrentItem(displayed.item(1))
        self.qt_app.processEvents()

        dialog._handle_outside_press(available.viewport())
        self.qt_app.processEvents()

        self.assertEqual(dialog.watchlist_editor.list_codes.currentRow(), -1)
        self.assertIsNone(displayed.currentItem())
        self.assertFalse(dialog.watchlist_editor.btn_del.isEnabled())

    def test_press_on_delete_or_top_button_keeps_selection(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        dialog.watchlist_editor.list_codes.setCurrentCell(1, 1)
        self.qt_app.processEvents()

        dialog._handle_outside_press(dialog.watchlist_editor.btn_del)
        dialog._handle_outside_press(dialog.watchlist_editor.btn_top)
        self.qt_app.processEvents()

        self.assertEqual(dialog.watchlist_editor.list_codes.currentRow(), 1)
        self.assertTrue(dialog.watchlist_editor.btn_del.isEnabled())

    def test_plain_metric_chip_does_not_keep_selection(self):
        dialog, _window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        available = dialog.metric_pool.available_pool

        # 单击没有独立面板的指标块：不保留选中状态
        displayed.itemClicked.emit(displayed.item(1))
        self.qt_app.processEvents()
        self.assertIsNone(displayed.currentItem())

        # 双击/拖动切换显示后，普通指标同样不保留选中
        dialog.metric_pool.move_metric("available", "displayed", "kline", 0)
        self.qt_app.processEvents()
        self.assertIsNone(displayed.currentItem())
        self.assertIsNone(available.currentItem())

    def test_double_click_moves_name_metric_to_the_other_pool(self):
        dialog, window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        name_item = next(
            displayed.item(row)
            for row in range(displayed.count())
            if displayed.item(row).data(Qt.ItemDataRole.UserRole) == "name"
        )

        # 单击结束后清除选中；所有指标均采用相同的双击交互。
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        dialog.show()
        self.qt_app.processEvents()
        pos = displayed.visualItemRect(name_item).center() - QPoint(8, 0)
        QTest.mouseClick(displayed.viewport(), Qt.LeftButton, pos=pos)
        self.assertIsNone(displayed.currentItem())

        # 双击文字快速移动到另一池。
        with patch.object(window.quotes, "refresh"):
            QTest.mouseDClick(displayed.viewport(), Qt.LeftButton, pos=pos)
        self.qt_app.processEvents()

        self.assertNotIn("name", window.visible_metrics)
        self.assertEqual(displayed.currentItem(), None)
