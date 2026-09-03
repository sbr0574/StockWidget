import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from stockwidget.core.metric_layout import METRIC_IDS
from stockwidget.ui.metric_pool import MetricPoolWidget
from stockwidget.ui.table_model import COLOR_ROLE_TEXT
from stockwidget.ui.widget import FloatLabel


class MetricPoolWidgetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])

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
        self.assertEqual(
            self.pool.displayed_pool.count() + self.pool.available_pool.count(),
            len(METRIC_IDS),
        )

    def test_available_pool_does_not_have_its_own_persisted_order(self):
        self.pool.set_visible_metrics(["price"])
        changes = []
        self.pool.visible_metrics_changed.connect(changes.append)

        self.pool.move_metric("available", "available", "kline", 0)

        self.assertEqual(self.pool.visible_metrics, ["price"])
        self.assertEqual(changes, [])


class FloatLabelMetricLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])

    def setUp(self):
        self._refresh_patcher = patch.object(FloatLabel, "_refresh_from_function")
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
        row = {header: header for header in FloatLabel.ALL_HEADERS}
        roles = {header: COLOR_ROLE_TEXT for header in FloatLabel.ALL_HEADERS}
        return row, roles

    def test_name_is_independent_and_always_projects_first(self):
        window = self._window(
            {
                "name_visible": True,
                "visible_metrics": ["kline", "price", "b1s1"],
            }
        )
        row, roles = self._row_and_roles()

        window._project_columns([row], [roles])

        self.assertEqual(
            window.model._headers,
            ["名称", "K线", "现价", "买一", "卖一"],
        )
        self.assertNotIn("name", window.visible_metrics)

    def test_setting_order_syncs_legacy_flags_and_saves_once(self):
        window = self._window({})
        save = Mock()
        window.set_on_change(save)
        changes = []
        window.display_flags_changed.connect(lambda: changes.append(True))

        window.set_visible_metrics(["amount", "price"])

        self.assertEqual(window.visible_metrics, ["amount", "price"])
        self.assertTrue(window.amount_visible)
        self.assertTrue(window.price_visible)
        self.assertFalse(window.change_pct_visible)
        self.assertEqual(window.current_config()["visible_metrics"], ["amount", "price"])
        save.assert_called_once_with()
        self.refresh.assert_called_once_with()
        self.assertEqual(changes, [True])

    def test_right_click_compatibility_appends_reenabled_metric(self):
        window = self._window({"visible_metrics": ["price"]})

        window.set_flag("成交额", True)
        window.set_flag("买一", True)
        window.set_flag("卖一", False)

        self.assertEqual(window.visible_metrics, ["price", "amount"])
        self.assertFalse(window.b1s1_visible)

    def test_kline_delegate_moves_to_new_column(self):
        window = self._window(
            {"name_visible": False, "visible_metrics": ["kline", "price"]}
        )
        row, roles = self._row_and_roles()
        window._project_columns([row], [roles])
        self.assertIs(window.table.itemDelegateForColumn(0), window.k_delegate)

        window.set_visible_metrics(["price", "kline"])
        window._project_columns([row], [roles])

        self.assertIs(
            window.table.itemDelegateForColumn(0),
            window._default_item_delegate,
        )
        self.assertIs(window.table.itemDelegateForColumn(1), window.k_delegate)

    def test_empty_layout_prompts_until_name_is_enabled(self):
        window = self._window(
            {"name_visible": False, "visible_metrics": []}
        )

        window._project_columns([], [])
        self.assertFalse(window.message_label.isHidden())
        self.assertIn("至少一个显示指标", window.message_label.text())

        window.set_flag("名称", True)
        window._project_columns([], [])

        self.assertEqual(window.model._headers, ["名称"])
        self.assertTrue(window.message_label.isHidden())


if __name__ == "__main__":
    unittest.main()
