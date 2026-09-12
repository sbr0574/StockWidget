"""控件销毁后，尚未执行的界面回调应随对应控件一起取消。"""

import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton, QWidget
from shiboken6 import delete

from stockwidget.ui.add_code_panel import AddCodePanel
from stockwidget.ui.widget import FloatLabel


class WidgetLifecycleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_deleting_float_window_cancels_pending_resize(self):
        with patch.object(FloatLabel, "_refresh_from_function"), patch("sys.excepthook") as errors:
            window = FloatLabel({}, {})
            window._defer_fit()
            delete(window)
            self.app.processEvents()
        errors.assert_not_called()

    def test_clearing_watchlist_discards_pending_quotes_and_errors(self):
        with patch.object(FloatLabel, "_refresh_from_function"):
            window = FloatLabel({"watchlist": {"au0": {"checked": True}}}, {})
        try:
            window.model.set_rows_headers([["800"]], ["现价"], [["text"]])
            window._refresh_thread = Mock()
            window._refresh_thread.is_alive.return_value = True
            window.set_watchlist({})
            self.assertEqual(window.model.rowCount(), 0)
            window._process_data((True, {"au0": {"current_price": 800}}, None))
            self.assertEqual(window.model.rowCount(), 0)
            window._process_data((False, None, "网络请求失败"))
            self.assertIn("添加自选股", window.message_label.text())
        finally:
            delete(window)
            self.app.processEvents()

    def test_deleting_add_panel_cancels_pending_focus(self):
        parent = QWidget()
        anchor = QPushButton(parent)
        panel = AddCodePanel("搜索", parent)
        with patch("sys.excepthook") as errors:
            panel.show_for(anchor)
            delete(parent)
            self.app.processEvents()
        errors.assert_not_called()
