"""添加面板的输入法、键盘导航与关闭行为。"""

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, Qt
from PySide6.QtGui import QInputMethodEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QPushButton, QWidget
from shiboken6 import delete

from stockwidget.core.code_search import build_search_index
from stockwidget.ui.add_code_panel import AddCodePanel, ENTRY_ROLE


class AddCodePanelInputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.parent = QWidget()
        self.anchor = QPushButton("添加", self.parent)
        self.panel = AddCodePanel("搜索", self.parent)
        self.panel.set_context(build_search_index({
            "sh600000": {"code": "600000", "market": "sh", "type": "沪", "name": "浦发银行"},
            "sh600036": {"code": "600036", "market": "sh", "type": "沪", "name": "招商银行"},
            "sh600519": {"code": "600519", "market": "sh", "type": "沪", "name": "贵州茅台"},
        }), {"sh600036"})
        self.added = []
        self.panel.entry_requested.connect(self.added.append)
        self.parent.show()
        self.panel.show_for(self.anchor)
        self.app.processEvents()

    def tearDown(self):
        delete(self.parent)
        self.app.processEvents()

    def test_opening_panel_activates_search_input(self):
        self.assertEqual(self.app.focusWidget(), self.panel.search_input)
        self.assertTrue(self.panel.isActiveWindow())
        self.assertEqual(self.panel.windowType(), Qt.WindowType.Tool)
        self.assertTrue(self.panel.search_input.testAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled))

    def test_arrow_keys_skip_added_rows_without_leaving_search(self):
        search = self.panel.search_input
        self.assertEqual(self.panel.result_list.currentIndex().row(), 0)
        QTest.keyClick(search, Qt.Key.Key_Down)
        self.assertEqual(self.panel.result_list.currentIndex().row(), 2)
        self.assertEqual(self.app.focusWidget(), search)
        QTest.keyClick(search, Qt.Key.Key_Up)
        self.assertEqual(self.panel.result_list.currentIndex().row(), 0)
        QTest.keyClick(search, Qt.Key.Key_Down)
        QTest.keyClick(search, Qt.Key.Key_Return)
        self.assertEqual([entry["key"] for entry in self.added], ["sh600519"])
        QTest.keyClicks(search, "600000")
        self.assertEqual(self.panel.current_result.total, 1)

    def test_chinese_composition_does_not_navigate_add_or_close(self):
        search = self.panel.search_input
        self.app.sendEvent(search, QInputMethodEvent("pu fa", []))
        self.assertTrue(search.composing)
        for key in (Qt.Key.Key_Down, Qt.Key.Key_Up, Qt.Key.Key_Return, Qt.Key.Key_Escape):
            QTest.keyClick(search, key)
        self.assertEqual(self.panel.result_list.currentIndex().row(), 0)
        self.assertTrue(self.panel.isVisible())
        self.assertEqual(self.added, [])

        commit = QInputMethodEvent()
        commit.setCommitString("浦发")
        self.app.sendEvent(search, commit)
        self.assertFalse(search.composing)
        self.assertEqual(search.text(), "浦发")
        self.assertEqual(self.panel.current_result.total, 1)
        self.assertEqual(self.panel.result_list.currentIndex().data(ENTRY_ROLE)["key"], "sh600000")
        QTest.keyClick(search, Qt.Key.Key_Return)
        self.assertEqual([entry["key"] for entry in self.added], ["sh600000"])

    def test_empty_and_all_added_results_do_not_activate(self):
        search = self.panel.search_input
        for query in ("missing", "招商"):
            search.setText(query)
            QTest.keyClick(search, Qt.Key.Key_Down)
            QTest.keyClick(search, Qt.Key.Key_Up)
            QTest.keyClick(search, Qt.Key.Key_Return)
        self.assertEqual(self.added, [])

    def test_outside_click_closes_panel_but_filter_click_does_not(self):
        checkbox = self.panel.category_filters.option_checkboxes["stock"]
        QTest.mouseClick(checkbox, Qt.MouseButton.LeftButton)
        self.assertTrue(self.panel.isVisible())
        QTest.mouseClick(self.anchor, Qt.MouseButton.LeftButton)
        self.assertFalse(self.panel.isVisible())

    def test_escape_and_parent_hide_close_panel(self):
        QTest.keyClick(self.panel.search_input, Qt.Key.Key_Escape)
        self.assertFalse(self.panel.isVisible())
        self.panel.show_for(self.anchor)
        self.app.processEvents()
        self.parent.hide()
        self.assertFalse(self.panel.isVisible())

    def test_switching_away_from_application_closes_panel(self):
        self.app.sendEvent(self.app, QEvent(QEvent.Type.ApplicationDeactivate))
        self.assertFalse(self.panel.isVisible())


if __name__ == "__main__":
    unittest.main()
