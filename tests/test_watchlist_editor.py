"""自选编辑组件可独立使用，且不破坏外层信号屏蔽与列表顺序。"""

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSignalBlocker
from PySide6.QtWidgets import QApplication, QPushButton, QTableWidget, QWidget

from stockwidget.ui.watchlist_editor import WatchlistEditor


class WatchlistEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.parent = QWidget()
        self.table = QTableWidget(0, 3, self.parent)
        self.codes = {
            "sz000001": {"code": "000001", "market": "sz", "name": "平安银行", "type": "深"},
            "sh600519": {"code": "600519", "market": "sh", "name": "贵州茅台", "type": "沪"},
        }
        self.editor = WatchlistEditor(
            self.table, *(QPushButton(self.parent) for _ in range(3)),
            {}, self.codes, self.parent,
        )

    def tearDown(self):
        self.parent.close()
        self.parent.deleteLater()
        self.app.processEvents()

    def test_emits_watchlist_without_application_or_float_window(self):
        received = []
        self.editor.watchlist_changed.connect(received.append)
        for code in self.codes:
            self.editor._add_entry_from_panel({"key": code})

        self.assertEqual(len(received), 2)
        self.assertEqual(list(received[-1]), ["sh600519", "sz000001"])
        self.assertEqual(received[-1]["sh600519"]["name"], "贵州茅台")

    def test_nested_row_updates_preserve_outer_signal_blocker(self):
        changed = []
        self.table.itemChanged.connect(changed.append)
        with QSignalBlocker(self.table):
            self.editor._append_code_row("sz000001", checked=True, cost=10)
            self.assertTrue(self.table.signalsBlocked())
            self.table.item(0, 2).setText("20")
        self.assertEqual(changed, [])
        self.assertFalse(self.table.signalsBlocked())

    def test_code_refresh_rebuilds_search_index(self):
        self.assertIsNone(self.editor._entry_for_text("新标的"))
        codes = {"new": {"code": "new", "market": "us", "name": "新标的", "type": "美"}}
        self.editor.refresh_code_search(codes)
        self.assertEqual(self.editor._entry_for_text("新标的")["key"], "new")
