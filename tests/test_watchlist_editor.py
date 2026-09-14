"""自选编辑组件可独立使用，且不破坏外层信号屏蔽与列表顺序。"""

import os
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent, QMimeData, QPointF, QSignalBlocker, Qt
from PySide6.QtGui import QDrag, QDropEvent
from PySide6.QtWidgets import QAbstractItemView, QApplication, QPushButton, QWidget

from stockwidget.ui.watchlist_editor import WatchlistEditor
from stockwidget.ui.watchlist_table import WatchlistTable
from shiboken6 import delete, isValid


class WatchlistEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.parent = QWidget()
        self.table = WatchlistTable(0, 3, self.parent)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.codes = {
            "sz000001": {"code": "000001", "market": "sz", "name": "平安银行", "type": "深"},
            "sh600519": {"code": "600519", "market": "sh", "name": "贵州茅台", "type": "沪"},
        }
        self.editor = WatchlistEditor(
            self.table, *(QPushButton(self.parent) for _ in range(3)),
            {}, self.codes, self.parent,
        )

    def tearDown(self):
        if isValid(self.parent):
            self.parent.close()
            self.parent.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)

    def test_deleting_parent_does_not_call_deleted_table(self):
        self.parent.show()
        self.app.processEvents()
        with patch("sys.excepthook") as errors:
            delete(self.parent)
        errors.assert_not_called()
        self.assertFalse(isValid(self.editor))

    def test_deleting_parent_with_active_editor_has_no_late_commits(self):
        self.parent.show()
        self.editor._start_quick_add()
        self.app.processEvents()
        with patch("sys.excepthook") as errors:
            delete(self.parent)
        errors.assert_not_called()

    def test_interpreter_shutdown_with_open_settings_has_no_qt_errors(self):
        script = textwrap.dedent('''
            from unittest.mock import patch
            from PySide6.QtWidgets import QApplication
            from stockwidget.ui.widget import FloatLabel
            from stockwidget.ui.settings_dialog import SettingsDialog
            app = QApplication([])
            with patch.object(FloatLabel, "_refresh_from_function"), patch.object(SettingsDialog, "_start_github_check"):
                window = FloatLabel({}, {})
                dialog = SettingsDialog(window, window)
                dialog.show()
                dialog.watchlist_editor._start_quick_add()
                app.processEvents()
        ''')
        result = subprocess.run(
            [sys.executable, "-c", script], cwd=Path(__file__).parents[1],
            capture_output=True, text=True, timeout=15,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertNotIn("Traceback", result.stderr)

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

    def test_move_return_after_nested_events_preserves_rows_and_later_edits(self):
        self.editor._append_code_row("sz000001", checked=True, cost=10)
        self.editor._append_code_row("sh600519", checked=False, cost=20)
        self.parent.show()
        self.app.processEvents()
        received = []
        self.editor.watchlist_changed.connect(received.append)
        original = self.editor._collect_watchlist_from_list()
        table = self.table

        class InternalDrop(QDropEvent):
            def source(self):
                return table

        # 下移、上移、原地放下、放到最后一行下面，再重复拖动。
        for source, target in ((0, 1), (1, 0), (0, 0), (0, -1), (1, 0)):
            with self.subTest(source=source, target=target):
                table.setCurrentCell(source, 1)
                dragged = table.item(source, 1)
                expected = list(self.editor._collect_watchlist_from_list())
                destination = target if target >= 0 else table.rowCount() - 1
                expected.insert(destination, expected.pop(source))
                if target < 0:
                    pos = QPointF(10, table.rowViewportPosition(1) + table.rowHeight(1) + 10)
                else:
                    pos = QPointF(table.visualItemRect(table.item(target, 1)).center())
                mime_data = QMimeData()
                event = InternalDrop(pos, Qt.MoveAction, mime_data, Qt.LeftButton, Qt.NoModifier)
                before = len(received)

                def native_drag(*_args):
                    self.assertTrue(self.editor.eventFilter(table.viewport(), event))
                    self.assertTrue(event.isAccepted())
                    self.assertEqual(event.dropAction(), Qt.MoveAction)
                    # 模拟 macOS：exec 尚未返回时仍处理 Qt 事件和零延时回调。
                    self.app.processEvents()
                    self.assertIn(dragged, table.selectedItems())
                    return Qt.MoveAction

                with patch.object(QDrag, "exec", side_effect=native_drag):
                    table.startDrag(Qt.MoveAction)
                self.app.processEvents()
                self.assertEqual(table.rowCount(), 2)
                self.assertIs(table.currentItem(), dragged)
                self.assertEqual(table.currentRow(), destination)
                self.assertEqual(len(received), before + 1)
                self.assertEqual(list(received[-1]), expected)
                self.assertEqual(received[-1], original)
                self.assertEqual(table.state(), QAbstractItemView.NoState)

        table.item(0, 2).setText("30")
        self.assertEqual(list(received[-1]), expected)
        self.assertEqual(received[-1][expected[0]]["cost"], 30)
        self.assertEqual(set(received[-1]), set(original))

    def test_cancelled_drag_and_external_drop_leave_watchlist_unchanged(self):
        self.editor._append_code_row("sz000001", checked=True, cost=10)
        self.table.setCurrentCell(0, 1)
        received = []
        self.editor.watchlist_changed.connect(received.append)
        original = self.editor._collect_watchlist_from_list()
        with patch.object(QDrag, "exec", return_value=Qt.IgnoreAction):
            self.table.startDrag(Qt.MoveAction)
        mime_data = QMimeData()
        event = QDropEvent(QPointF(10, 10), Qt.MoveAction, mime_data, Qt.LeftButton, Qt.NoModifier)
        self.editor.eventFilter(self.table.viewport(), event)
        self.app.processEvents()
        self.assertFalse(event.isAccepted())
        self.assertEqual(received, [])
        self.assertEqual(self.editor._collect_watchlist_from_list(), original)
