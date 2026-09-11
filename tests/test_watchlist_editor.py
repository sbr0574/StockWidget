"""自选编辑组件可独立使用，且不破坏外层信号屏蔽与列表顺序。"""

import os
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent, QSignalBlocker
from PySide6.QtWidgets import QApplication, QPushButton, QTableWidget, QWidget

from stockwidget.ui.watchlist_editor import WatchlistEditor
from shiboken6 import delete, isValid


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
