"""快捷键输入保留焦点，重复输入和仅收到修饰键都不清空已有设置。"""

import os
import unittest
from unittest.mock import Mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.ui.hotkey_sequence_edit import HotkeySequenceEdit


class HotkeySequenceEditTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.editor = HotkeySequenceEdit()
        self.editor.setMaximumSequenceLength(1)
        self.editor.setKeySequence(QKeySequence("Ctrl+Alt+F"))
        self.editor.show()
        self.editor.activateWindow()
        self.editor.setFocus()
        self.app.processEvents()

    def tearDown(self):
        delete(self.editor)
        self.app.processEvents()

    def test_modifiers_only_when_system_consumes_main_key_preserve_sequence(self):
        finished = Mock()
        self.editor.editingFinished.connect(finished)
        QTest.keyPress(self.editor, Qt.Key_Control)
        QTest.keyPress(self.editor, Qt.Key_Alt, Qt.ControlModifier)
        QTest.keyRelease(self.editor, Qt.Key_Alt, Qt.ControlModifier)
        QTest.keyRelease(self.editor, Qt.Key_Control)
        QTest.qWait(1100)
        self.assertEqual(self.editor.keySequence().toString(), "Ctrl+Alt+F")
        self.assertTrue(self.editor.hasFocus())
        finished.assert_not_called()

    def test_repeat_then_change_combination_keeps_focus(self):
        for key, text in ((Qt.Key_F, "Ctrl+Alt+F"), (Qt.Key_G, "Ctrl+Alt+G")):
            with self.subTest(text=text):
                QTest.keyClick(self.editor, key, Qt.ControlModifier | Qt.AltModifier)
                QTest.qWait(1100)
                self.assertEqual(self.editor.keySequence().toString(), text)
                self.assertTrue(self.editor.hasFocus())

    def test_bare_function_key_can_be_entered_and_repeated(self):
        for key, text in ((Qt.Key_F1, "F1"), (Qt.Key_F1, "F1"), (Qt.Key_F12, "F12")):
            with self.subTest(text=text):
                QTest.keyClick(self.editor, key)
                QTest.qWait(1100)
                self.assertEqual(self.editor.keySequence().toString(), text)
                self.assertTrue(self.editor.hasFocus())
