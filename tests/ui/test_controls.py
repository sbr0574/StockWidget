"""快捷键输入和 Designer 源文件的加载 / 布局绑定。"""

from pathlib import Path
from unittest.mock import Mock, patch
import importlib
import xml.etree.ElementTree as ET

from PySide6.QtCore import Qt, QBuffer, QByteArray, QIODevice
from PySide6.QtGui import QKeySequence
from PySide6.QtTest import QTest
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QGroupBox, QWidget
from shiboken6 import delete

from stockwidget.ui.controls.settings_widgets import HotkeySequenceEdit
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.settings.groups import TaskbarSettings

from tests.support import QtTestCase


class HotkeySequenceEditTests(QtTestCase):

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


DIRECTORY = Path(__file__).resolve().parents[2] / "stockwidget/ui/generated"


def load_form(xml):
    loader = QUiLoader()
    for custom in ET.fromstring(xml).findall("customwidgets/customwidget"):
        module = importlib.import_module(custom.findtext("header"))
        loader.registerCustomWidget(getattr(module, custom.findtext("class")))
    buffer = QBuffer()
    buffer.setData(QByteArray(xml))
    buffer.open(QIODevice.ReadOnly)
    root = loader.load(buffer)
    if root is None:
        raise AssertionError(loader.errorString())
    return root


class DesignerFormTests(QtTestCase):

    def test_all_forms_load_with_their_declared_controls(self):
        for path in sorted(DIRECTORY.glob("*.ui")):
            with self.subTest(form=path.name):
                xml = path.read_bytes()
                root = load_form(xml)
                try:
                    declared = {w.get("name") for w in ET.fromstring(xml).iter("widget")}
                    actual = {root.objectName()} | {w.objectName() for w in root.findChildren(QWidget)}
                    self.assertTrue(declared <= actual, declared - actual)
                finally:
                    delete(root)

    def test_designer_spacing_and_label_width_survive_binding_and_sync(self):
        tree = ET.parse(DIRECTORY / "settings.ui")
        tree.find(".//layout[@name='taskbar_settings_layout']/property[@name='spacing']/number").text = "9"
        tree.find(".//widget[@name='taskbar_font_size_label']/property[@name='minimumSize']/size/width").text = "47"
        root = load_form(ET.tostring(tree.getroot()))
        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({}, {})
        try:
            taskbar = root.findChild(QWidget, "taskbar_settings")
            binding = TaskbarSettings(taskbar)
            binding.bind(source)
            source.set_view_options(taskbar_enabled=True, taskbar_sync_appearance=False, taskbar_font_size=14)
            self.assertEqual(taskbar.body.layout().spacing(), 9)
            self.assertEqual(taskbar.style.font_size_label.minimumWidth(), 47)
            self.assertEqual(taskbar.style.font_size_label.text(), "14 pt")
            self.assertIs(taskbar.style.font_size, root.findChild(QWidget, "taskbar_font_size"))
        finally:
            delete(root)
            delete(source)

    def test_floating_layout_uses_the_generated_form_widgets(self):
        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({}, {})
        try:
            self.assertIs(source.panel, source.ui.panel)
            self.assertIs(source.table, source.ui.table)
            self.assertIs(source.pager, source.ui.pager)
            self.assertIs(source.message_label, source.ui.message_label)
        finally:
            delete(source)
