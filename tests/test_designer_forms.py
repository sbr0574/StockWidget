"""Designer files must load and remain the source of runtime layout properties."""

import importlib
import os
from pathlib import Path
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QBuffer, QByteArray, QIODevice
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QGroupBox, QWidget
from shiboken6 import delete

from stockwidget.ui.widget import FloatLabel
from stockwidget.ui.view_settings import TaskbarSettings


DIRECTORY = Path(__file__).resolve().parents[1] / "stockwidget/ui/generated"


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


class DesignerFormTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

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
        tree.find(".//layout[@name='taskbarControlsLayout']/property[@name='spacing']/number").text = "9"
        tree.find(".//widget[@name='taskbar_font_size_label']/property[@name='minimumSize']/size/width").text = "47"
        root = load_form(ET.tostring(tree.getroot()))
        with patch.object(FloatLabel, "_refresh_from_function"), patch("stockwidget.ui.widget.GlobalHotkeyManager"):
            source = FloatLabel({}, {})
        try:
            taskbar = root.findChild(QGroupBox, "taskbar_settings")
            binding = TaskbarSettings(taskbar)
            binding.bind(source)
            source.set_view_options(taskbar_enabled=True, taskbar_sync_appearance=False, taskbar_font_size=14)
            self.assertEqual(taskbar.left.layout().spacing(), 9)
            self.assertEqual(taskbar.style.font_size_label.minimumWidth(), 47)
            self.assertEqual(taskbar.style.font_size_label.text(), "14 pt")
            self.assertIs(taskbar.style.font_size, root.findChild(QWidget, "taskbar_font_size"))
        finally:
            delete(root)
            delete(source)

    def test_floating_layout_uses_the_generated_form_widgets(self):
        with patch.object(FloatLabel, "_refresh_from_function"), patch("stockwidget.ui.widget.GlobalHotkeyManager"):
            source = FloatLabel({}, {})
        try:
            self.assertIs(source.panel, source.ui.panel)
            self.assertIs(source.table, source.ui.table)
            self.assertIs(source.pager, source.ui.pager)
            self.assertIs(source.message_label, source.ui.message_label)
        finally:
            delete(source)
