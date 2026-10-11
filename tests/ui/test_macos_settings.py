"""macOS 组框标题的原生边框、留白与勾选交互；可在 Cocoa 会话验证。"""

from unittest.mock import patch

from PySide6.QtCore import Qt
from PySide6.QtGui import QFontInfo
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QGroupBox, QStyle, QStyleOptionGroupBox

from tests.support import SettingsTestCase


class MacOSSettingsTests(SettingsTestCase):
    def test_group_titles_match_content_and_leave_space_for_frames_and_rows(self):
        with patch("sys.platform", "darwin"):
            dialog, window = self._make_dialog()
            dialog.show()
            for mode in ("light", "dark"):
                window.set_view_options(color_mode=mode)
                for page, scroll in ((dialog.ui.general, dialog.ui.general_scroll),
                                     (dialog.ui.data, dialog.ui.data_scroll),
                                     (dialog.ui.floating, dialog.ui.floating_scroll)):
                    dialog.ui.settings_pages.setCurrentWidget(page)
                    self.qt_app.processEvents()
                    groups = [group for group in page.findChildren(QGroupBox) if not group.isHidden()]
                    for group in groups:
                        with self.subTest(mode=mode, group=group.title()):
                            option = QStyleOptionGroupBox()
                            group.initStyleOption(option)
                            label = group.style().subControlRect(
                                QStyle.CC_GroupBox, option, QStyle.SC_GroupBoxLabel, group)
                            frame = group.style().subControlRect(
                                QStyle.CC_GroupBox, option, QStyle.SC_GroupBoxFrame, group)
                            self.assertEqual(QFontInfo(group.font()).pixelSize(),
                                             QFontInfo(dialog.font()).pixelSize())
                            self.assertGreaterEqual(label.height(), group.fontMetrics().height())
                            self.assertTrue(group.rect().contains(label))
                            self.assertGreaterEqual(frame.top() - label.bottom(), 4)
                            self.assertFalse(group.isFlat())
                            item = group.layout().itemAt(0)
                            first = item.widget() or item.layout()
                            top = first.geometry().top()
                            self.assertGreaterEqual(top, 25)
                            self.assertGreaterEqual(top - frame.top(), 10)
                    for previous, current in zip(groups, groups[1:]):
                        if previous.parentWidget() is current.parentWidget():
                            self.assertGreaterEqual(current.y() - previous.geometry().bottom(), 10)
                    self.assertEqual(scroll.horizontalScrollBar().maximum(), 0)

    def test_checkable_title_and_checkbox_toggle_after_font_and_theme_changes(self):
        with patch("sys.platform", "darwin"):
            dialog, window = self._make_dialog()
            dialog.show()
            dialog.ui.settings_pages.setCurrentWidget(dialog.ui.general)
            group = dialog.ui.gb_scheduled_hide
            for mode in ("light", "dark"):
                window.set_view_options(color_mode=mode)
                dialog.ui.general_scroll.ensureWidgetVisible(group)
                self.qt_app.processEvents()
                option = QStyleOptionGroupBox()
                group.initStyleOption(option)
                for control in (QStyle.SC_GroupBoxLabel, QStyle.SC_GroupBoxCheckBox):
                    hit = group.style().subControlRect(QStyle.CC_GroupBox, option, control, group)
                    self.assertTrue(group.rect().contains(hit))
                    before = window.hide_enabled
                    QTest.mouseClick(group, Qt.LeftButton, pos=hit.center())
                    self.qt_app.processEvents()
                    self.assertEqual(window.hide_enabled, not before)
                    self.assertEqual(group.isChecked(), window.hide_enabled)
                    self.assertEqual(dialog.ui.btn_add_hide_time.isEnabled(), window.hide_enabled)
