import os
import threading
import time
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QAbstractItemDelegate

from stockwidget.ui.settings_dialog import (
    CodeSearchEditor,
    SettingsDialog,
    _SEARCH_PLACEHOLDER,
    _SUGGESTION_ADDED_ROLE,
    _SUGGESTION_ENTRY_ROLE,
    _build_settings_stylesheet,
)
from stockwidget.ui.widget import FloatLabel


CODES = {
    "sh600519": {
        "code": "600519",
        "market": "sh",
        "name": "贵州茅台",
        "type": "沪",
        "py": "guizhoumaotai",
        "abbr": "gzmt",
    },
    "sh501001": {
        "code": "501001",
        "market": "sh",
        "name": "财通精选混合LOF",
        "type": "基",
        "py": "caitongjingxuanhunhelof",
        "abbr": "ctjxhhlof",
    },
    "sh000001": {
        "code": "000001",
        "market": "sh",
        "name": "上证指数",
        "type": "指",
        "py": "shangzhengzhishu",
        "abbr": "szzs",
    },
    "ad0": {
        "code": "ad0",
        "market": "",
        "name": "铸造铝合金连续",
        "type": "期",
        "py": "zhuzaolvhejinlianxu",
        "abbr": "zzlhjlx",
    },
    "longname": {
        "code": "longname",
        "market": "us",
        "name": "用于验证候选列表根据条目内容自动扩展宽度的特别长证券名称",
        "type": "美",
        "py": "yongyuyanzhengchaotezhengquanmingcheng",
        "abbr": "cctm",
    },
}


class _SignalRecorder:
    def __init__(self):
        self.event = threading.Event()
        self.value = None

    def emit(self, value):
        self.value = value
        self.event.set()


class SettingsDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])
        cls.qt_app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self._windows = []

    def tearDown(self):
        for dialog, window in reversed(self._windows):
            for editor in dialog.list_codes.findChildren(CodeSearchEditor):
                editor.setProperty("_code_editor_committed", True)
                editor._code_completer.popup().hide()
                dialog.list_codes.closeEditor(
                    editor, QAbstractItemDelegate.EndEditHint.NoHint
                )
            self.qt_app.processEvents()
            dialog.close()
            window.close()
            dialog.deleteLater()
            window.deleteLater()
        self.qt_app.processEvents()

    def _make_dialog(self, watchlist=None):
        cfg = {"watchlist": watchlist or {}}
        window = FloatLabel(cfg, CODES)
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(window, window)
        self._windows.append((dialog, window))
        return dialog, window

    def _start_code_editor(self, dialog):
        dialog._add_code()
        self.qt_app.processEvents()
        editor = dialog.list_codes.findChild(CodeSearchEditor)
        self.assertIsNotNone(editor)
        return editor

    def test_macos_button_styles_are_platform_scoped(self):
        regular = _build_settings_stylesheet(dark=False)
        macos = _build_settings_stylesheet(dark=False, macos=True)

        self.assertNotIn("QPushButton#btn_add", regular)
        self.assertIn("QPushButton#btn_add", macos)
        self.assertIn("QPushButton#btn_icon_default", macos)
        self.assertIn("rgb(10, 132, 255)", macos)

    def test_source_toggle_only_updates_for_selected_button(self):
        calls = []
        owner = type(
            "Owner",
            (),
            {"win": type("Window", (), {"set_data_source": calls.append})()},
        )()

        SettingsDialog._on_source_toggled(owner, "sina", False)
        SettingsDialog._on_source_toggled(owner, "eastmoney", True)

        self.assertEqual(calls, ["eastmoney"])

    def test_github_check_does_not_block_dialog_thread(self):
        owner = type("Owner", (), {"github_check_finished": _SignalRecorder()})()

        def slow_check(timeout):
            time.sleep(0.2)
            return False

        with patch(
            "stockwidget.ui.settings_dialog.github_available",
            side_effect=slow_check,
        ):
            started = time.perf_counter()
            SettingsDialog._start_github_check(owner)
            elapsed = time.perf_counter() - started
            self.assertLess(elapsed, 0.1)
            self.assertTrue(owner.github_check_finished.event.wait(1))

        self.assertTrue(owner.github_check_finished.value)

    def test_inline_editor_contains_category_selector_and_placeholder(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)

        self.assertEqual(editor.placeholderText(), _SEARCH_PLACEHOLDER)
        self.assertEqual(
            [editor.category_combo.itemText(i)
             for i in range(editor.category_combo.count())],
            ["股票", "基金", "指数", "期货"],
        )
        self.assertEqual(editor.category(), "stock")
        widest_label = max(
            editor.fontMetrics().horizontalAdvance(
                editor.category_combo.itemText(i)
            )
            for i in range(editor.category_combo.count())
        )
        self.assertEqual(editor._category_width, max(50, widest_label + 28))
        combo_style = editor.category_combo.styleSheet()
        self.assertIn("background-color: rgba(127, 127, 127, 0.14)", combo_style)
        self.assertIn("border-right: 1px solid", combo_style)

    def test_category_is_session_only_and_new_dialog_defaults_to_stock(self):
        dialog, window = self._make_dialog()
        editor = self._start_code_editor(dialog)
        save_callback = Mock()
        window.set_on_change(save_callback)

        editor.category_combo.setCurrentIndex(
            editor.category_combo.findData("fund")
        )
        self.qt_app.processEvents()

        self.assertEqual(dialog._search_category, "fund")
        self.assertFalse(save_callback.called)
        self.assertNotIn("search_type_filters", window.current_config())

        other_dialog, _other_window = self._make_dialog()
        self.assertEqual(other_dialog._search_category, "stock")

    def test_category_change_filters_current_query_immediately(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)
        editor.setText("财通")

        dialog._update_suggestions(editor, editor.text())
        self.assertEqual(dialog.suggestion_model.rowCount(), 0)

        editor.category_combo.setCurrentIndex(
            editor.category_combo.findData("fund")
        )
        self.qt_app.processEvents()

        self.assertEqual(dialog.suggestion_model.rowCount(), 1)
        self.assertEqual(
            dialog.suggestion_model.item(0).data(_SUGGESTION_ENTRY_ROLE)["key"],
            "sh501001",
        )

    def test_added_result_has_prefix_and_cannot_be_selected(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            }
        }
        dialog, _window = self._make_dialog(watchlist)
        editor = self._start_code_editor(dialog)

        dialog._update_suggestions(editor, "茅台")
        item = dialog.suggestion_model.item(0)
        self.assertEqual(item.text(), "（已添加）沪/600519/贵州茅台")
        self.assertTrue(item.data(_SUGGESTION_ADDED_ROLE))
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsEnabled)
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsSelectable)

        completion_index = editor._code_completer.completionModel().index(0, 0)
        self.assertFalse(dialog._apply_suggestion(editor, completion_index))
        self.assertIsNone(editor.property("_selected_entry"))
        self.assertFalse(editor._code_completer.popup().currentIndex().isValid())

    def test_duplicate_manual_commit_preserves_original_row(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            }
        }
        dialog, _window = self._make_dialog(watchlist)
        editor = self._start_code_editor(dialog)
        editor.setText("600519")

        dialog.list_codes.itemDelegateForColumn(1)._commit_editor(editor)

        self.assertEqual(dialog.list_codes.rowCount(), 1)
        self.assertEqual(
            dialog.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(dialog.list_codes.item(0, 2).text(), "123")

    def test_popup_uses_table_width_and_expands_for_long_result(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)
        base_width = (
            dialog.list_codes.columnWidth(1)
            + dialog.list_codes.columnWidth(2)
        )

        dialog._update_suggestions(editor, "茅台")
        popup = editor._code_completer.popup()
        self.assertEqual(popup.minimumWidth(), base_width)
        self.assertEqual(popup.maximumWidth(), base_width)

        dialog._update_suggestions(editor, "特别长证券名称")
        self.assertGreater(popup.minimumWidth(), base_width)
        self.assertEqual(popup.minimumWidth(), popup.maximumWidth())


if __name__ == "__main__":
    unittest.main()
