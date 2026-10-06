"""自选规范化、代码搜索和编辑交互。"""

from pathlib import Path
from unittest.mock import patch, Mock
import subprocess
import sys
import textwrap

from PySide6.QtCore import QCoreApplication, QEvent, QMimeData, QPointF, QSignalBlocker, Qt, QPoint
from PySide6.QtGui import QDrag, QDropEvent, QInputMethodEvent, QPalette
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QAbstractItemView, QPushButton, QWidget
from shiboken6 import delete, isValid

from stockwidget.core.watchlist import build_search_index
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.watchlist.add_panel import (
    AddCodePanel,
    ENTRY_ROLE,
    ADDED_ROLE,
    PAGE_SIZE,
    RESULT_ROW_HEIGHT,
    SearchResultDelegate,
)
from stockwidget.ui.watchlist.editor import (
    WatchlistEditor,
    WatchlistTable,
    CodeSearchEditor,
    SEARCH_PLACEHOLDER,
)

from tests.support import CODES, QtTestCase, SettingsTestCase


class WatchlistEditorTests(QtTestCase):

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
            from stockwidget.ui.floating.presenter import QuotePresenter
            from stockwidget.ui.floating.widget import FloatLabel
            from stockwidget.ui.settings.dialog import SettingsDialog
            app = QApplication([])
            with patch.object(QuotePresenter, "refresh"), patch.object(SettingsDialog, "_start_github_check"):
                window = FloatLabel({}, {})
                dialog = SettingsDialog(window, window)
                dialog.show()
                dialog.watchlist_editor._start_quick_add()
                app.processEvents()
        ''')
        result = subprocess.run(
            [sys.executable, "-c", script], cwd=Path(__file__).parents[2],
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


class AddCodePanelInputTests(QtTestCase):

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


class WatchlistDialogTests(SettingsTestCase):
    def test_clear_watchlist_closes_editor_without_restoring_entries(self):
        dialog, window = self._make_dialog({"sh600519": {"checked": False, "cost": 100}})
        editor = self._start_code_editor(dialog)
        editor.setText("sh501001")
        changes = Mock()
        window.set_on_change(changes)
        before = window.current_config()
        dialog.ui.btn_clear_watchlist.click()
        self.qt_app.processEvents()
        self.assertEqual(window.watchlist, {})
        self.assertEqual(dialog.ui.list_codes.rowCount(), 0)
        self.assertFalse(dialog.ui.btn_del.isEnabled())
        self.assertFalse(dialog.ui.btn_top.isEnabled())
        self.assertFalse(dialog.watchlist_editor.add_code_panel.isVisible())
        self.assertEqual(window.current_config(), {**before, "watchlist": {}})
        changes.assert_called_once()

    def test_drop_preserves_dragged_row_selection(self):
        for target in (0, 1):
            with self.subTest(target=target):
                with patch.object(QuotePresenter, "refresh"):
                    dialog, _window = self._make_dialog({
                        "sh600519": {"checked": True},
                        "sh000001": {"checked": True},
                    })
                dialog.show()
                self.qt_app.processEvents()
                table = dialog.watchlist_editor.list_codes
                table.setCurrentCell(0, 1)
                dragged = table.item(0, 1)
                event = Mock()
                event.source.return_value = table
                event.position.return_value.toPoint.return_value = table.visualItemRect(table.item(target, 1)).center()
                dialog.watchlist_editor._handle_drop(event)
                event.setDropAction.assert_called_once_with(Qt.MoveAction)
                self.qt_app.processEvents()
                self.assertIs(table.currentItem(), dragged)
                self.assertIn(dragged, table.selectedItems())
                self.assertEqual(table.currentRow(), target)
                self.assertEqual(table.rowCount(), 2)
                self.assertTrue(table.hasFocus())
                self.assertTrue(dialog.watchlist_editor.btn_del.isEnabled())
                self.assertEqual(dialog.watchlist_editor.btn_top.isEnabled(), target > 0)

    def test_delete_button_disabled_without_watchlist_selection(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)

        self.assertFalse(dialog.watchlist_editor.btn_del.isEnabled())
        self.assertFalse(dialog.watchlist_editor.btn_top.isEnabled())

        dialog.watchlist_editor.list_codes.setCurrentCell(1, 1)
        self.qt_app.processEvents()
        self.assertTrue(dialog.watchlist_editor.btn_del.isEnabled())
        self.assertTrue(dialog.watchlist_editor.btn_top.isEnabled())

        # 点击空白处后取消选中，删除与置顶按钮均禁用
        dialog.watchlist_editor.list_codes.setCurrentCell(-1, -1)
        self.qt_app.processEvents()
        self.assertFalse(dialog.watchlist_editor.btn_del.isEnabled())
        self.assertFalse(dialog.watchlist_editor.btn_top.isEnabled())

    def test_quick_editor_has_no_category_selector_and_searches_all_types(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)

        self.assertEqual(editor.placeholderText(), SEARCH_PLACEHOLDER)
        self.assertFalse(hasattr(editor, "category_combo"))

        expected = {
            "茅台": "sh600519",
            "财通": "sh501001",
            "上证": "sh000001",
            "铝合金": "ad0",
        }
        for query, key in expected.items():
            dialog.watchlist_editor._update_suggestions(editor, query)
            self.assertEqual(
                dialog.watchlist_editor.suggestion_model.item(0).data(ENTRY_ROLE)["key"], key
            )

    def test_empty_hint_is_visible_and_does_not_block_double_click(self):
        dialog, _window = self._make_dialog()
        dialog.show()
        self.qt_app.processEvents()

        hint = dialog.watchlist_editor.empty_watchlist_hint
        self.assertEqual(hint.text(), "双击空白处添加条目")
        self.assertFalse(hint.isHidden())
        self.assertTrue(
            hint.testAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        )

        QTest.mouseDClick(
            dialog.watchlist_editor.list_codes.viewport(),
            Qt.MouseButton.LeftButton,
            pos=QPoint(20, 80),
        )
        self.qt_app.processEvents()

        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 1)
        self.assertTrue(hint.isHidden())
        self.assertIsNotNone(dialog.watchlist_editor.list_codes.findChild(CodeSearchEditor))

    def test_add_button_opens_panel_without_creating_a_row(self):
        dialog, _window = self._make_dialog()
        dialog.move(20, 20)
        dialog.show()
        self.qt_app.processEvents()

        dialog.watchlist_editor.btn_add.click()
        self.qt_app.processEvents()

        panel = dialog.watchlist_editor.add_code_panel
        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 0)
        self.assertTrue(panel.isVisible())
        self.assertEqual(panel.search_input.placeholderText(), SEARCH_PLACEHOLDER)
        self.assertEqual(
            [box.text() for box in panel.category_filters.option_checkboxes.values()],
            ["股票", "基金", "指数", "期货"],
        )
        self.assertEqual(
            [box.text() for box in panel.region_filters.option_checkboxes.values()],
            ["沪", "深", "京", "港", "美", "其他"],
        )
        button_bottom = dialog.watchlist_editor.btn_add.mapToGlobal(
            QPoint(0, dialog.watchlist_editor.btn_add.height())
        ).y()
        self.assertGreaterEqual(panel.y(), button_bottom)

    def test_add_panel_remains_readable_when_switching_color_modes(self):
        dialog, window = self._make_dialog()
        dialog.show()
        dialog.watchlist_editor.btn_add.click()
        panel = dialog.watchlist_editor.add_code_panel
        for mode in ("light", "dark", "light"):
            with self.subTest(mode=mode):
                window.set_view_options(color_mode=mode)
                self.qt_app.processEvents()
                for control, foreground, background in (
                    (panel.page_label, QPalette.WindowText, QPalette.Window),
                    (panel.search_input, QPalette.Text, QPalette.Base),
                    (panel.result_list, QPalette.Text, QPalette.Base),
                ):
                    palette = control.palette()
                    text = palette.color(foreground).lightness()
                    fill_color = palette.color(background)
                    if fill_color.alpha() == 0:
                        fill_color = panel.palette().color(QPalette.Window)
                    fill = fill_color.lightness()
                    self.assertEqual(text > fill, mode == "dark", control.objectName())
                    self.assertGreater(abs(text - fill), 128, control.objectName())

    def test_top_button_moves_selected_row_to_top_and_persists_order(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh501001": {
                "checked": False,
                "cost": 12.5,
                "name": "财通精选混合LOF",
                "type": "基",
                "market": "sh",
                "code": "501001",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, window = self._make_dialog(watchlist)

        self.assertEqual(dialog.watchlist_editor.btn_top.text(), "置顶")
        self.assertIs(dialog.watchlist_editor.btn_top.parentWidget(), dialog.ui.gb_list)
        self.assertLessEqual(
            dialog.watchlist_editor.btn_top.geometry().right(), dialog.ui.gb_list.width()
        )
        self.assertLess(dialog.watchlist_editor.btn_top.iconSize().width(), 20)
        # 没有选中条目时按钮禁用
        self.assertFalse(dialog.watchlist_editor.btn_top.isEnabled())

        dialog.watchlist_editor.list_codes.setCurrentCell(2, 1)
        self.assertTrue(dialog.watchlist_editor.btn_top.isEnabled())

        dialog.watchlist_editor.btn_top.click()

        self.assertEqual(
            dialog.watchlist_editor.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh000001",
        )
        self.assertEqual(
            dialog.watchlist_editor.list_codes.item(1, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(
            list(window.watchlist), ["sh000001", "sh600519", "sh501001"]
        )
        # 置顶后当前行回到顶部，按钮随之禁用
        self.assertEqual(dialog.watchlist_editor.list_codes.currentRow(), 0)
        self.assertFalse(dialog.watchlist_editor.btn_top.isEnabled())

        # 已在顶部的行再点置顶不产生变化
        dialog.watchlist_editor._top_code()
        self.assertEqual(
            list(window.watchlist), ["sh000001", "sh600519", "sh501001"]
        )

    def test_opening_panel_cancels_unfinished_quick_add_row(self):
        dialog, _window = self._make_dialog()
        self._start_code_editor(dialog)
        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 1)

        dialog.watchlist_editor._show_add_code_panel()
        self.qt_app.processEvents()

        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 0)
        self.assertTrue(dialog.watchlist_editor.add_code_panel.isVisible())
        self.assertFalse(dialog.watchlist_editor.empty_watchlist_hint.isHidden())

    def test_filter_select_all_checkbox_uses_three_states(self):
        dialog, _window = self._make_dialog()
        filters = dialog.watchlist_editor.add_code_panel.category_filters
        fund = filters.option_checkboxes["fund"]

        fund.setChecked(False)
        self.assertEqual(
            filters.all_checkbox.checkState(), Qt.CheckState.PartiallyChecked
        )
        filters.all_checkbox.click()
        self.assertEqual(filters.all_checkbox.checkState(), Qt.CheckState.Checked)
        self.assertTrue(all(box.isChecked() for box in filters.option_checkboxes.values()))
        filters.all_checkbox.click()
        self.assertEqual(filters.all_checkbox.checkState(), Qt.CheckState.Unchecked)
        self.assertFalse(any(box.isChecked() for box in filters.option_checkboxes.values()))

    def test_panel_paginates_ten_results_and_resets_after_filter_change(self):
        codes = {
            f"sh{code:06d}": {
                "code": f"{code:06d}",
                "market": "sh",
                "name": f"测试股票{code}",
                "type": "沪",
            }
            for code in range(21)
        }
        dialog, _window = self._make_dialog(codes=codes)
        panel = dialog.watchlist_editor.add_code_panel
        dialog.watchlist_editor._show_add_code_panel()

        self.assertEqual(panel.current_result.total, 21)
        self.assertEqual(panel.current_result.page_count, 3)
        self.assertEqual(panel.result_model.rowCount(), PAGE_SIZE)
        self.qt_app.processEvents()
        self.assertIsInstance(panel.result_list.itemDelegate(), SearchResultDelegate)
        self.assertEqual(
            panel.result_list.viewport().height(), PAGE_SIZE * RESULT_ROW_HEIGHT
        )
        self.assertEqual(
            [panel.result_list.sizeHintForRow(row) for row in range(PAGE_SIZE)],
            [RESULT_ROW_HEIGHT] * PAGE_SIZE,
        )
        self.assertEqual(
            panel.result_list.visualRect(
                panel.result_model.index(PAGE_SIZE - 1, 0)
            ).bottom(),
            panel.result_list.viewport().rect().bottom(),
        )
        panel.next_button.click()
        self.assertEqual(panel.current_result.page, 2)
        self.assertEqual(panel.result_model.rowCount(), PAGE_SIZE)

        panel.category_filters.option_checkboxes["stock"].setChecked(False)
        self.assertEqual(panel.current_result.page, 0)
        self.assertEqual(panel.current_result.total, 0)
        self.assertEqual(panel.result_model.rowCount(), 0)

    def test_panel_adds_complete_entry_and_marks_it_added(self):
        dialog, window = self._make_dialog()
        save_callback = Mock()
        window.set_on_change(save_callback)
        dialog.watchlist_editor._show_add_code_panel()
        panel = dialog.watchlist_editor.add_code_panel

        panel._activate_index(panel.result_model.index(0, 0))

        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 1)
        self.assertIn("sh600519", window.watchlist)
        self.assertEqual(window.watchlist["sh600519"]["market"], "sh")
        self.assertEqual(window.watchlist["sh600519"]["code"], "600519")
        self.assertTrue(window.watchlist["sh600519"]["checked"])
        self.assertEqual(save_callback.call_count, 1)
        self.assertTrue(panel.result_model.item(0).data(ADDED_ROLE))
        self.assertFalse(
            panel.result_model.item(0).flags() & Qt.ItemFlag.ItemIsEnabled
        )
        self.assertEqual(panel.result_model.item(0).text(), "沪/600519/贵州茅台")
        self.assertNotIn("已添加", panel.result_model.item(0).text())

    def test_panel_adds_new_entry_at_start_of_watchlist(self):
        watchlist = {
            "sh501001": {
                "checked": False,
                "cost": 12.5,
                "name": "财通精选混合LOF",
                "type": "基",
                "market": "sh",
                "code": "501001",
            }
        }
        dialog, window = self._make_dialog(watchlist)
        dialog.watchlist_editor._show_add_code_panel()
        panel = dialog.watchlist_editor.add_code_panel

        panel._activate_index(panel.result_model.index(0, 0))

        self.assertEqual(
            dialog.watchlist_editor.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(
            dialog.watchlist_editor.list_codes.item(1, 1).data(Qt.ItemDataRole.UserRole),
            "sh501001",
        )
        self.assertEqual(list(window.watchlist), ["sh600519", "sh501001"])
        self.assertFalse(window.watchlist["sh501001"]["checked"])
        self.assertEqual(window.watchlist["sh501001"]["cost"], 12.5)

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

        dialog.watchlist_editor._update_suggestions(editor, "茅台")
        item = dialog.watchlist_editor.suggestion_model.item(0)
        self.assertEqual(item.text(), "（已添加）沪/600519/贵州茅台")
        self.assertTrue(item.data(ADDED_ROLE))
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsEnabled)
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsSelectable)

        completion_index = editor._code_completer.completionModel().index(0, 0)
        self.assertFalse(dialog.watchlist_editor._apply_suggestion(editor, completion_index))
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

        dialog.watchlist_editor.list_codes.itemDelegateForColumn(1)._commit_editor(editor)

        self.assertEqual(dialog.watchlist_editor.list_codes.rowCount(), 1)
        self.assertEqual(
            dialog.watchlist_editor.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(dialog.watchlist_editor.list_codes.item(0, 2).text(), "123")

    def test_popup_uses_table_width_and_expands_for_long_result(self):
        codes = {**CODES, "longname": {**CODES["longname"], "name": "特别长证券名称" * 12}}
        dialog, _window = self._make_dialog(codes=codes)
        dialog.show()
        self.qt_app.processEvents()
        editor = self._start_code_editor(dialog)
        base_width = (
            dialog.watchlist_editor.list_codes.columnWidth(1)
            + dialog.watchlist_editor.list_codes.columnWidth(2)
        )

        dialog.watchlist_editor._update_suggestions(editor, "茅台")
        popup = editor._code_completer.popup()
        self.assertEqual(popup.minimumWidth(), base_width)
        self.assertEqual(popup.maximumWidth(), base_width)

        dialog.watchlist_editor._update_suggestions(editor, "特别长证券名称")
        self.assertGreater(popup.minimumWidth(), base_width)
        self.assertEqual(popup.minimumWidth(), popup.maximumWidth())
