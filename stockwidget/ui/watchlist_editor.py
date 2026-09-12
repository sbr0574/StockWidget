"""自选表格的编辑、搜索、排序和添加面板，通过信号提交完整自选列表。"""

from PySide6.QtCore import (
    Qt, QPoint, QEvent, QTimer, QItemSelectionModel, QModelIndex, QObject,
    QSignalBlocker, Signal,
)
from PySide6.QtGui import QGuiApplication, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication, QAbstractItemView, QTableWidgetItem, QAbstractItemDelegate,
    QStyledItemDelegate, QStyle, QStyleOptionViewItem, QLineEdit, QCompleter,
    QHeaderView, QLabel,
)
from shiboken6 import isValid

from stockwidget.core.code_search import build_search_index, search_suggestions
from stockwidget.core.watchlist import parse_positive_cost
from stockwidget.ui.add_code_panel import (
    ADDED_ROLE, ENTRY_ROLE, AddCodePanel, entry_display_text,
)

SEARCH_PLACEHOLDER = "搜索代码、名称、拼音或缩写，空格区分关键词"

class CodeSearchEditor(QLineEdit):
    """自选列表的全范围快速搜索输入框。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText(SEARCH_PLACEHOLDER)
        self.setToolTip(SEARCH_PLACEHOLDER)



class CodeCompleterDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def createEditor(self, parent, option, index):
        self.owner._ensure_search_index()
        editor = CodeSearchEditor(parent)
        editor.setProperty("_row", index.row())
        editor.setProperty("_code_editor_committed", False)
        editor.setProperty("_code_editor_initialized", False)
        completer = QCompleter(self.owner.suggestion_model, editor)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        completer.setWrapAround(False)
        completer.activated["QModelIndex"].connect(
            lambda index: self._accept_suggestion(editor, index)
        )
        editor.textEdited.connect(lambda text: self.owner._update_suggestions(editor, text))
        completer.setWidget(editor)
        editor._code_completer = completer
        return editor

    def setEditorData(self, editor, index):
        if editor.property("_code_editor_initialized"):
            return
        editor.setProperty("_code_editor_initialized", True)
        self.owner._remember_editor_value(editor, index)
        editor.clear()

    def setModelData(self, editor, model, index):
        self._commit_editor(editor)

    def destroyEditor(self, editor, index):
        if self.owner.is_active() and not editor.property("_code_editor_committed"):
            self.owner._cancel_code_editor(editor)
        super().destroyEditor(editor, index)

    def _accept_suggestion(self, editor, index):
        if editor.property("_code_editor_committed"):
            return
        if not self.owner._apply_suggestion(editor, index):
            return
        self.commitData.emit(editor)
        self._commit_editor(editor)
        self.closeEditor.emit(editor, QAbstractItemDelegate.EndEditHint.NoHint)

    def _commit_editor(self, editor):
        if not self.owner.is_active() or editor.property("_code_editor_committed"):
            return
        editor.setProperty("_code_editor_committed", True)
        self.owner._commit_code_editor(editor)



class CenteredCheckBoxDelegate(QStyledItemDelegate):
    """同步居中绘制与点击区域，避免勾选框显示位置和命中位置不一致。"""

    def _centered_check_rect(self, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        widget = opt.widget
        style = widget.style() if widget is not None else QApplication.style()
        rect = style.subElementRect(QStyle.SE_ItemViewItemCheckIndicator, opt, widget)
        rect.moveCenter(QPoint(opt.rect.center().x(), rect.center().y()))
        return opt, rect, style, widget

    def paint(self, painter, option, index):
        opt, check_rect, style, widget = self._centered_check_rect(option, index)
        style.drawPrimitive(QStyle.PE_PanelItemViewItem, opt, painter, widget)
        # 勾选框绘制：与 Qt 原生 CE_ItemViewItem 一致，需根据 checkState 手动设置
        # State_On / State_Off / State_NoChange，PE_IndicatorItemViewItemCheck 才会
        # 画出勾选/未勾选状态（否则永远显示为未勾选）。
        opt.rect = check_rect
        opt.state = opt.state & ~QStyle.State_HasFocus
        if opt.checkState == Qt.CheckState.Checked:
            opt.state |= QStyle.State_On
        elif opt.checkState == Qt.CheckState.PartiallyChecked:
            opt.state |= QStyle.State_NoChange
        else:
            opt.state |= QStyle.State_Off
        style.drawPrimitive(QStyle.PE_IndicatorItemViewItemCheck, opt, painter, widget)

    def editorEvent(self, event, model, option, index):
        if (event.type() == QEvent.MouseButtonRelease
                and event.button() == Qt.LeftButton
                and index.flags() & Qt.ItemIsUserCheckable):
            _, check_rect, _, _ = self._centered_check_rect(option, index)
            if check_rect.contains(event.position().toPoint()):
                # index.data(CheckStateRole) 在 PySide6 中返回整数而非枚举，
                # 需按数值比较（CheckState.Checked == 2）
                state = index.data(Qt.CheckStateRole)
                checked = getattr(state, "value", state) == 2
                new_state = Qt.CheckState.Unchecked if checked else Qt.CheckState.Checked
                return model.setData(index, new_state, Qt.ItemDataRole.CheckStateRole)
        return super().editorEvent(event, model, option, index)



class WatchlistEditor(QObject):
    # Python 对象保留插入顺序；Qt 的 QVariantMap 会按键排序。
    watchlist_changed = Signal(object)

    def __init__(self, table, add_button, delete_button, top_button, watchlist, codes, parent):
        super().__init__(table)
        self.list_codes = table
        self._viewport = table.viewport()
        self.btn_add = add_button
        self.btn_del = delete_button
        self.btn_top = top_button
        self._codes = codes
        self._search_index_source = None
        self._search_index_size = -1
        self._search_index = ()
        self._search_entries_by_key = {}
        self.suggestion_model = QStandardItemModel(self)
        self.add_code_panel = AddCodePanel(SEARCH_PLACEHOLDER, parent)
        self._init_code_table(watchlist)
        self.list_codes.itemChanged.connect(self._on_codes_changed)
        self.add_code_panel.entry_requested.connect(self._add_entry_from_panel)
        self.btn_add.clicked.connect(self._show_add_code_panel)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_top.clicked.connect(self._top_code)
        self.btn_top.setToolTip("选中条目移动到列表顶部")
        self.list_codes.itemSelectionChanged.connect(self._update_watchlist_action_buttons)
        self.list_codes.currentCellChanged.connect(self._update_watchlist_action_buttons)
        self._update_watchlist_action_buttons()

    def set_theme(self, dark: bool):
        self.add_code_panel.set_theme(dark)

    def is_active(self):
        return isValid(self.list_codes) and isValid(self._viewport)

    def close(self):
        self.add_code_panel.hide()
        self._on_codes_changed(None)

    def eventFilter(self, obj, ev):
        # 表格析构会先销毁模型，再销毁子对象；此时 Python 回调仍可能到达。
        if not self.is_active() or obj is not self._viewport:
            return False
        if ev.type() == QEvent.Resize:
            self._refresh_empty_watchlist_hint()
        elif ev.type() == QEvent.MouseButtonPress:
            # 单击自选列表空白处：清除选中条目与焦点
            pos = ev.position().toPoint()
            if self.list_codes.itemAt(pos) is None:
                self.list_codes.setCurrentCell(-1, -1)
        elif ev.type() == QEvent.MouseButtonDblClick:
            pos = ev.position().toPoint()
            if self.list_codes.itemAt(pos) is None:
                self._start_quick_add()
                return True
        elif ev.type() == QEvent.Drop:
            self._handle_drop(ev)
            return True
        return super().eventFilter(obj, ev)

    def _init_code_table(self, watchlist):
        self.list_codes.setHorizontalHeaderLabels(["显示", "代码", "成本"])
        self.list_codes.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.list_codes.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.list_codes.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        # 单击不进入编辑，便于整行拖动排序
        self.list_codes.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.AnyKeyPressed
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.list_codes.setDropIndicatorShown(True)
        self.list_codes.setItemDelegateForColumn(0, CenteredCheckBoxDelegate(self))
        self.list_codes.setItemDelegateForColumn(1, CodeCompleterDelegate(self))

        self.empty_watchlist_hint = QLabel(
            "双击空白处添加条目", self._viewport
        )
        self.empty_watchlist_hint.setObjectName("empty_watchlist_hint")
        self.empty_watchlist_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_watchlist_hint.setWordWrap(True)
        self.empty_watchlist_hint.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True
        )
        self.list_codes.model().rowsInserted.connect(self._refresh_empty_watchlist_hint)
        self.list_codes.model().rowsRemoved.connect(self._refresh_empty_watchlist_hint)
        self.list_codes.model().modelReset.connect(self._refresh_empty_watchlist_hint)

        for code, entry in watchlist.items():
            checked = bool(entry.get("checked", True))
            cost = entry.get("cost")
            self._append_code_row(code, entry.get("name", ""), checked, cost)
        self._refresh_empty_watchlist_hint()
        # 打开面板时不预选任何条目
        self.list_codes.setCurrentCell(-1, -1)
        self._viewport.installEventFilter(self)

    def refresh_code_search(self, codes):
        """代码表异步替换后刷新快速搜索和添加面板。"""
        self._codes = codes
        self._search_index_source = None
        self._ensure_search_index()
        for editor in self.list_codes.findChildren(CodeSearchEditor):
            if editor.isVisible():
                self._update_suggestions(editor, editor.text())
        if self.add_code_panel.isVisible():
            self.add_code_panel.set_context(
                self._search_index, self._existing_code_keys()
            )

    def _refresh_empty_watchlist_hint(self, *_args):
        if not self.is_active() or not isValid(self.empty_watchlist_hint):
            return
        self.empty_watchlist_hint.setGeometry(self._viewport.rect())
        self.empty_watchlist_hint.setVisible(self.list_codes.rowCount() == 0)
        self.empty_watchlist_hint.raise_()

    def _handle_drop(self, ev):
        """拖动调整顺序：将拖动的行移动到目标位置。
        1. 用 CopyAction（而非 MoveAction）结束拖放，让 drag->exec() 不返回
           MoveAction，从而不触发 startDrag() 里的 clearOrRemove()；
        2. 暂时清空选中，等拖放清理完成后恢复拖动行的高亮与焦点。
        否则源行会被二次删除，表现为拖拽后丢一行。
        """
        src_row = self.list_codes.currentRow()
        pos = ev.position().toPoint()
        target_row = self.list_codes.rowAt(pos.y())
        if ev.source() is not self.list_codes or src_row < 0:
            ev.ignore()
            return
        if target_row < 0:
            target_row = self.list_codes.rowCount() - 1
        if target_row != src_row:
            self._move_row(src_row, target_row)
        dragged_item = self.list_codes.item(target_row, 1)
        self.list_codes.clearSelection()
        ev.setDropAction(Qt.CopyAction)
        ev.accept()

        def finish_drop():
            if not self.is_active():
                return
            if dragged_item is not None and isValid(dragged_item):
                self.list_codes.setCurrentItem(
                    dragged_item,
                    QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows,
                )
                self.list_codes.setFocus(Qt.MouseFocusReason)
            self._on_codes_changed(None)

        QTimer.singleShot(0, self, finish_drop)

    def _append_code_row(self, code: str = "", name: str = "", checked: bool = False, cost=None):
        self._insert_code_row(
            self.list_codes.rowCount(), code, name, checked, cost
        )

    def _insert_code_row(self, row: int, code: str = "", name: str = "",
                         checked: bool = False, cost=None):
        with QSignalBlocker(self.list_codes):
            row = max(0, min(int(row), self.list_codes.rowCount()))
            self.list_codes.insertRow(row)
            self._set_code_row(row, code, code, name, checked, cost)

    def _ensure_search_index(self):
        """代码表对象变化时重建索引；普通查询复用已规范化记录。"""
        codes = self._codes if isinstance(self._codes, dict) else {}
        if codes is self._search_index_source and len(codes) == self._search_index_size:
            return
        self._search_index = build_search_index(codes)
        self._search_entries_by_key = {
            record.entry["key"]: record.entry for record in self._search_index
        }
        self._search_index_source = codes
        self._search_index_size = len(codes)

    def _entry_for_text(self, text: str) -> dict | None:
        self._ensure_search_index()
        key = str(text or "").strip().casefold()
        if key in self._search_entries_by_key:
            return self._search_entries_by_key[key]
        suggestions = search_suggestions(
            self._search_index,
            text,
            limit=1,
        )
        return suggestions[0] if suggestions else None

    def _set_code_row(self, row: int, value_key: str, display_code: str = "", name: str = "",
                      checked: bool = False, cost=None):
        with QSignalBlocker(self.list_codes):
            value_key = str(value_key or "").strip().lower()
            display_code = str(display_code or "").strip()
            cost = parse_positive_cost(cost)
            entry = dict(self._entry_for_text(value_key) or self._entry_for_text(display_code) or {})
            entry["code"] = entry.get("code") or display_code or value_key
            entry["name"] = str(name or "").strip() or entry.get("name", "")

            check_item = self.list_codes.item(row, 0)
            if check_item is None:
                check_item = QTableWidgetItem("")
                check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
                self.list_codes.setItem(row, 0, check_item)
            check_item.setCheckState(Qt.Checked if checked else Qt.Unchecked)

            code_item = self.list_codes.item(row, 1)
            if code_item is None:
                code_item = QTableWidgetItem("")
                code_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
                self.list_codes.setItem(row, 1, code_item)
            code_item.setText(entry_display_text(entry))
            code_item.setData(Qt.UserRole, value_key)

            cost_item = self.list_codes.item(row, 2)
            if cost_item is None:
                cost_item = QTableWidgetItem("")
                self.list_codes.setItem(row, 2, cost_item)
            if entry.get("type") == "指":
                # 指数不允许设置成本
                cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
                cost_item.setText("")
                cost_item.setData(Qt.UserRole, None)
            else:
                cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
                cost_item.setText("" if cost is None else f"{cost:g}")
                cost_item.setData(Qt.UserRole, cost)

    def _cleanup_code_rows(self):
        with QSignalBlocker(self.list_codes):
            seen = set()
            remove_rows = []
            # 从上到下判重，保留原有条目及其成本、勾选状态和排序。
            for row in range(self.list_codes.rowCount()):
                code_item = self.list_codes.item(row, 1)
                if code_item is None:
                    remove_rows.append(row)
                    continue
                value = str(code_item.data(Qt.UserRole) or "").strip().lower()
                text = str(code_item.text() or "").strip()
                if not text and not value:
                    remove_rows.append(row)
                    continue
                if value and value in seen:
                    remove_rows.append(row)
                    continue
                if value:
                    seen.add(value)
            for row in reversed(remove_rows):
                self.list_codes.removeRow(row)

    def _collect_watchlist_from_list(self):
        """从表格行收集自选列表及其显式 code/market 元数据。"""
        watchlist = {}
        for row in range(self.list_codes.rowCount()):
            check_item = self.list_codes.item(row, 0)
            code_item = self.list_codes.item(row, 1)
            cost_item = self.list_codes.item(row, 2)
            if code_item is None:
                continue

            value = str(code_item.data(Qt.UserRole) or "").strip().lower()
            if not value:
                value = str(code_item.text() or "").strip().lower()
            if not value:
                continue

            resolved = self._entry_for_text(value)
            type_ = str(resolved.get("type", "") or "") if resolved else ""
            entry = {
                "checked": False,
                "cost": None,
                "name": "",
                "type": type_,
                "code": str(resolved.get("code", "") or "") if resolved else "",
                "market": str(resolved.get("market", "") or "") if resolved else "",
            }
            if check_item is not None and check_item.checkState() == Qt.Checked:
                entry["checked"] = True
            # 指数不允许设置成本
            if type_ != "指" and cost_item is not None:
                entry["cost"] = parse_positive_cost(cost_item.text())
            if resolved:
                entry["name"] = str(resolved.get("name", "") or "")
            watchlist[value] = entry
        return watchlist

    def _on_codes_changed(self, _item):
        if not self.is_active():
            return
        self._cleanup_code_rows()
        watchlist = self._collect_watchlist_from_list()
        self.watchlist_changed.emit(watchlist)
        self._refresh_empty_watchlist_hint()
        if self.add_code_panel.isVisible():
            self.add_code_panel.set_context(
                self._search_index, self._existing_code_keys()
            )

    def _start_quick_add(self):
        """双击空白处时创建临时行并启动全范围快速搜索。"""
        self._append_code_row("", "", True)
        row = self.list_codes.rowCount() - 1
        self.list_codes.setCurrentCell(row, 1)
        self.list_codes.editItem(self.list_codes.item(row, 1))

    def _show_add_code_panel(self):
        """从添加按钮打开筛选、分页面板，不预先创建表格行。"""
        for editor in self.list_codes.findChildren(CodeSearchEditor):
            if not editor.property("_code_editor_committed"):
                editor.setProperty("_code_editor_committed", True)
                self._cancel_code_editor(editor)
                self.list_codes.closeEditor(
                    editor, QAbstractItemDelegate.EndEditHint.NoHint
                )
        self._ensure_search_index()
        self.add_code_panel.set_context(
            self._search_index, self._existing_code_keys()
        )
        self.add_code_panel.show_for(self.btn_add)

    def _add_entry_from_panel(self, entry):
        if not isinstance(entry, dict):
            return
        key = str(entry.get("key", "") or "").strip().casefold()
        if not key or key in self._existing_code_keys():
            return
        self._insert_code_row(
            0,
            key,
            str(entry.get("name", "") or ""),
            True,
        )
        self.list_codes.setCurrentCell(0, 1)
        self._on_codes_changed(None)

    def clear_watchlist(self):
        self.add_code_panel.hide()
        with QSignalBlocker(self.list_codes):
            for editor in self.list_codes.findChildren(QLineEdit):
                if isinstance(editor, CodeSearchEditor):
                    editor.setProperty("_code_editor_committed", True)
                    editor._code_completer.popup().hide()
                self.list_codes.closeEditor(
                    editor, QAbstractItemDelegate.EndEditHint.RevertModelCache
                )
            self.list_codes.setRowCount(0)
        self._update_watchlist_action_buttons()
        self._on_codes_changed(None)

    def _del_code(self):
        row = self.list_codes.currentRow()
        if row >= 0:
            self.list_codes.removeRow(row)
            self._on_codes_changed(None)

    def _top_code(self):
        """把当前选中的自选条目移动到列表顶部。"""
        row = self.list_codes.currentRow()
        if row <= 0:
            return
        self._move_row(row, 0)
        self._on_codes_changed(None)

    def _update_watchlist_action_buttons(self, *_args):
        """按自选列表选中状态更新置顶/删除按钮：
        无选中条目时删除按钮禁用；条目已在顶部时置顶按钮禁用。"""
        if not all(isValid(widget) for widget in (self.list_codes, self.btn_del, self.btn_top)):
            return
        row = self.list_codes.currentRow()
        self.btn_del.setEnabled(row >= 0)
        self.btn_top.setEnabled(row > 0)

    def _move_row(self, src: int, dst: int):
        """将 src 行移动到 dst 位置"""
        row_count = self.list_codes.rowCount()
        if not (0 <= src < row_count and 0 <= dst < row_count):
            return
        if src == dst:
            return
        with QSignalBlocker(self.list_codes):
            items = [self.list_codes.takeItem(src, c) for c in range(self.list_codes.columnCount())]
            self.list_codes.removeRow(src)
            self.list_codes.insertRow(dst)
            for c, item in enumerate(items):
                if item is not None:
                    self.list_codes.setItem(dst, c, item)
        self.list_codes.setCurrentCell(dst, 1)

    def _existing_code_keys(self) -> set[str]:
        keys = set()
        for row in range(self.list_codes.rowCount()):
            item = self.list_codes.item(row, 1)
            if item is None:
                continue
            key = str(item.data(Qt.UserRole) or "").strip().casefold()
            if key:
                keys.add(key)
        return keys

    def _apply_suggestion(self, editor: QLineEdit, index) -> bool:
        if not index.isValid():
            return False
        flags = index.flags()
        if not (flags & Qt.ItemFlag.ItemIsEnabled
                and flags & Qt.ItemFlag.ItemIsSelectable):
            return False
        if bool(index.data(ADDED_ROLE)):
            return False
        entry = index.data(ENTRY_ROLE)
        if not isinstance(entry, dict):
            return False

        editor.setProperty("_selected_entry", entry)
        code = str(entry.get("code", "") or "").strip()
        if not code:
            code = str(entry.get("key", "") or "").strip()
        editor.setText(code)
        editor.selectAll()
        return True

    def _update_suggestions(self, editor: QLineEdit, text: str):
        self._ensure_search_index()
        query = str(text or "").strip()
        candidates = search_suggestions(
            self._search_index,
            query,
            limit=10,
        )
        existing_keys = self._existing_code_keys()
        self.suggestion_model.clear()
        for entry in candidates:
            added = str(entry.get("key", "") or "").casefold() in existing_keys
            model_item = QStandardItem(entry_display_text(entry, added=added))
            model_item.setEditable(False)
            model_item.setData(entry, ENTRY_ROLE)
            model_item.setData(added, ADDED_ROLE)
            if added:
                model_item.setEnabled(False)
                model_item.setSelectable(False)
            self.suggestion_model.appendRow(model_item)
        editor.setProperty("_selected_entry", None)
        self._show_suggestions_for_editor(editor, bool(candidates))

    def _show_suggestions_for_editor(self, editor: QLineEdit, has_items: bool):
        completer = getattr(editor, "_code_completer", None)
        if completer is None:
            return
        popup = completer.popup()
        if not has_items:
            popup.hide()
            return

        base_width = self.list_codes.columnWidth(1) + self.list_codes.columnWidth(2)
        content_width = max(0, popup.sizeHintForColumn(0))
        content_width += 2 * popup.frameWidth()
        content_width += 2 * popup.style().pixelMetric(QStyle.PM_FocusFrameHMargin)
        if self.suggestion_model.rowCount() > completer.maxVisibleItems():
            content_width += popup.style().pixelMetric(QStyle.PM_ScrollBarExtent)

        screen = editor.screen() or QGuiApplication.primaryScreen()
        available = screen.availableGeometry() if screen is not None else None
        popup_width = max(base_width, content_width)
        if available is not None:
            popup_width = min(popup_width, available.width())

        popup.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        popup.setMinimumWidth(popup_width)
        popup.setMaximumWidth(popup_width)
        rect = editor.rect()
        rect.setWidth(popup_width)
        if available is not None:
            editor_left = editor.mapToGlobal(editor.rect().topLeft()).x()
            room_on_right = available.right() - editor_left + 1
            if popup_width > room_on_right:
                shift = min(popup_width - room_on_right, editor_left - available.left())
                rect.moveLeft(-max(0, shift))
        completer.complete(rect)
        completion_model = completer.completionModel()
        popup.selectionModel().clearSelection()
        popup.setCurrentIndex(QModelIndex())
        for row in range(completion_model.rowCount()):
            index = completion_model.index(row, completer.completionColumn())
            item_flags = index.flags()
            if (item_flags & Qt.ItemFlag.ItemIsEnabled
                    and item_flags & Qt.ItemFlag.ItemIsSelectable):
                selection_flags = (
                    QItemSelectionModel.SelectionFlag.ClearAndSelect
                    | QItemSelectionModel.SelectionFlag.Rows
                )
                popup.selectionModel().setCurrentIndex(index, selection_flags)
                break

    def _commit_code_editor(self, editor: QLineEdit):
        row = editor.property("_row")
        row = int(row) if row is not None else -1
        if row < 0:
            return
        cost_item = self.list_codes.item(row, 2)
        cost = parse_positive_cost(cost_item.text()) if cost_item is not None else None
        entry = editor.property("_selected_entry")
        text = str(editor.text() or "").strip()
        if not isinstance(entry, dict):
            entry = self._entry_for_text(text)

        existing_keys = self._existing_code_keys()
        entry_key = str(entry.get("key", "") or "").casefold() if entry else ""
        if entry and entry_key and entry_key not in existing_keys:
            self._set_code_row(
                row, entry["key"], entry["code"], entry["name"], True, cost
            )
        else:
            # 无效或重复输入均恢复旧行；新增空行则直接移除。
            self._restore_or_remove_row(row, editor.property("_previous_editor_value"))
        self._on_codes_changed(None)

    def _remember_editor_value(self, editor: QLineEdit, index):
        row = index.row()
        code_item = self.list_codes.item(row, 1)
        cost_item = self.list_codes.item(row, 2)
        check_item = self.list_codes.item(row, 0)
        previous = {
            "key": str(code_item.data(Qt.UserRole) or "") if code_item else "",
            "code": str(code_item.text() or "") if code_item else "",
            "cost": parse_positive_cost(cost_item.text()) if cost_item else None,
            "checked": check_item.checkState() == Qt.Checked if check_item else True,
        }
        editor.setProperty("_previous_editor_value", previous)

    def _cancel_code_editor(self, editor: QLineEdit):
        row = editor.property("_row")
        row = int(row) if row is not None else -1
        if row < 0:
            return
        self._restore_or_remove_row(row, editor.property("_previous_editor_value"))
        self._on_codes_changed(None)

    def _restore_or_remove_row(self, row: int, previous):
        if isinstance(previous, dict) and (previous["key"] or previous["code"]):
            self._set_code_row(row, previous["key"], previous["code"], "", previous["checked"], previous.get("cost"))
        elif 0 <= row < self.list_codes.rowCount():
            self.list_codes.removeRow(row)
