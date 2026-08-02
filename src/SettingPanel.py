import os
import platform
import re
from functools import partial

from PySide6.QtCore import Qt, QSize, QStringListModel, QFile, QEvent
from PySide6.QtGui import QColor, QFontDatabase, QKeySequence
from PySide6.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QPushButton, QSlider, QLabel, QColorDialog,
    QComboBox, QAbstractItemView, QCheckBox, QTableWidget, QTableWidgetItem,
    QKeySequenceEdit, QFileDialog, QStyledItemDelegate, QLineEdit, QCompleter,
    QFontComboBox, QHeaderView
)
from ui.ui_settings import Ui_SettingDialog
from src.utils import find_suggestions
from src.WidgetPanel import FloatLabel
from services.stock_data import request_sina


class CodeCompleterDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setProperty("_row", index.row())
        editor.setProperty("_column", index.column())
        completer = QCompleter(self.owner.suggestion_model, editor)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.activated.connect(lambda text: self.owner._apply_suggestion(editor, text))
        editor.textEdited.connect(lambda text: self.owner._update_suggestions(editor, text))
        editor.editingFinished.connect(lambda: self.owner._commit_code_editor(editor))
        editor.setCompleter(completer)
        return editor

    def setEditorData(self, editor, index):
        editor.setText(index.data(Qt.DisplayRole) or "")


class SettingsDialog(QDialog):
    APP_NAME = "StockWidget"

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self.ui = Ui_SettingDialog()
        self.ui.setupUi(self)
        self.setModal(False)
        self.suggestion_model = QStringListModel(self)
        self._suggestion_map = {}

        self._init_code_table()
        self._bind_widgets()
        self._load_settings()

    def _is_macos(self) -> bool:
        return platform.system() == "Darwin"

    def _refresh_code_index(self) -> dict:
        if self._refresh_index_fn is None:
            return self.code_index
        try:
            result = self._refresh_index_fn()
        except TypeError:
            result = self._refresh_index_fn(self.APP_NAME)
        if isinstance(result, dict):
            codes = result.get("codes", result)
            if isinstance(codes, dict):
                self.code_index = codes
                return codes
        return self.code_index

    def _display_code_for_ui(self, code: str) -> str:
        value = str(code or "").strip().lower()
        if len(value) == 8 and value[:2] in {"sh", "sz", "bj"}:
            return value[2:]
        return value

    def _init_code_table(self):
        self.list_codes = self.ui.list_codes
        self.list_codes.setColumnCount(3)
        self.list_codes.setHorizontalHeaderLabels(["显示", "代码", "名称"])
        self.list_codes.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.list_codes.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.list_codes.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.list_codes.verticalHeader().setVisible(False)
        self.list_codes.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.list_codes.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_codes.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.list_codes.viewport().installEventFilter(self)
        self.list_codes.setItemDelegateForColumn(1, CodeCompleterDelegate(self))

        for code in self.win.codes:
            name = getattr(self.win, "code_names", {}).get(code, "")
            checked = code in getattr(self.win, "checked_codes", [])
            self._append_code_row(code, name, checked)

    def _bind_widgets(self):
        self.slider_interval = self.ui.slider_interval
        self.label_interval = self.ui.label_interval
        self.cb_code = self.ui.cb_code
        self.cb_name = self.ui.cb_name
        self.cb_type = self.ui.cb_type
        self.cb_price = self.ui.cb_price
        self.cb_diff = self.ui.cb_diff
        self.cb_pct = self.ui.cb_pct
        self.cb_vol = self.ui.cb_vol
        self.cb_amount = self.ui.cb_amount
        self.cb_avg = self.ui.cb_avg
        self.cb_b1s1 = self.ui.cb_b1s1
        self.cb_commi = self.ui.cb_commi
        self.cb_kline = self.ui.cb_kline
        self.cmb_namelen = self.ui.cmb_namelen
        self.btn_update = self.ui.btn_update

        self.cb_default_color = self.ui.cb_default_color
        self.btn_fg = self.ui.btn_fg_color
        self.btn_bg = self.ui.btn_bg_color
        self.slider_bg_alpha = self.ui.slider_bg_alpha
        self.slider_all_alpha = self.ui.slider_all_alpha
        self.label_bg_alpha = self.ui.label_bg_alpha
        self.label_all_alpha = self.ui.label_all_alpha

        self.cmb_family = self.ui.cmb_font
        self.slider_font = self.ui.slider_font_size
        self.slider_line = self.ui.slider_line_interval
        self.label_font = self.ui.label_current_font_size
        self.label_line = self.ui.label_current_line_interval

        self.cmb_icon = self.ui.cmb_icon
        self.btn_pick_icon = self.ui.btn_icon
        self.keyseq_hide = self.ui.keyseq_hide
        self.cb_auto_start = self.ui.cb_auto_start
        self.cb_head = self.ui.cb_head
        self.cb_grid = self.ui.cb_grid

        self.slider_interval.valueChanged.connect(self._on_interval_changed)
        self.btn_update.clicked.connect(self._refresh_stock_index)
        self.list_codes.itemChanged.connect(self._on_codes_changed)

        self.cb_code.toggled.connect(self._on_code_toggled)
        self.cb_name.toggled.connect(self._on_name_toggled)
        self.cb_type.toggled.connect(self._on_type_toggled)
        self.cb_price.toggled.connect(partial(self._on_flag_toggled, "现价"))
        self.cb_diff.toggled.connect(partial(self._on_flag_toggled, "涨跌值"))
        self.cb_pct.toggled.connect(partial(self._on_flag_toggled, "涨跌幅"))
        self.cb_vol.toggled.connect(partial(self._on_flag_toggled, "成交量"))
        self.cb_amount.toggled.connect(partial(self._on_flag_toggled, "成交额"))
        self.cb_avg.toggled.connect(partial(self._on_flag_toggled, "均价"))
        self.cb_b1s1.toggled.connect(self._on_b1s1_toggled)
        self.cb_commi.toggled.connect(partial(self._on_flag_toggled, "委比"))
        self.cb_kline.toggled.connect(partial(self._on_flag_toggled, "K线"))

        self.btn_add = self.ui.btn_add
        self.btn_del = self.ui.btn_del
        self.btn_up = self.ui.btn_up
        self.btn_down = self.ui.btn_down
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_down.clicked.connect(self._move_down)

        self.cmb_namelen.currentIndexChanged.connect(self._on_name_length_changed)
        self.cb_default_color.toggled.connect(self._on_default_color_toggled)
        self.btn_fg.clicked.connect(self.pick_fg)
        self.btn_bg.clicked.connect(self.pick_bg)
        self.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.slider_all_alpha.valueChanged.connect(self.apply_win_opacity)

        self.cmb_family.currentTextChanged.connect(self._on_family_changed)
        self.slider_font.valueChanged.connect(self.apply_font_size)
        self.slider_line.valueChanged.connect(self._on_line_changed)
        self.keyseq_hide.editingFinished.connect(self._on_hotkey_changed)
        self.cb_auto_start.toggled.connect(self._on_start_on_boot_toggled)
        self.cb_head.toggled.connect(self._on_header_toggled)
        self.cb_grid.toggled.connect(self._on_grid_toggled)
        self.cmb_icon.currentIndexChanged.connect(self._on_icon_changed)
        if self._is_macos():
            self.cmb_icon.setEnabled(False)
            self.btn_pick_icon.setEnabled(False)
        
        if self._is_macos():
            self.keyseq_hide.setEnabled(False)

    def _load_settings(self):
        self.slider_interval.setValue(self.win.refresh_seconds)
        self.label_interval.setText(f"{self.win.refresh_seconds}s")
        self.cb_code.setChecked(self.win.header_is_visible("代码"))
        self.cb_name.setChecked(self.win.header_is_visible("名称"))
        self.cb_type.setChecked(bool(getattr(self.win, "type_visible", False)))
        self.cb_price.setChecked(self.win.header_is_visible("现价"))
        self.cb_diff.setChecked(self.win.header_is_visible("涨跌值"))
        self.cb_pct.setChecked(self.win.header_is_visible("涨跌幅"))
        self.cb_vol.setChecked(self.win.header_is_visible("成交量"))
        self.cb_amount.setChecked(self.win.header_is_visible("成交额"))
        self.cb_avg.setChecked(self.win.header_is_visible("均价"))
        self.cb_b1s1.setChecked(self.win.b1s1_visible)
        self.cb_commi.setChecked(self.win.header_is_visible("委比"))
        self.cb_kline.setChecked(self.win.header_is_visible("K线"))

        self.cmb_namelen.clear()
        for length in [0, 1, 2, 3, 4]:
            self.cmb_namelen.addItem(f"{length}个字" if length > 0 else "完整", userData=length)
        idx_name = self.cmb_namelen.findData(self.win.name_length)
        self.cmb_namelen.setCurrentIndex(idx_name if idx_name >= 0 else 0)

        self.cb_default_color.setChecked(self.win.default_color)
        self.btn_fg.setEnabled(not self.win.default_color)
        self.slider_bg_alpha.setValue(int(round(self.win.bg.alpha() / 2.55)))
        self.label_bg_alpha.setText(f"{self.slider_bg_alpha.value()}%")
        self.slider_all_alpha.setValue(int(round(self.win.windowOpacity() * 100)))
        self.label_all_alpha.setText(f"{self.slider_all_alpha.value()}%")

        self.cmb_family.setCurrentText(self.win.font.family())
        self.slider_font.setValue(self.win.font.pointSize())
        self.label_font.setText(f"{self.win.font.pointSize()} pt")
        self.slider_line.setValue(getattr(self.win, "line_extra_px", 4))
        self.label_line.setText(f"+{self.slider_line.value()} px")

        self.keyseq_hide.setKeySequence(QKeySequence(self.win.hotkey))
        self.cb_auto_start.setChecked(bool(self.win.start_on_boot))
        self.cb_head.setChecked(self.win.header_visible)
        self.cb_grid.setChecked(self.win.grid_visible)

        self._setup_icon_choices()

    def _setup_icon_choices(self):
        self.cmb_icon.clear()
        icon_items = [
            ("默认", 'default'),
            ("系统：计算机", 'std:computer'),
            ("系统：网络", 'std:network'),
            ("系统：文件夹", 'std:folder'),
            ("系统：文件", 'std:file'),
            ("系统：回收站", 'std:trash'),
        ]
        for label, val in icon_items:
            self.cmb_icon.addItem(label, userData=val)

        cur_choice = getattr(self.app, '_app_icon_choice', None) if self.app is not None else None
        if cur_choice is None:
            cur_choice = 'default'
        idx = self.cmb_icon.findData(cur_choice)
        if idx < 0 and isinstance(cur_choice, str) and os.path.exists(cur_choice):
            self.cmb_icon.addItem('自定义', userData=cur_choice)
            idx = self.cmb_icon.count() - 1
        self.cmb_icon.setCurrentIndex(idx if idx >= 0 else 0)

    def eventFilter(self, obj, ev):
        if obj is self.list_codes.viewport() and ev.type() == QEvent.MouseButtonDblClick:
            pos = ev.position().toPoint() if hasattr(ev, 'position') else ev.pos()
            if self.list_codes.itemAt(pos) is None:
                self._append_code_row("", "", True)
                row = self.list_codes.rowCount() - 1
                self.list_codes.setCurrentCell(row, 1)
                self.list_codes.editItem(self.list_codes.item(row, 1))
                return True
        return super().eventFilter(obj, ev)

    def _append_code_row(self, code: str = "", name: str = "", checked: bool = False):
        self.list_codes.blockSignals(True)
        row = self.list_codes.rowCount()
        self.list_codes.insertRow(row)
        self._set_code_row(row, code, code, name, checked)
        self.list_codes.blockSignals(False)

    def _resolve_name_for_code(self, value_key: str, display_code: str = "") -> str:
        codes_list = getattr(self.win, "codes_list", None) or {}
        if not isinstance(codes_list, dict):
            return ""
        values = [str(value_key or "").strip().lower(), str(display_code or "").strip().lower()]
        for entry_key, info in codes_list.items():
            entry_key = str(entry_key or "").strip().lower()
            if entry_key in values:
                return str(info.get("name", "") or "")
            entry_code = str(info.get("code", "") or "").strip().lower()
            if entry_code in values:
                return str(info.get("name", "") or "")
        return ""

    def _set_code_row(self, row: int, value_key: str, display_code: str = "", name: str = "", checked: bool = False):
        self.list_codes.blockSignals(True)
        value_key = str(value_key or "").strip().lower()
        display_code = str(display_code or "").strip()
        resolved_name = str(name or "").strip() or self._resolve_name_for_code(value_key, display_code)

        check_item = self.list_codes.item(row, 0)
        if check_item is None:
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            self.list_codes.setItem(row, 0, check_item)
        check_item.setCheckState(Qt.Checked if checked else Qt.Unchecked)

        code_item = self.list_codes.item(row, 1)
        if code_item is None:
            code_item = QTableWidgetItem("")
            code_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.list_codes.setItem(row, 1, code_item)
        code_item.setText(self._display_code_for_ui(display_code or value_key))
        code_item.setData(Qt.UserRole, value_key)

        name_item = self.list_codes.item(row, 2)
        if name_item is None:
            name_item = QTableWidgetItem("")
            name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.list_codes.setItem(row, 2, name_item)
        name_item.setText(resolved_name)
        name_item.setData(Qt.UserRole, resolved_name)
        self.list_codes.blockSignals(False)

    def _cleanup_code_rows(self):
        self.list_codes.blockSignals(True)
        seen = set()
        remove_rows = []
        for row in range(self.list_codes.rowCount() - 1, -1, -1):
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
        for row in remove_rows:
            self.list_codes.removeRow(row)
        self.list_codes.blockSignals(False)

    def _collect_codes_from_list(self):
        codes = []
        checked_codes = []
        code_names = {}
        seen = set()

        for row in range(self.list_codes.rowCount()):
            code_item = self.list_codes.item(row, 1)
            name_item = self.list_codes.item(row, 2)
            check_item = self.list_codes.item(row, 0)
            if code_item is None:
                continue

            value = str(code_item.data(Qt.UserRole) or "").strip().lower()
            if not value:
                value = str(code_item.text() or "").strip().lower()
            if not value:
                continue
            if value not in seen:
                seen.add(value)
                codes.append(value)
            if name_item is not None:
                name = str(name_item.text() or "").strip()
                if name:
                    code_names[value] = name
            if check_item is not None and check_item.checkState() == Qt.Checked:
                checked_codes.append(value)

        if hasattr(self.win, 'set_code_names') and callable(getattr(self.win, 'set_code_names')):
            self.win.set_code_names(code_names)
        else:
            setattr(self.win, 'code_names', code_names)
        return codes, checked_codes

    def _on_codes_changed(self, _item):
        self._cleanup_code_rows()
        codes, checked_codes = self._collect_codes_from_list()
        self.win.set_codes(codes)
        self.win.set_checked_codes(checked_codes)

    def _add_code(self):
        self._append_code_row("", "", True)
        row = self.list_codes.rowCount() - 1
        self.list_codes.setCurrentCell(row, 1)
        self.list_codes.editItem(self.list_codes.item(row, 1))

    def _del_code(self):
        row = self.list_codes.currentRow()
        if row >= 0:
            self.list_codes.removeRow(row)
            self._on_codes_changed(None)

    def _move_up(self):
        row = self.list_codes.currentRow()
        if row > 0:
            self._swap_rows(row, row - 1)
            self.list_codes.setCurrentCell(row - 1, 1)
            self._on_codes_changed(None)

    def _move_down(self):
        row = self.list_codes.currentRow()
        if 0 <= row < self.list_codes.rowCount() - 1:
            self._swap_rows(row, row + 1)
            self.list_codes.setCurrentCell(row + 1, 1)
            self._on_codes_changed(None)

    def _swap_rows(self, row_a: int, row_b: int):
        for col in range(self.list_codes.columnCount()):
            item_a = self.list_codes.takeItem(row_a, col)
            item_b = self.list_codes.takeItem(row_b, col)
            self.list_codes.setItem(row_a, col, item_b)
            self.list_codes.setItem(row_b, col, item_a)

    def _on_interval_changed(self, value: int):
        self.label_interval.setText(f"{value}s")
        self.win.set_refresh_interval(value)

    def _on_code_toggled(self, checked: bool):
        self.win.set_flag("代码", checked)
        self.win.set_code_type(checked)

    def _on_name_toggled(self, checked: bool):
        self.win.set_flag("名称", checked)

    def _on_type_toggled(self, checked: bool):
        if hasattr(self.win, 'set_type_visible'):
            self.win.set_type_visible(checked)
        else:
            setattr(self.win, 'type_visible', bool(checked))
            if hasattr(self.win, '_notify_change'):
                self.win._notify_change()

    def _on_flag_toggled(self, header: str, checked: bool):
        self.win.set_flag(header, checked)

    def _on_b1s1_toggled(self, checked: bool):
        self.win.set_flag("买一", checked)

    def _on_default_color_toggled(self, checked: bool):
        self.btn_fg.setEnabled(not checked)
        self.win.set_default_color(bool(checked))

    def _on_name_length_changed(self, idx: int):
        value = self.cmb_namelen.itemData(idx)
        if isinstance(value, int):
            self.win.set_name_length(value)

    def _on_family_changed(self, fam: str):
        self.win.set_font_family(fam)

    def _on_hotkey_changed(self):
        new_hotkey = self.keyseq_hide.keySequence().toString()
        try:
            self.win.update_hotkey(new_hotkey)
        except Exception:
            pass

    def _on_header_toggled(self, checked: bool):
        self.win.set_header_visible(bool(checked))

    def _on_grid_toggled(self, checked: bool):
        self.win.set_grid_visible(bool(checked))

    def _on_icon_changed(self, idx: int):
        try:
            val = self.cmb_icon.itemData(idx)
            if not val:
                return
            if hasattr(self, 'app') and self.app is not None:
                try:
                    self.app.set_app_icon(val)
                    try:
                        self.app.save_now()
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

    def _on_start_on_boot_toggled(self, checked: bool):
        self.app.set_start_on_boot(bool(checked))

    def pick_fg(self):
        c = QColorDialog.getColor(self.win.fg, self, "选择文字颜色")
        if c.isValid():
            self.win.set_fg_color(c)

    def pick_bg(self):
        base = QColor(self.win.bg)
        base.setAlpha(255)
        c = QColorDialog.getColor(base, self, "选择背景颜色")
        if c.isValid():
            self.win.set_bg_rgb_keep_alpha(c)

    def apply_bg_alpha(self, v: int):
        self.label_bg_alpha.setText(f"{v}%")
        self.win.set_bg_alpha_percent(v)

    def apply_win_opacity(self, v: int):
        self.label_all_alpha.setText(f"{v}%")
        self.win.set_window_opacity_percent(v)

    def apply_font_size(self, v: int):
        self.label_font.setText(f"{v} pt")
        self.win.set_font_size(v)

    def _on_line_changed(self, v: int):
        self.label_line.setText(f"+{v} px")
        self.win.set_line_extra(v)

    def _display_text(self, item: dict) -> str:
        parts = [str(item.get("type", "") or "").strip(), str(item.get("code", "") or "").strip(), str(item.get("name", "") or "").strip()]
        return " / ".join([p for p in parts if p])

    def _apply_suggestion(self, editor: QLineEdit, text: str):
        entry = self._suggestion_map.get(text)
        if isinstance(entry, dict):
            editor.setProperty("_selected_entry", entry)
            code = str(entry.get("code", "") or "").strip() or str(entry.get("key", "") or "").strip()
            editor.setText(code)
            editor.selectAll()
        else:
            editor.setProperty("_selected_entry", None)

    def _update_suggestions(self, editor: QLineEdit, text: str):
        query = str(text or "").strip()
        candidates = find_suggestions(self.win.codes_list, query, limit=10)
        labels = []
        self._suggestion_map = {}
        for item in candidates:
            label = self._display_text(item)
            labels.append(label)
            self._suggestion_map[label] = item
        self.suggestion_model.setStringList(labels)
        editor.setProperty("_selected_entry", None)

    def _commit_code_editor(self, editor: QLineEdit):
        row = int(editor.property("_row") or -1)
        if row < 0:
            return
        entry = editor.property("_selected_entry")
        text = str(editor.text() or "").strip()
        if isinstance(entry, dict):
            self._set_code_row(row, entry.get("key", text), entry.get("code", text), entry.get("name", ""), True)
        else:
            self._set_code_row(row, text, text, "", bool(text))
        self._on_codes_changed(None)

    def _refresh_stock_index(self):
        self.btn_update.setEnabled(False)
        original_text = self.btn_update.text()
        self.btn_update.setText("更新中...")
        try:
            self.code_index = self._refresh_code_index()
        except Exception:
            self.code_index = {}
        finally:
            self.btn_update.setEnabled(True)
            self.btn_update.setText(original_text)

    def closeEvent(self, event):
        self._cleanup_code_rows()
        self._on_codes_changed(None)
        super().closeEvent(event)
