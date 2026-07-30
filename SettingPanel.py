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
from PySide6.QtUiTools import QUiLoader

from WidgetPanel import FloatLabel
from code_index import find_suggestions, refresh_index_from_akshare
from stock_data import request_sina


class CodeCompleterDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        completer = QCompleter(self.owner.suggestion_model, editor)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.activated.connect(lambda text: self.owner._apply_suggestion(editor, text))
        editor.textEdited.connect(self.owner._update_suggestions)
        editor.setCompleter(completer)
        return editor


class SettingsDialog(QDialog):
    APP_NAME = "StockWidget"

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self.setModal(False)
        self.code_index = list(getattr(app, "code_index", []) or [])
        self.suggestion_model = QStringListModel(self)
        self._suggestion_map = {}

        self._load_ui()
        self._init_code_table()
        self._bind_widgets()
        self._load_settings()

    def _is_macos(self) -> bool:
        return platform.system() == "Darwin"

    def _display_code_for_ui(self, code: str) -> str:
        value = str(code or "").strip().lower()
        if len(value) == 8 and value[:2] in {"sh", "sz", "bj"}:
            return value[2:]
        return value

    def _load_ui(self):
        ui_path = os.path.join(os.path.dirname(__file__), "settings.ui")
        ui_file = QFile(ui_path)
        if not ui_file.open(QFile.ReadOnly):
            raise RuntimeError(f"Cannot open settings UI file: {ui_path}")
        loader = QUiLoader()
        self.root = loader.load(ui_file, self)
        ui_file.close()
        self.setLayout(QVBoxLayout(self))
        self.layout().setContentsMargins(0, 0, 0, 0)
        self.layout().addWidget(self.root)

    def _find(self, widget_type, name):
        widget = self.root.findChild(widget_type, name)
        if widget is None:
            raise AttributeError(f"Missing widget '{name}' in settings.ui")
        return widget

    def _init_code_table(self):
        self.list_codes = self._find(QTableWidget, "list_codes")
        self.list_codes.setColumnCount(3)
        self.list_codes.setHorizontalHeaderLabels(["启用", "代码", "名称"])
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
        self.slider_interval = self._find(QSlider, "slider_interval")
        self.label_interval = self._find(QLabel, "label_interval")
        self.cb_code = self._find(QCheckBox, "cb_code")
        self.cb_name = self._find(QCheckBox, "cb_name")
        self.cb_type = self._find(QCheckBox, "cb_type")
        self.cb_price = self._find(QCheckBox, "cb_price")
        self.cb_diff = self._find(QCheckBox, "cb_diff")
        self.cb_pct = self._find(QCheckBox, "cb_pct")
        self.cb_vol = self._find(QCheckBox, "cb_vol")
        self.cb_amount = self._find(QCheckBox, "cb_amount")
        self.cb_avg = self._find(QCheckBox, "cb_avg")
        self.cb_b1s1 = self._find(QCheckBox, "cb_b1s1")
        self.cb_commi = self._find(QCheckBox, "cb_commi")
        self.cb_kline = self._find(QCheckBox, "cb_kline")
        self.cmb_namelen = self._find(QComboBox, "cmb_namelen")
        self.btn_update = self._find(QPushButton, "btn_update")

        self.cb_default_color = self._find(QCheckBox, "cb_default_color")
        self.btn_fg = self._find(QPushButton, "btn_fg_color")
        self.btn_bg = self._find(QPushButton, "btn_bg_color")
        self.slider_bg_alpha = self._find(QSlider, "slider_bg_alpha")
        self.slider_all_alpha = self._find(QSlider, "slider_all_alpha")
        self.label_bg_alpha = self._find(QLabel, "label_bg_alpha")
        self.label_all_alpha = self._find(QLabel, "label_all_alpha")

        self.cmb_family = self._find(QFontComboBox, "cmb_font")
        self.slider_font = self._find(QSlider, "slider_font_size")
        self.slider_line = self._find(QSlider, "slider_line_interval")
        self.label_font = self._find(QLabel, "label_current_font_size")
        self.label_line = self._find(QLabel, "label_current_line_interval")

        self.cmb_icon = self._find(QComboBox, "cmb_icon")
        self.btn_pick_icon = self._find(QPushButton, "btn_icon")
        self.keyseq_hide = self._find(QKeySequenceEdit, "keyseq_hide")
        self.cb_auto_start = self._find(QCheckBox, "cb_auto_start")
        self.cb_head = self._find(QCheckBox, "cb_head")
        self.cb_grid = self._find(QCheckBox, "cb_grid")

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

        self.btn_add = self._find(QPushButton, "btn_add")
        self.btn_del = self._find(QPushButton, "btn_del")
        self.btn_up = self._find(QPushButton, "btn_up")
        self.btn_down = self._find(QPushButton, "btn_down")
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
                self._append_code_row("", "", False)
                row = self.list_codes.rowCount() - 1
                self.list_codes.setCurrentCell(row, 1)
                self.list_codes.editItem(self.list_codes.item(row, 1))
                return True
        return super().eventFilter(obj, ev)

    def _append_code_row(self, code: str = "", name: str = "", checked: bool = False):
        self.list_codes.blockSignals(True)
        row = self.list_codes.rowCount()
        self.list_codes.insertRow(row)

        check_item = QTableWidgetItem("")
        check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        check_item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.list_codes.setItem(row, 0, check_item)

        display_code = self._display_code_for_ui(code)
        code_item = QTableWidgetItem(display_code)
        code_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
        code_item.setData(Qt.UserRole, str(code or "").strip().lower())
        self.list_codes.setItem(row, 1, code_item)

        name_item = QTableWidgetItem(name)
        name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
        name_item.setData(Qt.UserRole, name)
        self.list_codes.setItem(row, 2, name_item)
        self.list_codes.blockSignals(False)

    def _cleanup_empty_rows(self):
        row = 0
        while row < self.list_codes.rowCount():
            code_item = self.list_codes.item(row, 1)
            name_item = self.list_codes.item(row, 2)
            if code_item is None or name_item is None:
                self.list_codes.removeRow(row)
                continue
            code = str(code_item.text()).strip()
            prev_code = str(code_item.data(Qt.UserRole) or "").strip()
            if not code and not prev_code:
                self.list_codes.removeRow(row)
                continue
            row += 1

    def _find_matching_entry(self, raw_code: str, raw_name: str):
        for query in [raw_code, raw_name]:
            query = str(query or "").strip()
            if not query:
                continue
            for entry in find_suggestions(self.code_index, query, limit=3):
                entry_code = str(entry.get("code", "") or "").strip().lower()
                entry_num = str(entry.get("code_num", "") or "").strip().lower()
                entry_name = str(entry.get("name", "") or "").strip().lower()
                query_norm = query.strip().lower()
                if query_norm in {entry_code, entry_num, entry_name}:
                    return entry
        return None

    def _validate_code(self, code: str) -> bool:
        if not code:
            return False
        try:
            result, _ = request_sina([code])
            return bool(result and result[0])
        except Exception:
            return False

    def _collect_codes_from_list(self):
        codes = []
        checked_codes = []
        code_names = {}
        seen = set()

        for row in range(self.list_codes.rowCount()):
            code_item = self.list_codes.item(row, 1)
            name_item = self.list_codes.item(row, 2)
            check_item = self.list_codes.item(row, 0)
            if code_item is None or name_item is None:
                continue
            raw_code = str(code_item.text() or "").strip()
            raw_name = str(name_item.text() or "").strip()
            prev_code = str(code_item.data(Qt.UserRole) or "").strip().lower()
            prev_name = str(name_item.data(Qt.UserRole) or "").strip()

            resolved_code = None
            matched_entry = None
            display_name = raw_name or prev_name
            if raw_code:
                matched_entry = self._find_matching_entry(raw_code, raw_name)
                if matched_entry is not None:
                    resolved_code = str(matched_entry.get("code", "") or "").strip().lower()
                    display_name = str(matched_entry.get("name", "") or "")
                else:
                    resolved_code = self._to_prefixed_code(raw_code)
                    if resolved_code and not self._validate_code(resolved_code):
                        resolved_code = None
                        display_name = "无效代码"
            elif raw_name:
                matched_entry = self._find_matching_entry("", raw_name)
                if matched_entry is not None:
                    resolved_code = str(matched_entry.get("code", "") or "").strip().lower()
                    display_name = str(matched_entry.get("name", "") or "")

            if resolved_code:
                if resolved_code not in seen:
                    seen.add(resolved_code)
                    codes.append(resolved_code)
                if display_name:
                    code_names[resolved_code] = display_name
                if check_item is not None and check_item.checkState() == Qt.Checked:
                    checked_codes.append(resolved_code)
                self.list_codes.blockSignals(True)
                code_item.setText(self._display_code_for_ui(resolved_code))
                code_item.setData(Qt.UserRole, resolved_code)
                if name_item.text() != display_name:
                    name_item.setText(display_name)
                    name_item.setData(Qt.UserRole, display_name)
                self.list_codes.blockSignals(False)
            else:
                if prev_code:
                    self.list_codes.blockSignals(True)
                    code_item.setText(self._display_code_for_ui(prev_code))
                    code_item.setData(Qt.UserRole, prev_code)
                    self.list_codes.blockSignals(False)
                if raw_name == "" and prev_name:
                    self.list_codes.blockSignals(True)
                    name_item.setText(prev_name)
                    self.list_codes.blockSignals(False)
                elif display_name == "无效代码":
                    self.list_codes.blockSignals(True)
                    name_item.setText("无效代码")
                    self.list_codes.blockSignals(False)

        if hasattr(self.win, 'set_code_names') and callable(getattr(self.win, 'set_code_names')):
            self.win.set_code_names(code_names)
        else:
            setattr(self.win, 'code_names', code_names)
        return codes, checked_codes

    def _on_codes_changed(self, _item):
        codes, checked_codes = self._collect_codes_from_list()
        self.win.set_codes(codes)
        self.win.set_checked_codes(checked_codes)

    def _add_code(self):
        self._append_code_row("", "", False)
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

    def _code_from_input(self, text: str) -> str | None:
        s = str(text or "").strip()
        if not s:
            return None
        if s in self._suggestion_map:
            return self._suggestion_map[s]
        token = s.split()[0]
        return self._to_prefixed_code(token)

    def _to_prefixed_code(self, text: str) -> str | None:
        s = str(text or "").strip().lower()
        if not s:
            return None
        if re.match(r'^(sh|sz|bj)\d{6}$', s):
            return s
        if re.match(r'^\d{6}$', s):
            if s[0] in ("6", "5", "9"):
                return f"sh{s}"
            if s[0] in ("0", "1", "2", "3"):
                return f"sz{s}"
            if s[0] in ("4", "8"):
                return f"bj{s}"
        return None

    def _display_text(self, item: dict) -> str:
        return f"{item.get('code', '')} {item.get('name', '')}".strip()

    def _apply_suggestion(self, editor: QLineEdit, text: str):
        code = self._suggestion_map.get(text)
        if code:
            editor.setText(code)
            editor.selectAll()

    def _update_suggestions(self, text: str):
        query = str(text or "").strip()
        candidates = find_suggestions(self.code_index, query, limit=20)
        labels = []
        self._suggestion_map = {}
        for item in candidates:
            label = self._display_text(item)
            labels.append(label)
            self._suggestion_map[label] = item.get("code", "")
        self.suggestion_model.setStringList(labels)

    def _refresh_stock_index(self):
        self.btn_update.setEnabled(False)
        original_text = self.btn_update.text()
        self.btn_update.setText("更新中...")
        try:
            self.code_index = refresh_index_from_akshare(self.APP_NAME)
        except Exception:
            pass
        finally:
            self.btn_update.setEnabled(True)
            self.btn_update.setText(original_text)

    def closeEvent(self, event):
        self._cleanup_empty_rows()
        self._on_codes_changed(None)
        super().closeEvent(event)
