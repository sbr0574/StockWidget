import os
import shutil
from functools import partial

from PySide6.QtCore import Qt, QStringListModel, QEvent, QTimer
from PySide6.QtGui import QColor, QKeySequence
from PySide6.QtWidgets import (
    QWidget, QDialog, QColorDialog, QAbstractItemView, QTableWidgetItem,
    QStyledItemDelegate, QLineEdit, QCompleter, QHeaderView, QFileDialog,
    QMessageBox
)
from ui.ui_settings import Ui_SettingDialog
from src.utils import code_without_market, find_suggestions, config_paths
from src.WidgetPanel import FloatLabel
from src.platform_support import (
    hotkeys_supported,
    click_through_supported,
    opacity_supported,
    force_top_supported,
    start_on_boot_supported,
    unsupported_tooltip,
    custom_icon_supported,
)
from services.update_check import PROJECT_URL


def _hotkey_error_message(result) -> str:
    """把 HotkeyResult 转成用户可读的中文提示。"""
    if result.reason == "conflict":
        return "该快捷键已被其他程序占用,请更换后重试。"
    if result.reason == "reserved":
        return "该快捷键为系统或通用快捷键(如复制、粘贴、保存等),为避免影响其他应用,请更换为 Ctrl+Alt+某键 之类的组合。"
    if result.reason == "invalid":
        return "快捷键无效,需包含至少一个修饰键(Ctrl/Alt/Shift/Win)和一个主键。"
    if result.reason == "unsupported":
        return "当前平台暂不支持全局快捷键。"
    if result.reason == "permission":
        return ("macOS 需要「辅助功能/输入监听」权限才能使用全局快捷键。\n"
                "请前往 系统设置 → 隐私与安全性 → 辅助功能(或输入监控),"
                "勾选本程序后,重新勾选「启用快捷键」即可。")
    return "快捷键注册失败,请更换后重试。"


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
        completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        completer.activated.connect(lambda text: self.owner._apply_suggestion(editor, text))
        editor.textEdited.connect(lambda text: self.owner._update_suggestions(editor, text))
        editor.editingFinished.connect(lambda: self.owner._commit_code_editor(editor))
        editor.setCompleter(completer)
        return editor

    def setEditorData(self, editor, index):
        editor.setText(index.data(Qt.DisplayRole) or "")
        self.owner._remember_editor_value(editor, index)


class SettingsDialog(QDialog):

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self.ui = Ui_SettingDialog()
        self.ui.setupUi(self)
        self.setModal(False)
        self.suggestion_model = QStringListModel(self)
        self._suggestion_map = {}
        self._previous_editor_values = {}

        self._init_code_table()
        self._bind_widgets()
        self._load_settings()

    def _display_code_for_ui(self, code: str) -> str:
        value = str(code or "").strip().lower()
        if len(value) == 8 and value[:2] in {"sh", "sz", "bj"}:
            return value[2:]
        return code_without_market(value)

    def _init_code_table(self):
        self.list_codes = self.ui.list_codes
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
        self.list_codes.viewport().installEventFilter(self)
        self.list_codes.setItemDelegateForColumn(1, CodeCompleterDelegate(self))

        for code, entry in self.win.watchlist.items():
            checked = bool(entry.get("checked", True))
            cost = entry.get("cost")
            self._append_code_row(code, entry.get("name", ""), checked, cost)

    def _bind_widgets(self):
        self.slider_interval = self.ui.slider_interval
        self.label_interval = self.ui.label_interval
        self.gb_name = self.ui.gb_name
        self.cb_code = self.ui.cb_code
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
        self.cb_profit = self.ui.cb_profit
        self.cmb_namelen = self.ui.cmb_namelen

        self.cb_default_color = self.ui.cb_default_color
        self.btn_fg = self.ui.btn_fg_color
        self.btn_bg = self.ui.btn_bg_color
        self.slider_bg_alpha = self.ui.slider_bg_alpha
        self.slider_all_alpha = self.ui.slider_all_alpha
        self.label_bg_alpha = self.ui.label_bg_alpha
        self.label_all_alpha = self.ui.label_all_alpha
        self.label_all = self.ui.label_all

        self.cmb_family = self.ui.cmb_font
        self.slider_font = self.ui.slider_font_size
        self.slider_line = self.ui.slider_line_interval
        self.label_font = self.ui.label_current_font_size
        self.label_line = self.ui.label_current_line_interval

        self.cmb_icon = self.ui.cmb_icon
        self.btn_pick_icon = self.ui.btn_icon
        self.cb_auto_start = self.ui.cb_auto_start
        self.cb_force_top = self.ui.cb_force_top
        self.cb_click_through = self.ui.cb_click_through
        self.cb_head = self.ui.cb_head
        self.cb_grid = self.ui.cb_grid

        self.cb_hotkey_hide = self.ui.cb_hotkey_hide
        self.cb_hotkey_click_through = self.ui.cb_hotkey_click_through
        self.keyseq_hide = self.ui.keyseq_hide
        self.keyseq_click_through = self.ui.keyseq_click_through

        self.slider_interval.valueChanged.connect(self._on_interval_changed)
        self.list_codes.itemChanged.connect(self._on_codes_changed)

        self.gb_name.toggled.connect(self._on_name_toggled)
        self.cb_code.toggled.connect(self._on_code_toggled)
        self.cb_type.toggled.connect(self._on_type_toggled)
        self.cb_price.toggled.connect(partial(self._on_flag_toggled, "现价"))
        self.cb_diff.toggled.connect(partial(self._on_flag_toggled, "涨跌"))
        self.cb_pct.toggled.connect(partial(self._on_flag_toggled, "涨幅"))
        self.cb_vol.toggled.connect(partial(self._on_flag_toggled, "成交量"))
        self.cb_amount.toggled.connect(partial(self._on_flag_toggled, "成交额"))
        self.cb_avg.toggled.connect(partial(self._on_flag_toggled, "均价"))
        self.cb_b1s1.toggled.connect(self._on_b1s1_toggled)
        self.cb_commi.toggled.connect(partial(self._on_flag_toggled, "委比"))
        self.cb_kline.toggled.connect(partial(self._on_flag_toggled, "K线"))
        self.cb_profit.toggled.connect(partial(self._on_flag_toggled, "浮盈"))

        self.btn_add = self.ui.btn_add
        self.btn_del = self.ui.btn_del
        self.btn_up = self.ui.btn_up
        self.btn_down = self.ui.btn_down
        self.btn_top = self.ui.btn_top
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_down.clicked.connect(self._move_down)
        self.btn_top.clicked.connect(self._top_code)

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
        self.keyseq_click_through.editingFinished.connect(self._on_click_through_hotkey_changed)
        self.cb_auto_start.toggled.connect(self._on_start_on_boot_toggled)
        self.cb_force_top.toggled.connect(self._on_force_top_toggled)
        self.cb_click_through.toggled.connect(self._on_click_through_toggled)
        self.win.click_through_changed.connect(self._sync_click_through_from_win)
        self.cb_hotkey_hide.toggled.connect(self._on_hotkey_hide_enabled_toggled)
        self.cb_hotkey_click_through.toggled.connect(self._on_click_through_hotkey_enabled_toggled)
        self.cb_head.toggled.connect(self._on_header_toggled)
        self.cb_grid.toggled.connect(self._on_grid_toggled)
        self.cmb_icon.currentIndexChanged.connect(self._on_icon_changed)
        self.btn_pick_icon.clicked.connect(self._pick_custom_icon)
        if not custom_icon_supported():
            self.cmb_icon.setEnabled(False)
            self.btn_pick_icon.setEnabled(False)

    def _load_settings(self):
        self.slider_interval.setValue(self.win.refresh_seconds)
        self.label_interval.setText(f"{self.win.refresh_seconds}s")
        self.cb_code.setChecked(self.win.code_visible)
        self.gb_name.setChecked(self.win.name_visible)
        self.cb_type.setChecked(self.win.type_visible)
        self.cb_price.setChecked(self.win.header_is_visible("现价"))
        self.cb_diff.setChecked(self.win.header_is_visible("涨跌"))
        self.cb_pct.setChecked(self.win.header_is_visible("涨幅"))
        self.cb_vol.setChecked(self.win.header_is_visible("成交量"))
        self.cb_amount.setChecked(self.win.header_is_visible("成交额"))
        self.cb_avg.setChecked(self.win.header_is_visible("均价"))
        self.cb_b1s1.setChecked(self.win.b1s1_visible)
        self.cb_commi.setChecked(self.win.header_is_visible("委比"))
        self.cb_kline.setChecked(self.win.header_is_visible("K线"))
        self.cb_profit.setChecked(self.win.profit_visible)

        # 名称显示字数: 0=不显示, -1=全部显示, 1-4=前 N 个字
        # 填充选项时会触发 currentIndexChanged，需屏蔽信号避免意外修改配置
        self.cmb_namelen.blockSignals(True)
        self.cmb_namelen.clear()
        self.cmb_namelen.addItem("不显示", userData=0)
        self.cmb_namelen.addItem("全部显示", userData=-1)
        for length in [1, 2, 3, 4]:
            self.cmb_namelen.addItem(f"{length}个字", userData=length)
        idx_name = self.cmb_namelen.findData(self.win.name_length)
        self.cmb_namelen.setCurrentIndex(idx_name if idx_name >= 0 else 1)
        self.cmb_namelen.blockSignals(False)

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
        self.keyseq_hide.setEnabled(self.win.hotkey_enabled)
        self.keyseq_click_through.setKeySequence(QKeySequence(self.win.hotkey_click_through))
        self.keyseq_click_through.setEnabled(self.win.hotkey_click_through_enabled)
        self.cb_hotkey_hide.setChecked(self.win.hotkey_enabled)
        self.cb_hotkey_click_through.setChecked(self.win.hotkey_click_through_enabled)
        self.cb_auto_start.setChecked(bool(self.win.start_on_boot))
        self.cb_force_top.setChecked(self.win.force_top)
        self.cb_click_through.setChecked(self.win.click_through)
        self.cb_head.setChecked(self.win.header_visible)
        self.cb_grid.setChecked(self.win.grid_visible)

        self._apply_platform_limits()
        self._setup_icon_choices()
        self._setup_about()

    def _apply_platform_limits(self):
        """按当前平台禁用不支持的功能控件:
        - Wayland 下:全局快捷键、鼠标穿透、窗口整体透明度不可用。
        - Linux 下:强制置顶不可用(raise_ 受窗口管理器/合成器限制)。
        """
        if not hotkeys_supported():
            for w in (self.cb_hotkey_hide, self.cb_hotkey_click_through,
                      self.keyseq_hide, self.keyseq_click_through):
                w.setEnabled(False)
                w.setToolTip(unsupported_tooltip("全局快捷键"))
        if not click_through_supported():
            self.cb_click_through.setEnabled(False)
            self.cb_click_through.setToolTip(unsupported_tooltip("鼠标穿透"))
        if not opacity_supported():
            # 整体不透明度滑块:Wayland 平台插件不支持设置窗口透明度
            for w in (self.slider_all_alpha, self.label_all, self.label_all_alpha):
                w.setEnabled(False)
            self.slider_all_alpha.setToolTip(unsupported_tooltip("整体不透明度"))
        if not force_top_supported():
            # 强制置顶:仅 Windows / macOS 支持;Linux 下 raise_ 不可靠
            self.cb_force_top.setEnabled(False)
            self.cb_force_top.setToolTip(unsupported_tooltip("强制置顶", suggest_x11=False))
        if not start_on_boot_supported():
            self.cb_auto_start.setEnabled(False)
            self.cb_auto_start.setToolTip(unsupported_tooltip("开机自启"))

    def _setup_icon_choices(self):
        self.cmb_icon.blockSignals(True)
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

        cur_choice = self.app._icon_choice if self.app is not None else None
        if cur_choice is None:
            cur_choice = 'default'
        idx = self.cmb_icon.findData(cur_choice)
        if idx < 0 and isinstance(cur_choice, str) and os.path.exists(cur_choice):
            self.cmb_icon.addItem('自定义', userData=cur_choice)
            idx = self.cmb_icon.count() - 1
        self.cmb_icon.setCurrentIndex(idx if idx >= 0 else 0)
        self.cmb_icon.blockSignals(False)

    def eventFilter(self, obj, ev):
        if obj is self.list_codes.viewport() and ev.type() == QEvent.MouseButtonDblClick:
            pos = ev.position().toPoint() if hasattr(ev, 'position') else ev.pos()
            if self.list_codes.itemAt(pos) is None:
                self._append_code_row("", "", True)
                row = self.list_codes.rowCount() - 1
                self.list_codes.setCurrentCell(row, 1)
                self.list_codes.editItem(self.list_codes.item(row, 1))
                return True
        if obj is self.list_codes.viewport() and ev.type() == QEvent.Drop:
            self._handle_drop(ev)
            return True
        return super().eventFilter(obj, ev)

    def _handle_drop(self, ev):
        """拖动调整顺序：将拖动的行移动到目标位置"""
        src_row = self.list_codes.currentRow()
        pos = ev.position().toPoint()
        target_row = self.list_codes.rowAt(pos.y())
        if ev.source() is self.list_codes and src_row >= 0:
            if target_row < 0:
                target_row = self.list_codes.rowCount() - 1
            if target_row != src_row:
                self._move_row(src_row, target_row)
        ev.acceptProposedAction()
        QTimer.singleShot(0, lambda: self._on_codes_changed(None))

    def _append_code_row(self, code: str = "", name: str = "", checked: bool = False, cost=None):
        self.list_codes.blockSignals(True)
        row = self.list_codes.rowCount()
        self.list_codes.insertRow(row)
        self._set_code_row(row, code, code, name, checked, cost)
        self.list_codes.blockSignals(False)

    def _entry_for_text(self, text: str) -> dict | None:
        suggestions = find_suggestions(self.win.codes_list, text, limit=1)
        return suggestions[0] if suggestions else None

    def _resolve_name_for_code(self, value_key: str, display_code: str = "") -> str:
        entry = self._entry_for_text(value_key) or self._entry_for_text(display_code)
        return entry["name"] if entry else ""

    def _row_type(self, value_key: str, display_code: str = "") -> str:
        """解析代码对应的市场类型（如 沪/深/创/科/京/基/指）"""
        entry = self._entry_for_text(value_key) or self._entry_for_text(display_code)
        return str(entry.get("type", "") or "") if entry else ""

    def _code_display_for_row(self, value_key: str, display_code: str = "", name: str = "") -> str:
        """生成“类型/代码/名称”格式的合并显示文本"""
        entry = self._entry_for_text(value_key) or self._entry_for_text(display_code)
        if entry:
            type_ = str(entry.get("type", "") or "").strip()
            code = str(entry.get("code", "") or "").strip()
            name = str(name or "").strip() or str(entry.get("name", "") or "").strip()
        else:
            type_ = ""
            code = str(display_code or value_key or "").strip()
            name = str(name or "").strip()
        code = self._display_code_for_ui(code)
        return "/".join([p for p in [type_, code, name] if p])

    def _set_code_row(self, row: int, value_key: str, display_code: str = "", name: str = "",
                      checked: bool = False, cost=None):
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
        code_item.setText(self._code_display_for_row(value_key, display_code, resolved_name))
        code_item.setData(Qt.UserRole, value_key)

        cost_item = self.list_codes.item(row, 2)
        if cost_item is None:
            cost_item = QTableWidgetItem("")
            self.list_codes.setItem(row, 2, cost_item)
        if self._row_type(value_key, display_code) == "指":
            # 指数不允许设置成本
            cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            cost_item.setText("")
            cost_item.setData(Qt.UserRole, None)
        else:
            cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            cost_item.setText("" if cost is None else f"{cost:g}")
            cost_item.setData(Qt.UserRole, cost)
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

    def _collect_watchlist_from_list(self):
        """从表格行收集自选列表（代码 -> {checked, cost, name, type}）"""
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
            entry = {"checked": False, "cost": None, "name": "", "type": type_}
            if check_item is not None and check_item.checkState() == Qt.Checked:
                entry["checked"] = True
            # 指数不允许设置成本
            if type_ != "指" and cost_item is not None:
                cost_text = str(cost_item.text() or "").strip()
                try:
                    cost_val = float(cost_text)
                    if cost_val > 0:
                        entry["cost"] = cost_val
                except ValueError:
                    entry["cost"] = None
            if resolved:
                entry["name"] = str(resolved.get("name", "") or "")
            watchlist[value] = entry
        return watchlist

    def _on_codes_changed(self, _item):
        self._cleanup_code_rows()
        watchlist = self._collect_watchlist_from_list()
        self.win.set_watchlist(watchlist)

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

    def _move_row(self, src: int, dst: int):
        """将 src 行移动到 dst 位置"""
        row_count = self.list_codes.rowCount()
        if not (0 <= src < row_count and 0 <= dst < row_count):
            return
        if src == dst:
            return
        self.list_codes.blockSignals(True)
        items = [self.list_codes.takeItem(src, c) for c in range(self.list_codes.columnCount())]
        self.list_codes.removeRow(src)
        self.list_codes.insertRow(dst)
        for c, item in enumerate(items):
            if item is not None:
                self.list_codes.setItem(dst, c, item)
        self.list_codes.blockSignals(False)
        self.list_codes.setCurrentCell(dst, 1)

    def _move_up(self):
        row = self.list_codes.currentRow()
        if row > 0:
            self._move_row(row, row - 1)
            self._on_codes_changed(None)

    def _move_down(self):
        row = self.list_codes.currentRow()
        if 0 <= row < self.list_codes.rowCount() - 1:
            self._move_row(row, row + 1)
            self._on_codes_changed(None)

    def _top_code(self):
        """将选中个股移到列表第一位"""
        row = self.list_codes.currentRow()
        if row > 0:
            self._move_row(row, 0)
            self._on_codes_changed(None)

    def _on_interval_changed(self, value: int):
        self.label_interval.setText(f"{value}s")
        self.win.set_refresh_interval(value)

    def _on_code_toggled(self, checked: bool):
        self.win.set_code_visible(checked)

    def _on_name_toggled(self, checked: bool):
        self.win.set_flag("名称", checked)

    def _on_type_toggled(self, checked: bool):
        self.win.set_type_visible(checked)

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
        result = self.win.update_hotkey(new_hotkey)
        if not result:
            self.keyseq_hide.setKeySequence(QKeySequence(self.win.hotkey))
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))

    def _on_header_toggled(self, checked: bool):
        self.win.set_header_visible(bool(checked))

    def _on_grid_toggled(self, checked: bool):
        self.win.set_grid_visible(bool(checked))

    def _on_icon_changed(self, idx: int):
        val = self.cmb_icon.itemData(idx)
        if val and self.app is not None:
            self.app.set_app_icon(val)
            self.app.save_now()

    def _pick_custom_icon(self):
        """选择自定义图标并保存到配置文件同目录，立即应用"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择自定义图标", "", "图标文件 (*.ico *.png *.jpg *.jpeg);;所有文件 (*)"
        )
        if not path:
            return
        try:
            dest_dir = config_paths(self.app.app_name)
            os.makedirs(dest_dir, exist_ok=True)
            dest = os.path.join(dest_dir, "custom_icon.ico")
            shutil.copyfile(path, dest)
        except Exception as exc:
            QMessageBox.warning(self, "图标错误", f"无法保存自定义图标：\n{exc}")
            return
        if self.app is not None:
            self.app.set_app_icon(dest)
            self.app.save_now()
        self._setup_icon_choices()

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
        return "/".join([p for p in parts if p])

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
        self._show_suggestions_for_editor(editor, bool(labels))


    def _show_suggestions_for_editor(self, editor: QLineEdit, has_items: bool):
        completer = editor.completer()
        if completer is None:
            return
        popup_width = max(self.list_codes.viewport().width(), editor.width())
        completer.popup().setMinimumWidth(popup_width)
        completer.popup().setMaximumWidth(popup_width)
        if not has_items:
            completer.popup().hide()
            return
        rect = editor.rect()
        rect.setWidth(popup_width)
        completer.complete(rect)

    def _commit_code_editor(self, editor: QLineEdit):
        row = editor.property("_row")
        row = int(row) if row is not None else -1
        if row < 0:
            return
        cost_item = self.list_codes.item(row, 2)
        cost = None
        if cost_item is not None:
            cost_text = str(cost_item.text() or "").strip()
            try:
                cost_val = float(cost_text)
                if cost_val > 0:
                    cost = cost_val
            except ValueError:
                cost = None
        entry = editor.property("_selected_entry")
        text = str(editor.text() or "").strip()
        if isinstance(entry, dict):
            self._set_code_row(row, entry["key"], entry["code"], entry["name"], True, cost)
        else:
            entry = self._entry_for_text(text)
            if entry:
                self._set_code_row(row, entry["key"], entry["code"], entry["name"], True, cost)
            else:
                self._restore_or_remove_row(row)
        self._on_codes_changed(None)


    def _remember_editor_value(self, editor: QLineEdit, index):
        row = index.row()
        code_item = self.list_codes.item(row, 1)
        cost_item = self.list_codes.item(row, 2)
        check_item = self.list_codes.item(row, 0)
        self._previous_editor_values[row] = {
            "key": str(code_item.data(Qt.UserRole) or "") if code_item else "",
            "code": str(code_item.text() or "") if code_item else "",
            "cost": str(cost_item.text() or "") if cost_item else "",
            "checked": check_item.checkState() == Qt.Checked if check_item else True,
        }

    def _restore_or_remove_row(self, row: int):
        previous = self._previous_editor_values.get(row)
        if previous and (previous["key"] or previous["code"]):
            self._set_code_row(row, previous["key"], previous["code"], "", previous["checked"], previous.get("cost"))
        elif 0 <= row < self.list_codes.rowCount():
            self.list_codes.removeRow(row)

    def _sync_click_through_from_win(self, checked: bool):
        """浮窗鼠标穿透状态变化（如快捷键触发）时同步设置窗口复选框"""
        self.cb_click_through.blockSignals(True)
        self.cb_click_through.setChecked(bool(checked))
        self.cb_click_through.blockSignals(False)

    def _on_click_through_toggled(self, checked: bool):
        self.win.set_click_through(bool(checked))

    def _on_force_top_toggled(self, checked: bool):
        self.win.set_force_top(bool(checked))

    def _on_hotkey_hide_enabled_toggled(self, checked: bool):
        self.keyseq_hide.setEnabled(bool(checked))
        result = self.win.set_hotkey_enabled(bool(checked))
        if not result:
            # 启用失败(如冲突):回滚复选框与输入框状态,并提示用户
            self.keyseq_hide.setEnabled(False)
            self.cb_hotkey_hide.blockSignals(True)
            self.cb_hotkey_hide.setChecked(False)
            self.cb_hotkey_hide.blockSignals(False)
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))

    def _on_click_through_hotkey_enabled_toggled(self, checked: bool):
        self.keyseq_click_through.setEnabled(bool(checked))
        result = self.win.set_click_through_hotkey_enabled(bool(checked))
        if not result:
            # 启用失败(如冲突):回滚复选框与输入框状态,并提示用户
            self.keyseq_click_through.setEnabled(False)
            self.cb_hotkey_click_through.blockSignals(True)
            self.cb_hotkey_click_through.setChecked(False)
            self.cb_hotkey_click_through.blockSignals(False)
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))

    def _on_click_through_hotkey_changed(self):
        new_hotkey = self.keyseq_click_through.keySequence().toString()
        result = self.win.update_click_through_hotkey(new_hotkey)
        if not result:
            # 冲突/无效:回滚输入框显示,并提示用户
            self.keyseq_click_through.setKeySequence(QKeySequence(self.win.hotkey_click_through))
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))

    def _setup_about(self):
        label = self.ui.label_about_info
        label.setWordWrap(True)
        has_update = bool(self.app is not None and getattr(self.app, "_has_update", False))
        first_line = f"当前版本 v{self.app.app_version}"
        if has_update:
            first_line += "（有新版本）"
        html = (
            f'<a href="{PROJECT_URL}" style="text-decoration:none; color:#4a90d9;">'
            f"{first_line}<br>Copyright 2026 sbr0574</a>"
        )
        label.setTextFormat(Qt.RichText)
        label.setText(html)
        label.setOpenExternalLinks(True)
        label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        label.setCursor(Qt.PointingHandCursor)

    def refresh_about(self):
        self._setup_about()

    def closeEvent(self, event):
        self._cleanup_code_rows()
        self._on_codes_changed(None)
        super().closeEvent(event)
