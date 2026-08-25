from functools import partial

from PySide6.QtCore import Qt, QPoint, QStringListModel, QEvent, QTimer
from PySide6.QtGui import QColor, QGuiApplication, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QWidget, QDialog, QColorDialog, QAbstractItemView, QTableWidgetItem,
    QStyledItemDelegate, QStyle, QStyleOptionViewItem, QLineEdit, QCompleter,
    QHeaderView, QButtonGroup, QMessageBox
)
from stockwidget.ui.generated.ui_settings import Ui_SettingDialog
from stockwidget.core.code_search import find_suggestions
from stockwidget.ui.widget import FloatLabel
from stockwidget.platform.capabilities import (
    hotkeys_supported,
    click_through_supported,
    opacity_supported,
    force_top_supported,
    start_on_boot_supported,
    unsupported_tooltip,
)
from stockwidget.data.update_check import GITHUB, project_links


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
    return "快捷键注册失败,请更换后重试。"


# 扁平化分组框列表（自选列表 gb_list 保持默认带边框样式，不在其中）
_FLAT_GROUPS = ("gb_data", "gb_data_setting", "gb_name", "gb_icon", "gb_fcn",
                "gb_color", "gb_text", "gb_tabel", "gb_hotkeys", "gb_about")


def _build_settings_stylesheet(dark: bool) -> str:
    """按系统深浅色生成设置窗口样式表（Qt QSS 不支持媒体查询，故在运行时按主题构建）。"""
    if dark:
        sep = "rgba(255, 255, 255, 0.35)"
        header_bg, header_line = "rgba(255, 255, 255, 0.10)", "rgba(255, 255, 255, 0.30)"
    else:
        sep = "rgba(0, 0, 0, 0.25)"
        header_bg, header_line = "rgba(0, 0, 0, 0.06)", "rgba(0, 0, 0, 0.20)"

    flat_boxes = ",\n".join(f"QGroupBox#{n}" for n in _FLAT_GROUPS)
    flat_titles = ",\n".join(f"QGroupBox#{n}::title" for n in _FLAT_GROUPS)

    return f"""
{flat_boxes} {{
    border: none;
    border-top: 1px solid {sep};
    margin-top: 9px;
    padding-top: 0px;
}}
{flat_titles} {{
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 2 8px;
}}
QTableWidget#list_codes QHeaderView::section {{
    background-color: {header_bg};
    border: none;
    border-bottom: 1px solid {header_line};
    padding: 4px 8px;
    font-weight: 600;
}}
"""


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


class CenteredCheckBoxDelegate(QStyledItemDelegate):
    """把自选列表第一列的勾选框在单元格内水平居中，同时保持可点击勾选。

    关键点：Qt 的勾选框“点击命中”走的是 delegate 的 editorEvent（视图在
    mouseReleaseEvent 里调用 edit() → sendDelegateEvent()），而不是 paint。
    因此必须同时重写 paint（居中绘制）和 editorEvent（按居中位置命中），
    否则会出现“画在中间、却要点左边才能切换”的问题。
    """

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


class SettingsDialog(QDialog):

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self.ui = Ui_SettingDialog()
        self.ui.setupUi(self)
        self.setModal(False)
        self._apply_theme_stylesheet()
        # 系统深浅色切换时跟随更新样式
        QGuiApplication.styleHints().colorSchemeChanged.connect(self._on_color_scheme_changed)
        self.suggestion_model = QStringListModel(self)
        self._suggestion_map = {}
        self._previous_editor_values = {}

        self._init_code_table()
        self._bind_widgets()
        self._load_settings()

    def _apply_theme_stylesheet(self):
        """按当前系统深浅色应用样式表。"""
        dark = QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
        self.setStyleSheet(_build_settings_stylesheet(dark))

    def _on_color_scheme_changed(self, _scheme=None):
        self._apply_theme_stylesheet()

    def _display_code_for_ui(self, code: str) -> str:
        return str(code or "").strip()

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
        self.list_codes.setItemDelegateForColumn(0, CenteredCheckBoxDelegate(self))
        self.list_codes.setItemDelegateForColumn(1, CodeCompleterDelegate(self))

        for code, entry in self.win.watchlist.items():
            checked = bool(entry.get("checked", True))
            cost = entry.get("cost")
            self._append_code_row(code, entry.get("name", ""), checked, cost)

    def _bind_widgets(self):
        self.sb_interval = self.ui.sb_interval
        self.cmb_source = self.ui.cmb_source
        self.label_data_state = self.ui.label_data_state
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

        self.cb_auto_start = self.ui.cb_auto_start
        self.cb_force_top = self.ui.cb_force_top
        self.cb_click_through = self.ui.cb_click_through
        self.cb_head = self.ui.cb_head
        self.cb_grid = self.ui.cb_grid

        self.cb_hotkey_hide = self.ui.cb_hotkey_hide
        self.cb_hotkey_click_through = self.ui.cb_hotkey_click_through
        self.keyseq_hide = self.ui.keyseq_hide
        self.keyseq_click_through = self.ui.keyseq_click_through

        self.sb_interval.valueChanged.connect(self._on_interval_changed)
        self.cmb_source.currentIndexChanged.connect(self._on_source_changed)
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
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)

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
        # 浮窗右键菜单等外部途径修改显示指标时，同步设置窗口复选框
        self.win.display_flags_changed.connect(self._sync_display_flags_from_win)
        self.cb_hotkey_hide.toggled.connect(self._on_hotkey_hide_enabled_toggled)
        self.cb_hotkey_click_through.toggled.connect(self._on_click_through_hotkey_enabled_toggled)
        self.cb_head.toggled.connect(self._on_header_toggled)
        self.cb_grid.toggled.connect(self._on_grid_toggled)

    def _load_settings(self):
        self.sb_interval.setValue(self.win.refresh_seconds)
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
        self._setup_source_combo()
        self._setup_about()
        self.refresh_data_state()

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
            # 鼠标穿透不可用（如 macOS）时，其快捷键一并关闭
            for w in (self.cb_hotkey_click_through, self.keyseq_click_through):
                w.setEnabled(False)
                w.setToolTip(unsupported_tooltip("鼠标穿透"))
        if not opacity_supported():
            # 整体不透明度滑块:Wayland 平台插件不支持设置窗口透明度
            for w in (self.slider_all_alpha, self.label_all, self.label_all_alpha):
                w.setEnabled(False)
            self.slider_all_alpha.setToolTip(unsupported_tooltip("整体不透明度"))
        if not force_top_supported():
            # 强制置顶:仅 Windows 支持
            self.cb_force_top.setEnabled(False)
            self.cb_force_top.setToolTip(unsupported_tooltip("强制置顶", suggest_x11=False))
        if not start_on_boot_supported():
            self.cb_auto_start.setEnabled(False)
            self.cb_auto_start.setToolTip(unsupported_tooltip("开机自启"))

    def _setup_icon_choices(self):
        self.icon_buttons = {
            "default": self.ui.btn_icon_default,
            "dark": self.ui.btn_icon_dark,
            "lightG": self.ui.btn_icon_lightG,
            "darkG": self.ui.btn_icon_darkG,
        }
        self._icon_button_group = QButtonGroup(self)
        self._icon_button_group.setExclusive(True)
        for key, btn in self.icon_buttons.items():
            self._icon_button_group.addButton(btn)
            btn.setCheckable(True)
            btn.setStyleSheet("QPushButton:checked { border: 2px solid #4a90d9; border-radius: 4px; }")
            btn.toggled.connect(partial(self._on_icon_button_toggled, key))

        cur_choice = self.app._icon_choice if self.app is not None else None
        if cur_choice not in self.icon_buttons:
            cur_choice = "default"
            if self.app is not None:
                self.app.set_app_icon(cur_choice)
                self.app.save_now()
        for btn in self.icon_buttons.values():
            btn.blockSignals(True)
        self.icon_buttons[cur_choice].setChecked(True)
        for btn in self.icon_buttons.values():
            btn.blockSignals(False)

    def _setup_source_combo(self):
        """填充行情数据源下拉框（新浪 / 东方财富），并按当前配置选中。"""
        self.cmb_source.blockSignals(True)
        self.cmb_source.clear()
        self.cmb_source.addItem("新浪", "sina")
        self.cmb_source.addItem("东方财富", "eastmoney")
        idx = self.cmb_source.findData(getattr(self.win, "data_source", "sina"))
        self.cmb_source.setCurrentIndex(idx if idx >= 0 else 0)
        self.cmb_source.blockSignals(False)

    def refresh_data_state(self):
        """更新“市场代码数据”状态文字：在线/缓存/离线 + 更新日期(yyyymmdd)。"""
        if self.app is not None and hasattr(self.app, "code_data_state"):
            state, date = self.app.code_data_state()
        else:
            state, date = "offline", ""
        d = str(date or "").replace("-", "")
        if state == "online":
            text = f"✅ 市场代码数据：在线 ({d})"
        elif state == "cached":
            text = f"⚠️ 市场代码数据：缓存 ({d})"
        else:
            text = f"❌ 市场代码数据：离线 ({d})"
        self.label_data_state.setText(text)

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
        """拖动调整顺序：将拖动的行移动到目标位置。
        1. 用 CopyAction（而非 MoveAction）结束拖放，让 drag->exec() 不返回
           MoveAction，从而不触发 startDrag() 里的 clearOrRemove()；
        2. 同时清空选中——即使 clearOrRemove() 仍被触发，也没有选中行可删。
        否则源行会被二次删除，表现为拖拽后丢一行。
        """
        src_row = self.list_codes.currentRow()
        pos = ev.position().toPoint()
        target_row = self.list_codes.rowAt(pos.y())
        if ev.source() is self.list_codes and src_row >= 0:
            if target_row < 0:
                target_row = self.list_codes.rowCount() - 1
            if target_row != src_row:
                self._move_row(src_row, target_row)
        self.list_codes.clearSelection()
        ev.accept()
        ev.setDropAction(Qt.CopyAction)
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
            check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
            self.list_codes.setItem(row, 0, check_item)
        check_item.setCheckState(Qt.Checked if checked else Qt.Unchecked)

        code_item = self.list_codes.item(row, 1)
        if code_item is None:
            code_item = QTableWidgetItem("")
            code_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
            self.list_codes.setItem(row, 1, code_item)
        code_item.setText(self._code_display_for_row(value_key, display_code, resolved_name))
        code_item.setData(Qt.UserRole, value_key)

        cost_item = self.list_codes.item(row, 2)
        if cost_item is None:
            cost_item = QTableWidgetItem("")
            self.list_codes.setItem(row, 2, cost_item)
        if self._row_type(value_key, display_code) == "指":
            # 指数不允许设置成本
            cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
            cost_item.setText("")
            cost_item.setData(Qt.UserRole, None)
        else:
            cost_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
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

    def _on_interval_changed(self, value: int):
        self.win.set_refresh_interval(value)

    def _on_source_changed(self, idx: int):
        src = self.cmb_source.itemData(idx)
        if src:
            self.win.set_data_source(str(src))

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

    def _on_icon_button_toggled(self, key: str, checked: bool):
        if not checked:
            return
        if self.app is not None:
            self.app.set_app_icon(key)
            self.app.save_now()

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
        popup_width = max(self.list_codes.columnWidth(1), editor.width())
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

    @staticmethod
    def _set_checked_blocked(widget, checked: bool):
        """设置可勾选控件的状态，并屏蔽其 toggled 信号，避免反向触发浮窗改动。"""
        widget.blockSignals(True)
        widget.setChecked(bool(checked))
        widget.blockSignals(False)

    def _sync_display_flags_from_win(self):
        """浮窗右键菜单等外部途径修改显示指标时，同步设置窗口对应复选框。"""
        for cb, header in (
            (self.cb_price, "现价"),
            (self.cb_diff, "涨跌"),
            (self.cb_pct, "涨幅"),
            (self.cb_vol, "成交量"),
            (self.cb_amount, "成交额"),
            (self.cb_avg, "均价"),
            (self.cb_commi, "委比"),
            (self.cb_kline, "K线"),
            (self.cb_profit, "浮盈"),
        ):
            self._set_checked_blocked(cb, self.win.header_is_visible(header))
        self._set_checked_blocked(self.cb_b1s1, self.win.b1s1_visible)
        self._set_checked_blocked(self.gb_name, self.win.name_visible)
        self._set_checked_blocked(self.cb_type, self.win.type_visible)
        self._set_checked_blocked(self.cb_code, self.win.code_visible)
        self._set_checked_blocked(self.cb_head, self.win.header_visible)
        self._set_checked_blocked(self.cb_grid, self.win.grid_visible)
        self._set_checked_blocked(self.cb_default_color, self.win.default_color)
        # 默认颜色开启时禁用前景色选择按钮
        self.btn_fg.setEnabled(not self.win.default_color)

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
        app_version = self.app.app_version if self.app is not None else "1.0.0"
        has_update = bool(self.app is not None and getattr(self.app, "_has_update", False))
        latest_version = getattr(self.app, "_latest_version", None) if self.app is not None else None
        latest_url = getattr(self.app, "_latest_release_url", None) if self.app is not None else None
        if not has_update:
            latest_version = None
        source = getattr(self.app, "_remote_source", GITHUB) if self.app is not None else GITHUB
        links = project_links(source)
        latest_url = latest_url or links["releases"]

        version_line = f"当前版本 v{app_version}"
        if latest_version:
            version_line += f" 最新 v{latest_version}"
        html = (
            f'<p style="margin:2px 0;"><a href="{latest_url}" style="text-decoration:none; color:#4a90d9;">{version_line}</a></p>'
            f'<p style="margin:2px 0;"><a href="{links["license"]}" style="text-decoration:none; color:#4a90d9;">License</a> · '
            f'<a href="{links["readme"]}" style="text-decoration:none; color:#4a90d9;">使用帮助</a> · '
            f'<a href="{links["issues"]}" style="text-decoration:none; color:#4a90d9;">反馈建议</a></p>'
            f'<p style="margin:2px 0;"><a href="{links["project"]}" style="text-decoration:none; color:#4a90d9;">{links["repository_label"]}</a></p>'
            f'<p style="margin:2px 0;">Copyright 2026 sbr0574</p>'
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
