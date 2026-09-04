import sys
import threading
from functools import partial
from math import isfinite

from PySide6.QtCore import (
    Qt, QPoint, QEvent, QTimer, QItemSelectionModel, QModelIndex, Signal,
)
from PySide6.QtGui import (
    QColor, QGuiApplication, QIcon, QKeySequence, QPainter, QPixmap,
    QStandardItem, QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QDialog, QColorDialog, QAbstractItemView, QTableWidgetItem,
    QAbstractItemDelegate, QStyledItemDelegate, QStyle, QStyleOptionViewItem,
    QLineEdit, QCompleter, QHeaderView, QButtonGroup, QMessageBox, QLabel,
    QVBoxLayout,
)
from stockwidget.ui.generated.ui_settings import Ui_SettingDialog
from stockwidget.core.code_search import build_search_index, search_suggestions
from stockwidget.ui.widget import FloatLabel
from stockwidget.ui.metric_pool import MetricPoolWidget
from stockwidget.ui.add_code_panel import (
    ADDED_ROLE, ENTRY_ROLE, AddCodePanel, entry_display_text,
)
from stockwidget.platform.capabilities import (
    hotkeys_supported,
    click_through_supported,
    opacity_supported,
    force_top_supported,
    start_on_boot_supported,
    unsupported_tooltip,
)
from stockwidget.data.update_check import github_available, project_links


def _parse_positive_cost(value):
    """把成本输入规范化为有限正数；空值或无效值返回 None。"""
    if value in (None, ""):
        return None
    try:
        cost = float(value)
    except (TypeError, ValueError):
        return None
    return cost if isfinite(cost) and cost > 0 else None


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
_FLAT_GROUPS = ("gb_data", "gb_data_setting", "gb_name", "gb_icon", "gb_fcn", "gb_opacity",
                "gb_color", "gb_text", "gb_tabel", "gb_hotkeys", "gb_about")

_SEARCH_PLACEHOLDER = "搜索代码、名称、拼音或缩写，空格区分关键词"
_COLOR_SWATCH_SIZE = 12
_UNICOLOR_EXTRA_WIDTH = 8


def _color_swatch_icon(color: QColor, device_pixel_ratio: float = 1.0) -> QIcon:
    """按屏幕像素比绘制无描边圆形色标，避免高 DPI 缩放发糊。"""
    ratio = max(1.0, float(device_pixel_ratio))
    pixel_size = max(_COLOR_SWATCH_SIZE, round(_COLOR_SWATCH_SIZE * ratio))
    ratio = pixel_size / _COLOR_SWATCH_SIZE
    pixmap = QPixmap(pixel_size, pixel_size)
    pixmap.setDevicePixelRatio(ratio)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(color)
    painter.drawEllipse(1, 1, _COLOR_SWATCH_SIZE - 2, _COLOR_SWATCH_SIZE - 2)
    painter.end()

    icon = QIcon()
    icon.addPixmap(pixmap, QIcon.Mode.Normal, QIcon.State.Off)
    icon.addPixmap(pixmap, QIcon.Mode.Disabled, QIcon.State.Off)
    return icon


def _build_settings_stylesheet(dark: bool, *, macos: bool = False) -> str:
    """按系统深浅色生成设置窗口样式表（Qt QSS 不支持媒体查询，故在运行时按主题构建）。"""
    if dark:
        sep = "rgba(255, 255, 255, 0.35)"
        header_bg, header_line = "rgba(255, 255, 255, 0.10)", "rgba(255, 255, 255, 0.30)"
        empty_hint = "rgba(255, 255, 255, 0.28)"
    else:
        sep = "rgba(0, 0, 0, 0.25)"
        header_bg, header_line = "rgba(0, 0, 0, 0.06)", "rgba(0, 0, 0, 0.20)"
        empty_hint = "rgba(0, 0, 0, 0.28)"

    flat_boxes = ",\n".join(f"QGroupBox#{n}" for n in _FLAT_GROUPS)
    flat_titles = ",\n".join(f"QGroupBox#{n}::title" for n in _FLAT_GROUPS)

    macos_buttons = ""
    if macos:
        if dark:
            button_bg = "rgba(255, 255, 255, 0.10)"
            button_hover = "rgba(255, 255, 255, 0.16)"
            button_pressed = "rgba(255, 255, 255, 0.22)"
            button_border = "rgba(255, 255, 255, 0.24)"
            button_disabled = "rgba(255, 255, 255, 0.05)"
            icon_selected = "rgba(10, 132, 255, 0.24)"
        else:
            button_bg = "rgba(255, 255, 255, 0.86)"
            button_hover = "rgba(255, 255, 255, 1.00)"
            button_pressed = "rgba(0, 0, 0, 0.08)"
            button_border = "rgba(0, 0, 0, 0.22)"
            button_disabled = "rgba(0, 0, 0, 0.04)"
            icon_selected = "rgba(0, 122, 255, 0.14)"

        regular_selectors = (
            "QPushButton#btn_add",
            "QPushButton#btn_del",
        )
        icon_selectors = (
            "QPushButton#btn_icon_default",
            "QPushButton#btn_icon_lightG",
            "QPushButton#btn_icon_dark",
            "QPushButton#btn_icon_darkG",
        )
        color_selectors = (
            "QPushButton#btn_fg_color",
            "QPushButton#btn_bg_color",
            "QPushButton#btn_up_color",
            "QPushButton#btn_down_color",
            "QPushButton#btn_neutral_color",
        )
        flat_selectors = icon_selectors + color_selectors
        regular_buttons = ",\n".join(regular_selectors)
        regular_hover = ",\n".join(f"{selector}:hover" for selector in regular_selectors)
        regular_pressed = ",\n".join(f"{selector}:pressed" for selector in regular_selectors)
        regular_focus = ",\n".join(f"{selector}:focus" for selector in regular_selectors)
        regular_disabled = ",\n".join(f"{selector}:disabled" for selector in regular_selectors)
        flat_buttons = ",\n".join(flat_selectors)
        flat_hover = ",\n".join(f"{selector}:hover" for selector in flat_selectors)
        flat_pressed = ",\n".join(f"{selector}:pressed" for selector in flat_selectors)
        icon_checked = ",\n".join(f"{selector}:checked" for selector in icon_selectors)
        macos_buttons = f"""
{regular_buttons} {{
    background-color: {button_bg};
    border: 1px solid {button_border};
    border-radius: 6px;
    padding: 3px 9px;
}}
{regular_hover} {{
    background-color: {button_hover};
}}
{regular_pressed} {{
    background-color: {button_pressed};
}}
{regular_focus} {{
    border: 2px solid rgba(10, 132, 255, 0.82);
}}
{regular_disabled} {{
    background-color: {button_disabled};
    border-color: transparent;
}}
{flat_buttons} {{
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 3px;
}}
{flat_hover} {{
    background-color: {button_hover};
}}
{flat_pressed} {{
    background-color: {button_pressed};
}}
{icon_checked} {{
    background-color: {icon_selected};
    border: 2px solid rgb(10, 132, 255);
}}
"""

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
QLabel#empty_watchlist_hint {{
    color: {empty_hint};
    background: transparent;
    font-size: 18px;
    font-weight: 500;
}}
{macos_buttons}
"""


class CodeSearchEditor(QLineEdit):
    """自选列表的全范围快速搜索输入框。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText(_SEARCH_PLACEHOLDER)
        self.setToolTip(_SEARCH_PLACEHOLDER)


class CodeCompleterDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def createEditor(self, parent, option, index):
        self.owner._ensure_search_index()
        editor = CodeSearchEditor(parent)
        editor.setProperty("_row", index.row())
        editor.setProperty("_column", index.column())
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
        if not editor.property("_code_editor_committed"):
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
        if editor.property("_code_editor_committed"):
            return
        editor.setProperty("_code_editor_committed", True)
        self.owner._commit_code_editor(editor)


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
    github_check_finished = Signal(bool)

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self._use_gitee_links = False
        self.ui = Ui_SettingDialog()
        self.ui.setupUi(self)
        self._init_metric_pool()
        self.setModal(False)
        self._apply_theme_stylesheet()
        # 系统深浅色切换时跟随更新样式
        QGuiApplication.styleHints().colorSchemeChanged.connect(self._on_color_scheme_changed)
        self._search_index_source = None
        self._search_index_size = -1
        self._search_index = ()
        self._search_entries_by_key = {}
        self.suggestion_model = QStandardItemModel(self)

        self._init_code_table()
        self._bind_widgets()
        self._load_settings()
        self.github_check_finished.connect(self._on_github_check_finished)
        self._start_github_check()

    def _init_metric_pool(self):
        """用动态双池替换固定指标复选框区域。"""
        layout = QVBoxLayout(self.ui.gb_data)
        layout.setContentsMargins(5, 20, 5, 5)
        layout.setSpacing(0)
        self.metric_pool = MetricPoolWidget(self.ui.gb_data)
        layout.addWidget(self.metric_pool)

    def _start_github_check(self):
        """后台选择关于页链接平台，不阻塞设置窗口构造。"""

        def _worker():
            use_gitee = not github_available(timeout=2)
            try:
                self.github_check_finished.emit(use_gitee)
            except RuntimeError:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_github_check_finished(self, use_gitee: bool):
        if self._use_gitee_links == use_gitee:
            return
        self._use_gitee_links = use_gitee
        self._setup_about()

    def _apply_theme_stylesheet(self):
        """按当前系统深浅色应用样式表。"""
        dark = QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
        self.setStyleSheet(_build_settings_stylesheet(dark, macos=sys.platform == "darwin"))
        if hasattr(self, "metric_pool"):
            self.metric_pool.set_theme(dark)
        if hasattr(self, "add_code_panel"):
            self.add_code_panel.set_theme(dark)
        self._refresh_color_buttons()

    def _refresh_color_buttons(self):
        """用无描边圆形图标展示五个颜色按钮的当前色值。"""
        if not hasattr(self, "_color_buttons"):
            return

        for button, attr, title in self._color_buttons:
            color = QColor(getattr(self.win, attr))
            color_name = color.name(QColor.NameFormat.HexRgb)
            button.setToolTip(f"{title}: {color_name}")
            button.setIcon(_color_swatch_icon(color, button.devicePixelRatioF()))

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

        self.empty_watchlist_hint = QLabel(
            "双击空白处添加条目", self.list_codes.viewport()
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

        for code, entry in self.win.watchlist.items():
            checked = bool(entry.get("checked", True))
            cost = entry.get("cost")
            self._append_code_row(code, entry.get("name", ""), checked, cost)
        self._refresh_empty_watchlist_hint()

    def _bind_widgets(self):
        self.sb_interval = self.ui.sb_interval
        self.rb_sina = self.ui.rb_sina
        self.rb_em = self.ui.rb_em
        self._source_buttons = {
            "sina": self.rb_sina,
            "eastmoney": self.rb_em,
        }
        self.label_data_state = self.ui.label_data_state
        self.gb_name = self.ui.gb_name
        self.cb_code = self.ui.cb_code
        self.cb_type = self.ui.cb_type
        self.cmb_namelen = self.ui.cmb_namelen

        self.cb_unicolor = self.ui.cb_unicolor
        # macOS 原生复选框的 sizeHint 较紧，额外留出字形右侧空间避免裁切。
        self.cb_unicolor.setMinimumWidth(
            self.cb_unicolor.sizeHint().width() + _UNICOLOR_EXTRA_WIDTH
        )
        self.btn_fg = self.ui.btn_fg_color
        self.btn_bg = self.ui.btn_bg_color
        self.btn_up = self.ui.btn_up_color
        self.btn_down = self.ui.btn_down_color
        self.btn_neutral = self.ui.btn_neutral_color
        self._color_buttons = (
            (self.btn_bg, "bg", "背景颜色"),
            (self.btn_fg, "fg", "文字颜色"),
            (self.btn_up, "up_color", "上涨颜色"),
            (self.btn_down, "down_color", "下跌颜色"),
            (self.btn_neutral, "neutral_color", "中性颜色"),
        )
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
        for source, button in self._source_buttons.items():
            button.toggled.connect(partial(self._on_source_toggled, source))
        self.list_codes.itemChanged.connect(self._on_codes_changed)

        self.gb_name.toggled.connect(self._on_name_toggled)
        self.cb_code.toggled.connect(self._on_code_toggled)
        self.cb_type.toggled.connect(self._on_type_toggled)
        self.metric_pool.visible_metrics_changed.connect(
            self.win.set_visible_metrics
        )

        self.btn_add = self.ui.btn_add
        self.btn_del = self.ui.btn_del
        self.add_code_panel = AddCodePanel(_SEARCH_PLACEHOLDER, self)
        self.add_code_panel.set_theme(
            QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
        )
        self.add_code_panel.entry_requested.connect(self._add_entry_from_panel)
        self.btn_add.clicked.connect(self._show_add_code_panel)
        self.btn_del.clicked.connect(self._del_code)

        self.cmb_namelen.currentIndexChanged.connect(self._on_name_length_changed)
        self.cb_unicolor.toggled.connect(self._on_unicolor_toggled)
        self.btn_fg.clicked.connect(self.pick_fg)
        self.btn_bg.clicked.connect(self.pick_bg)
        self.btn_up.clicked.connect(self.pick_up)
        self.btn_down.clicked.connect(self.pick_down)
        self.btn_neutral.clicked.connect(self.pick_neutral)
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
        self.metric_pool.set_visible_metrics(self.win.visible_metrics)

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

        self._set_checked_blocked(self.cb_unicolor, self.win.unicolor)
        self._update_direction_color_controls()
        self._refresh_color_buttons()
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
        self._setup_source_buttons()
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
            btn.setFlat(sys.platform == "darwin")
            if sys.platform != "darwin":
                btn.setStyleSheet(
                    "QPushButton:checked { border: 2px solid #4a90d9; border-radius: 4px; }"
                )
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

    def _setup_source_buttons(self):
        """按当前配置选中行情数据源单选按钮。"""
        source = getattr(self.win, "data_source", "sina")
        if source not in self._source_buttons:
            source = "sina"
        for button in self._source_buttons.values():
            button.blockSignals(True)
        self._source_buttons[source].setChecked(True)
        for button in self._source_buttons.values():
            button.blockSignals(False)

    def refresh_data_state(self):
        """更新市场代码状态；qrc 与本地文件都属于缓存。"""
        if self.app is not None and hasattr(self.app, "code_data_state"):
            state, date = self.app.code_data_state()
        else:
            state, date = "cached", ""
        d = str(date or "").replace("-", "")
        if state == "current":
            text = f"✅ 市场代码数据：最新 ({d})"
        else:
            text = f"⚠️ 市场代码数据：缓存 ({d})"
        self.label_data_state.setText(text)

    def refresh_code_search(self):
        """代码表异步替换后刷新快速搜索和添加面板。"""
        self._search_index_source = None
        self._search_index_size = -1
        self._ensure_search_index()
        for editor in self.list_codes.findChildren(CodeSearchEditor):
            if editor.isVisible():
                self._update_suggestions(editor, editor.text())
        if hasattr(self, "add_code_panel") and self.add_code_panel.isVisible():
            self.add_code_panel.set_context(
                self._search_index, self._existing_code_keys()
            )

    def eventFilter(self, obj, ev):
        if obj is self.list_codes.viewport() and ev.type() == QEvent.Resize:
            self._refresh_empty_watchlist_hint()
        if obj is self.list_codes.viewport() and ev.type() == QEvent.MouseButtonDblClick:
            pos = ev.position().toPoint() if hasattr(ev, 'position') else ev.pos()
            if self.list_codes.itemAt(pos) is None:
                self._start_quick_add()
                return True
        if obj is self.list_codes.viewport() and ev.type() == QEvent.Drop:
            self._handle_drop(ev)
            return True
        return super().eventFilter(obj, ev)

    def _refresh_empty_watchlist_hint(self, *_args):
        if not hasattr(self, "empty_watchlist_hint"):
            return
        self.empty_watchlist_hint.setGeometry(self.list_codes.viewport().rect())
        self.empty_watchlist_hint.setVisible(self.list_codes.rowCount() == 0)
        self.empty_watchlist_hint.raise_()

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
        self._insert_code_row(
            self.list_codes.rowCount(), code, name, checked, cost
        )

    def _insert_code_row(self, row: int, code: str = "", name: str = "",
                         checked: bool = False, cost=None):
        self.list_codes.blockSignals(True)
        row = max(0, min(int(row), self.list_codes.rowCount()))
        self.list_codes.insertRow(row)
        self._set_code_row(row, code, code, name, checked, cost)
        self.list_codes.blockSignals(False)

    def _ensure_search_index(self):
        """代码表对象变化时重建索引；普通查询复用已规范化记录。"""
        codes = self.win.codes_list if isinstance(self.win.codes_list, dict) else {}
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
        cost = _parse_positive_cost(cost)
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
                entry["cost"] = _parse_positive_cost(cost_item.text())
            if resolved:
                entry["name"] = str(resolved.get("name", "") or "")
            watchlist[value] = entry
        return watchlist

    def _on_codes_changed(self, _item):
        self._cleanup_code_rows()
        watchlist = self._collect_watchlist_from_list()
        self.win.set_watchlist(watchlist)
        self._refresh_empty_watchlist_hint()
        if hasattr(self, "add_code_panel") and self.add_code_panel.isVisible():
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

    def _on_source_toggled(self, source: str, checked: bool):
        if checked:
            self.win.set_data_source(source)

    def _on_code_toggled(self, checked: bool):
        self.win.set_code_visible(checked)

    def _on_name_toggled(self, checked: bool):
        self.win.set_flag("名称", checked)

    def _on_type_toggled(self, checked: bool):
        self.win.set_type_visible(checked)

    def _update_direction_color_controls(self):
        enabled = not self.win.unicolor
        for widget in (
            self.btn_up,
            self.btn_down,
            self.btn_neutral,
        ):
            widget.setEnabled(enabled)

    def _on_unicolor_toggled(self, checked: bool):
        self.win.set_unicolor(bool(checked))
        self._update_direction_color_controls()
        self._refresh_color_buttons()

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
            self._refresh_color_buttons()

    def pick_bg(self):
        base = QColor(self.win.bg)
        base.setAlpha(255)
        c = QColorDialog.getColor(base, self, "选择背景颜色")
        if c.isValid():
            self.win.set_bg_rgb_keep_alpha(c)
            self._refresh_color_buttons()

    def pick_up(self):
        c = QColorDialog.getColor(self.win.up_color, self, "选择上涨颜色")
        if c.isValid():
            self.win.set_up_color(c)
            self._refresh_color_buttons()

    def pick_down(self):
        c = QColorDialog.getColor(self.win.down_color, self, "选择下跌颜色")
        if c.isValid():
            self.win.set_down_color(c)
            self._refresh_color_buttons()

    def pick_neutral(self):
        c = QColorDialog.getColor(self.win.neutral_color, self, "选择中性颜色")
        if c.isValid():
            self.win.set_neutral_color(c)
            self._refresh_color_buttons()

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
        cost = _parse_positive_cost(cost_item.text()) if cost_item is not None else None
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
            "cost": _parse_positive_cost(cost_item.text()) if cost_item else None,
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
        """浮窗右键菜单等外部途径修改显示状态时，同步指标池。"""
        self.metric_pool.set_visible_metrics(self.win.visible_metrics)
        self._set_checked_blocked(self.gb_name, self.win.name_visible)
        self._set_checked_blocked(self.cb_type, self.win.type_visible)
        self._set_checked_blocked(self.cb_code, self.win.code_visible)
        self._set_checked_blocked(self.cb_head, self.win.header_visible)
        self._set_checked_blocked(self.cb_grid, self.win.grid_visible)
        self._set_checked_blocked(self.cb_unicolor, self.win.unicolor)
        self._update_direction_color_controls()
        self._refresh_color_buttons()

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
        if not has_update:
            latest_version = None
        links = project_links(use_gitee=self._use_gitee_links)
        github_links = project_links()
        gitee_links = project_links(use_gitee=True)

        version_line = f"当前版本 v{app_version}"
        if latest_version:
            version_line += f" 最新 v{latest_version}"
        html = (
            f'<p style="margin:2px 0;"><a href="{links["releases"]}" style="text-decoration:none; color:#4a90d9;">{version_line}</a></p>'
            f'<p style="margin:2px 0;"><a href="{links["license"]}" style="text-decoration:none; color:#4a90d9;">License</a> · '
            f'<a href="{links["readme"]}" style="text-decoration:none; color:#4a90d9;">使用帮助</a> · '
            f'<a href="{links["issues"]}" style="text-decoration:none; color:#4a90d9;">反馈建议</a></p>'
            f'<p style="margin:2px 0;"><a href="{github_links["project"]}" style="text-decoration:none; color:#4a90d9;">GitHub仓库</a> · '
            f'<a href="{gitee_links["project"]}" style="text-decoration:none; color:#4a90d9;">Gitee仓库</a></p>'
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
        if hasattr(self, "add_code_panel"):
            self.add_code_panel.hide()
        self._cleanup_code_rows()
        self._on_codes_changed(None)
        super().closeEvent(event)
