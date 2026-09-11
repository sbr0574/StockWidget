import os
import sys
import threading
from contextlib import ExitStack
from functools import partial
from html import escape

from PySide6.QtCore import (
    Qt, QEvent, QUrl, QSignalBlocker, Signal,
)
from PySide6.QtGui import (
    QColor, QDesktopServices, QGuiApplication, QIcon, QKeySequence,
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QDialog, QColorDialog, QButtonGroup, QFileDialog, QMessageBox,
    QScrollBar, QVBoxLayout,
)
from stockwidget.ui.generated.ui_settings import Ui_SettingDialog
from stockwidget.constants import APP_VERSION
from stockwidget.core.config_store import config_paths
from stockwidget.ui.widget import FloatLabel
from stockwidget.ui.metric_pool import MetricPoolWidget
from stockwidget.ui.name_settings_panel import NameSettingsPanel
from stockwidget.ui.unit_settings_panel import UnitSettingsPanel
from stockwidget.ui.watchlist_editor import WatchlistEditor
from stockwidget.ui.settings_style import build_settings_stylesheet, color_swatch_icon
from stockwidget.platform.capabilities import (
    hotkeys_supported,
    click_through_supported,
    opacity_supported,
    force_top_supported,
    start_on_boot_supported,
    unsupported_tooltip,
)
from stockwidget.data.update_check import (
    get_update_info,
    github_available,
    project_links,
)


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


_ICON_FILE_FILTER = (
    "图标文件 (*.png *.ico *.icns *.svg *.jpg *.jpeg *.bmp *.gif *.webp);;"
    "所有文件 (*)"
)


class SettingsDialog(QDialog):
    github_check_finished = Signal(bool)
    update_check_finished = Signal(object)

    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.win = win
        self.app = app
        self._use_gitee_links = False
        self.ui = Ui_SettingDialog()
        self.ui.setupUi(self)
        # Linux 下用 Tool 窗口避开任务栏/程序坞条目；
        # macOS 的 Dock 图标由应用级 Accessory 激活策略隐藏
        # （见 app._hide_macos_dock_icon），窗口保持普通标题栏，
        # 避免 Tool 窗口的小号红黄绿按钮与小标题。
        if sys.platform == "linux":
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.Tool)
        self._init_metric_pool()
        self.setModal(False)
        self.watchlist_editor = WatchlistEditor(
            self.ui.list_codes, self.ui.btn_add, self.ui.btn_del, self.ui.btn_top,
            self.win.watchlist, self.win.codes_list, self,
        )
        self.watchlist_editor.watchlist_changed.connect(self.win.set_watchlist)
        self._bind_widgets()
        self._load_settings()
        self._apply_theme_stylesheet()
        QGuiApplication.styleHints().colorSchemeChanged.connect(self._apply_theme_stylesheet)
        QApplication.instance().paletteChanged.connect(self._apply_theme_stylesheet)
        self.github_check_finished.connect(self._on_github_check_finished)
        self.update_check_finished.connect(self._on_update_check_finished)
        self._start_github_check()
        # 全局监听鼠标按下：点击设置页空白/其他区域时清除自选列表与指标池的选中
        QApplication.instance().installEventFilter(self)

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

    def _apply_theme_stylesheet(self, *_args):
        dark = QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
        self.setStyleSheet(build_settings_stylesheet(dark))
        for component in (self.metric_pool, self.watchlist_editor,
                          self.name_settings_panel, self.unit_settings_panel):
            component.set_theme(dark)
        self._refresh_color_buttons()

    def _refresh_color_buttons(self):
        """用无描边圆形图标展示五个颜色按钮的当前色值。"""
        for button, attr, title in self._color_buttons:
            color = QColor(getattr(self.win, attr))
            color_name = color.name(QColor.NameFormat.HexRgb)
            button.setToolTip(f"{title}: {color_name}")
            button.setIcon(color_swatch_icon(color, button.devicePixelRatioF()))

    def _bind_widgets(self):
        self.sb_interval = self.ui.sb_interval
        self.rb_sina = self.ui.rb_sina
        self.rb_em = self.ui.rb_em
        self._source_buttons = {
            "sina": self.rb_sina,
            "eastmoney": self.rb_em,
        }
        self.label_data_state = self.ui.label_data_state

        self.cb_unicolor = self.ui.cb_unicolor
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

        self.sb_interval.valueChanged.connect(self.win.set_refresh_interval)
        for source, button in self._source_buttons.items():
            button.toggled.connect(partial(self._on_source_toggled, source))

        self.metric_pool.visible_metrics_changed.connect(
            self.win.set_visible_metrics
        )
        self.metric_pool.name_settings_requested.connect(
            self._show_name_settings_panel
        )
        self.metric_pool.unit_settings_requested.connect(
            self._show_unit_settings_panel
        )
        self.name_settings_panel = NameSettingsPanel(self)
        self.name_settings_panel.name_length_changed.connect(
            self.win.set_name_length
        )
        self.name_settings_panel.code_visible_changed.connect(
            self.win.set_code_visible
        )
        self.name_settings_panel.type_visible_changed.connect(
            self.win.set_type_visible
        )
        self.unit_settings_panel = UnitSettingsPanel(self)
        self.unit_settings_panel.unit_mode_changed.connect(
            self.win.set_unit_mode
        )
        # 面板关闭（含点击外部空白关闭）后清除指标块的选中状态
        self.name_settings_panel.panel_closed.connect(
            self.metric_pool.clear_selections
        )
        self.unit_settings_panel.panel_closed.connect(
            self.metric_pool.clear_selections
        )

        self.btn_check_update = self.ui.btn_check_update
        self.btn_open_cache_dir = self.ui.btn_open_cache_dir
        self.btn_check_update.clicked.connect(self._check_update_manually)
        self.btn_open_cache_dir.clicked.connect(self._open_cache_dir)

        self.cb_unicolor.toggled.connect(self._on_unicolor_toggled)
        color_setters = {
            "bg": self.win.set_bg_rgb_keep_alpha,
            "fg": self.win.set_fg_color,
            "up_color": self.win.set_up_color,
            "down_color": self.win.set_down_color,
            "neutral_color": self.win.set_neutral_color,
        }
        for button, attr, title in self._color_buttons:
            button.clicked.connect(partial(self._pick_color, attr, title, color_setters[attr]))
        self.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.slider_all_alpha.valueChanged.connect(self.apply_win_opacity)

        self.cmb_family.currentTextChanged.connect(self.win.set_font_family)
        self.slider_font.valueChanged.connect(self.apply_font_size)
        self.slider_line.valueChanged.connect(self._on_line_changed)
        self.keyseq_hide.editingFinished.connect(self._on_hotkey_changed)
        self.keyseq_click_through.editingFinished.connect(self._on_click_through_hotkey_changed)
        self.cb_auto_start.toggled.connect(self._on_start_on_boot_toggled)
        self.cb_force_top.toggled.connect(self.win.set_force_top)
        self.cb_click_through.toggled.connect(self.win.set_click_through)
        self.win.click_through_changed.connect(self._sync_click_through_from_win)
        # 浮窗右键菜单等外部途径修改显示指标时，同步设置窗口复选框
        self.win.display_flags_changed.connect(self._sync_display_flags_from_win)
        self.cb_hotkey_hide.toggled.connect(self._on_hotkey_hide_enabled_toggled)
        self.cb_hotkey_click_through.toggled.connect(self._on_click_through_hotkey_enabled_toggled)
        self.cb_head.toggled.connect(self.win.set_header_visible)
        self.cb_grid.toggled.connect(self.win.set_grid_visible)

    def _load_settings(self):
        self.sb_interval.setValue(self.win.refresh_seconds)
        self.metric_pool.set_visible_metrics(self.win.visible_metrics)
        self.name_settings_panel.sync_from(self.win)
        self.unit_settings_panel.sync_from(self.win)

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
        self.slider_line.setValue(self.win.line_extra_px)
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
            "custom": self.ui.btn_icon_custom,
        }
        custom_path = getattr(self.app, "_custom_icon_path", "") if self.app else ""
        custom_icon = QIcon(custom_path) if custom_path else QIcon()
        self.ui.btn_icon_custom.set_custom_icon(custom_icon)
        self._update_custom_icon_tooltip(custom_path if not custom_icon.isNull() else "")
        self.ui.btn_icon_custom.deleteRequested.connect(self._clear_custom_icon)

        self._icon_button_group = QButtonGroup(self)
        self._icon_button_group.setExclusive(True)
        for key, btn in self.icon_buttons.items():
            self._icon_button_group.addButton(btn)
            btn.setCheckable(True)
            btn.toggled.connect(partial(self._on_icon_button_toggled, key))

        cur_choice = self.app._icon_choice if self.app is not None else None
        if cur_choice == "custom" and not self.ui.btn_icon_custom.has_custom_icon():
            cur_choice = "default"
        if cur_choice not in self.icon_buttons:
            cur_choice = "default"
            if self.app is not None:
                self.app.set_app_icon(cur_choice)
                self.app.save_now()
        self._set_checked_icon(cur_choice)
        self._active_icon_choice = cur_choice

    def _set_checked_icon(self, choice: str):
        self._select_button(self.icon_buttons, choice)

    @staticmethod
    def _select_button(buttons, choice):
        with ExitStack() as stack:
            for button in buttons.values():
                stack.enter_context(QSignalBlocker(button))
            buttons[choice].setChecked(True)

    def _update_custom_icon_tooltip(self, path: str):
        if path:
            self.ui.btn_icon_custom.setToolTip(
                f"自定义图标：{path}\n悬停右上角可删除"
            )
        else:
            self.ui.btn_icon_custom.setToolTip("选择自定义图标")

    def _choose_custom_icon(self):
        previous = self._active_icon_choice
        initial_path = getattr(self.app, "_custom_icon_path", "") if self.app else ""
        path, _selected_filter = QFileDialog.getOpenFileName(
            self, "选择自定义图标", initial_path, _ICON_FILE_FILTER
        )
        if not path:
            self._set_checked_icon(previous)
            return

        if self.app is None or not self.app.set_custom_icon(path):
            QMessageBox.warning(self, "图标无效", "无法读取所选图标文件，请选择其他文件。")
            self._set_checked_icon(previous)
            return

        custom_path = self.app._custom_icon_path
        self.ui.btn_icon_custom.set_custom_icon(QIcon(custom_path))
        self._update_custom_icon_tooltip(custom_path)
        self._active_icon_choice = "custom"
        self._set_checked_icon("custom")
        self.app.save_now()

    def _clear_custom_icon(self):
        self.ui.btn_icon_custom.clear_custom_icon()
        self._update_custom_icon_tooltip("")
        self._active_icon_choice = "default"
        self._set_checked_icon("default")
        if self.app is not None:
            self.app.clear_custom_icon()
            self.app.save_now()

    def _setup_source_buttons(self):
        """按当前配置选中行情数据源单选按钮。"""
        source = self.win.data_source
        self._select_button(self._source_buttons, source)

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
        self.watchlist_editor.refresh_code_search(self.win.codes_list)

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonPress:
            self._handle_outside_press(obj)
        return super().eventFilter(obj, ev)

    def _widget_in_dialog(self, widget) -> bool:
        """widget 是否属于设置对话框本体（悬浮面板等独立顶层窗口不算）。"""
        w = widget
        while w is not None:
            if w is self:
                return True
            if w.window() is w:
                return False
            w = w.parentWidget()
        return False

    def _inside_watchlist(self, widget) -> bool:
        w = widget
        while w is not None and w is not self:
            if w is self.ui.list_codes:
                return True
            w = w.parentWidget()
        return False

    def _pool_for_widget(self, widget):
        w = widget
        while w is not None and w is not self:
            for pool in (self.metric_pool.available_pool, self.metric_pool.displayed_pool):
                if w is pool:
                    return pool
            w = w.parentWidget()
        return None

    def _clear_watchlist_selection(self):
        if self.ui.list_codes.currentRow() >= 0 or self.ui.list_codes.selectedItems():
            self.ui.list_codes.setCurrentCell(-1, -1)

    def _handle_outside_press(self, obj):
        """设置页内按下鼠标时联动清除另一处选中：
        - 按下自选列表：清除指标池选中；
        - 按下指标池：清除自选列表选中与另一池选中；
        - 按下其余空白区域：全部清除并把焦点移回对话框。
        删除/置顶按钮与滚动条除外，避免影响其自身操作。
        """
        if not isinstance(obj, QWidget) or not self._widget_in_dialog(obj):
            return
        if isinstance(obj, QScrollBar):
            return
        if obj in (self.ui.btn_del, self.ui.btn_top):
            return

        pool = self._pool_for_widget(obj)
        if pool is not None:
            self._clear_watchlist_selection()
            for other in (self.metric_pool.available_pool, self.metric_pool.displayed_pool):
                if other is not pool:
                    other.clearSelection()
                    other.setCurrentItem(None)
                    other.clearFocus()
        elif self._inside_watchlist(obj):
            self.metric_pool.clear_selections()
        else:
            self._clear_watchlist_selection()
            self.metric_pool.clear_selections()
            if obj is self:
                self.setFocus()


    def _on_source_toggled(self, source: str, checked: bool):
        if checked:
            self.win.set_data_source(source)

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

    def _show_name_settings_panel(self, anchor=None):
        """点击“名称”后的 ⓘ 切换名称显示设置面板。"""
        self._show_metric_settings_panel(
            self.name_settings_panel, self.unit_settings_panel, anchor
        )

    def _show_unit_settings_panel(self, anchor=None):
        """点击“成交量/成交额”后的 ⓘ 切换单位设置面板。"""
        self._show_metric_settings_panel(
            self.unit_settings_panel, self.name_settings_panel, anchor
        )

    def _show_metric_settings_panel(self, panel, other_panel, anchor):
        if not isinstance(anchor, QWidget):
            anchor = self.metric_pool
        item = anchor.currentItem() if hasattr(anchor, "currentItem") else None
        other_panel.hide()
        # 关闭旧面板会清除两池选中，切换入口时恢复新入口的高亮。
        if item is not None:
            anchor.setCurrentItem(item)
        panel.sync_from(self.win)
        panel.show_for(anchor)

    def _on_hotkey_changed(self):
        new_hotkey = self.keyseq_hide.keySequence().toString()
        result = self.win.update_hotkey(new_hotkey)
        if not result:
            self.keyseq_hide.setKeySequence(QKeySequence(self.win.hotkey))
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))


    def _on_icon_button_toggled(self, key: str, checked: bool):
        if not checked:
            return
        if key == "custom":
            if not self.ui.btn_icon_custom.has_custom_icon():
                self._choose_custom_icon()
                return
            if self.app is not None:
                self.app.set_app_icon("custom")
                if self.app._icon_choice != "custom":
                    self._clear_custom_icon()
                    return
                self.app.save_now()
            self._active_icon_choice = "custom"
            return

        self._active_icon_choice = key
        if self.app is not None:
            self.app.set_app_icon(key)
            self.app.save_now()

    def _on_start_on_boot_toggled(self, checked: bool):
        self.app.set_start_on_boot(bool(checked))




    def _pick_color(self, attr, title, setter):
        base = QColor(getattr(self.win, attr))
        base.setAlpha(255)
        color = QColorDialog.getColor(base, self, f"选择{title}")
        if color.isValid():
            setter(color)
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

    def _sync_click_through_from_win(self, checked: bool):
        """浮窗鼠标穿透状态变化（如快捷键触发）时同步设置窗口复选框"""
        self._set_checked_blocked(self.cb_click_through, checked)

    @staticmethod
    def _set_checked_blocked(widget, checked: bool):
        """设置可勾选控件的状态，并屏蔽其 toggled 信号，避免反向触发浮窗改动。"""
        with QSignalBlocker(widget):
            widget.setChecked(bool(checked))

    def _sync_display_flags_from_win(self):
        """浮窗右键菜单等外部途径修改显示状态时，同步指标池。"""
        self.metric_pool.set_visible_metrics(self.win.visible_metrics)
        self.name_settings_panel.sync_from(self.win)
        self.unit_settings_panel.sync_from(self.win)
        self._set_checked_blocked(self.cb_head, self.win.header_visible)
        self._set_checked_blocked(self.cb_grid, self.win.grid_visible)
        self._set_checked_blocked(self.cb_unicolor, self.win.unicolor)
        self._update_direction_color_controls()
        self._refresh_color_buttons()


    def _on_hotkey_hide_enabled_toggled(self, checked: bool):
        self.keyseq_hide.setEnabled(bool(checked))
        result = self.win.set_hotkey_enabled(bool(checked))
        if not result:
            # 启用失败(如冲突):回滚复选框与输入框状态,并提示用户
            self.keyseq_hide.setEnabled(False)
            self._set_checked_blocked(self.cb_hotkey_hide, False)
            QMessageBox.warning(self, "快捷键无效", _hotkey_error_message(result))

    def _on_click_through_hotkey_enabled_toggled(self, checked: bool):
        self.keyseq_click_through.setEnabled(bool(checked))
        result = self.win.set_click_through_hotkey_enabled(bool(checked))
        if not result:
            # 启用失败(如冲突):回滚复选框与输入框状态,并提示用户
            self.keyseq_click_through.setEnabled(False)
            self._set_checked_blocked(self.cb_hotkey_click_through, False)
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
        app_version = self.app.app_version if self.app is not None else APP_VERSION
        has_update = bool(self.app is not None and getattr(self.app, "_has_update", False))
        latest_version = getattr(self.app, "_latest_version", None) if self.app is not None else None
        if not has_update:
            latest_version = None
        links = project_links(use_gitee=self._use_gitee_links)
        github_links = project_links()
        gitee_links = project_links(use_gitee=True)

        def link(url, text):
            return (
                f'<a href="{escape(url, quote=True)}" '
                f'style="text-decoration:none; color:#4a90d9;">{escape(text)}</a>'
            )

        version_line = f"当前版本 v{escape(str(app_version))}"
        if latest_version:
            version_line += link(
                github_links["releases"] + "/latest", f"（有更新 v{latest_version}）"
            )
        lines = [
            version_line,
            f'本项目基于 {link(links["license"], "Apache-2.0 License")} 开源',
            "Copyright 2026 sbr0574",
            "&nbsp;",
            "支持：",
            f'仓库地址：{link(github_links["project"], github_links["project"])}',
            f'镜像仓库：{link(gitee_links["project"], gitee_links["project"])}',
            f'发行下载：{link(links["releases"], links["releases"])}',
            f'使用帮助：{link(links["readme"], links["readme"])}',
            f'问题反馈：{link(links["issues"], links["issues"])}',
        ]
        html = "".join(f'<p style="margin:2px 0;">{line}</p>' for line in lines)
        label.setTextFormat(Qt.RichText)
        label.setText(html)
        label.setOpenExternalLinks(True)
        label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        label.unsetCursor()

    def refresh_about(self):
        self._setup_about()

    def _check_update_manually(self):
        """后台检查更新，完成后弹窗提示结果。"""
        self.btn_check_update.setEnabled(False)
        self.btn_check_update.setText("检查中…")

        def _worker():
            try:
                result = get_update_info(
                    self.app.app_version if self.app is not None else APP_VERSION
                )
            except Exception:
                result = (False, None)
            try:
                self.update_check_finished.emit(result)
            except RuntimeError:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_check_finished(self, result):
        self.btn_check_update.setEnabled(True)
        self.btn_check_update.setText("检查更新")
        has_update, latest_version = result
        if self.app is not None:
            self.app._has_update = bool(has_update)
            self.app._latest_version = latest_version if has_update else None
        self._setup_about()

        current_version = (
            self.app.app_version if self.app is not None else APP_VERSION
        )
        if has_update and latest_version:
            QMessageBox.information(
                self,
                "检查更新",
                f"发现新版本 v{latest_version}（当前 v{current_version}）。\n"
                "请前往 Releases 页面下载更新。",
            )
        elif latest_version:
            QMessageBox.information(
                self, "检查更新", f"当前已是最新版本 v{current_version}。"
            )
        else:
            QMessageBox.warning(
                self, "检查更新", "检查更新失败，请检查网络连接后重试。"
            )

    def _open_cache_dir(self):
        """打开配置/缓存目录所在的文件夹。"""
        path = config_paths()
        os.makedirs(path, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def closeEvent(self, event):
        self.watchlist_editor.close()
        self.name_settings_panel.hide()
        self.unit_settings_panel.hide()
        super().closeEvent(event)
