"""设置窗口：装配各页绑定、应用配置及管理面板交互。"""

from contextlib import ExitStack
from functools import partial
from html import escape
import os
import sys
import threading

from PySide6.QtCore import Qt, QEvent, QUrl, QSignalBlocker, Signal
from PySide6.QtGui import QColor, QDesktopServices, QGuiApplication, QIcon, QKeySequence, QWheelEvent
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QDialog,
    QComboBox,
    QColorDialog,
    QButtonGroup,
    QFileDialog,
    QMessageBox,
    QScrollBar,
    QScrollArea,
    QSlider,
)

from stockwidget.constants import APP_VERSION
from stockwidget.core.config_store import config_paths
from stockwidget.core.window_rules import MAX_HIDE_TIMES
from stockwidget.data.update_check import get_update_info, github_available, project_links
from stockwidget.platform.capabilities import (
    hotkeys_supported,
    click_through_supported,
    opacity_supported,
    force_top_supported,
    start_on_boot_supported,
    unsupported_tooltip,
)
from stockwidget.platform.hotkeys import HotkeyResult
from stockwidget.ui.controls.style import (
    LINUX_FONT_RULES,
    build_settings_stylesheet,
    is_dark_theme,
    set_color_button,
    settings_navigation_icon,
    theme_palette,
)
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.generated.ui_settings import Ui_SettingDialog
from stockwidget.ui.settings.groups import (
    FloatRowSettings,
    FloatSplitSettings,
    TaskbarSettings,
)
from stockwidget.ui.watchlist.editor import WatchlistEditor


def _hotkey_error_message(result) -> str:
    """把 HotkeyResult 转成用户可读的中文提示。"""
    if result.reason == "conflict":
        return "快捷键已占用，请更换"
    if result.reason == "reserved":
        return "系统保留快捷键，请更换"
    if result.reason == "invalid":
        if sys.platform in ("win32", "linux"):
            return "请用组合键或 F1–F12"
        return "请使用修饰键 + 主键"
    if result.reason == "unsupported":
        return "当前平台不支持"
    return "注册失败，请重试"


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
        self._designer_stylesheet = self.styleSheet()
        self._init_view_settings()
        # Linux 下用 Tool 窗口避开任务栏/程序坞条目；
        # macOS 的 Dock 图标由应用级 Accessory 激活策略隐藏
        # （见 app._hide_macos_dock_icon），窗口保持普通标题栏，
        # 避免 Tool 窗口的小号红黄绿按钮与小标题。
        if sys.platform == "linux":
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.Tool)
            # 在指标池按文字宽度计算尺寸前应用字号，最终主题仍保留这些规则。
            self.setStyleSheet(LINUX_FONT_RULES)
        self._init_hotkey_status()
        self._init_metric_pool()
        self.setModal(False)
        self.watchlist_editor = WatchlistEditor(
            self.ui.list_codes, self.ui.btn_add, self.ui.btn_del, self.ui.btn_top,
            self.win.watchlist, self.win.codes_list, self,
        )
        self.watchlist_editor.watchlist_changed.connect(self.win.set_watchlist)
        self._connect_controls()
        self._setup_icon_choices()
        self._load_settings()
        self.ui.settings_navigation.currentRowChanged.connect(self._on_settings_page_changed)
        self.ui.settings_navigation.setCurrentRow(0)
        self._apply_theme_stylesheet()
        QGuiApplication.styleHints().colorSchemeChanged.connect(self._apply_theme_stylesheet)
        QApplication.instance().paletteChanged.connect(self._apply_theme_stylesheet)
        self.github_check_finished.connect(self._on_github_check_finished)
        self.update_check_finished.connect(self._on_update_check_finished)
        self._start_github_check()
        # 全局监听鼠标按下：点击设置页空白/其他区域时清除自选列表与指标池的选中
        QApplication.instance().installEventFilter(self)

    def _init_hotkey_status(self):
        self._hotkey_rows = (
            ("hotkey", self.ui.cb_hotkey_hide, self.ui.keyseq_hide),
            ("hotkey_click_through", self.ui.cb_hotkey_click_through,
             self.ui.keyseq_click_through),
        )
        self._hotkey_status = {}
        for key, _checkbox, editor in self._hotkey_rows:
            editor.setMaximumSequenceLength(1)
            status = getattr(self.ui, key + "_status")
            self._hotkey_status[key] = status

    def _refresh_hotkey_status(self):
        for key, checkbox, editor in self._hotkey_rows:
            status = self._hotkey_status[key]
            status.setVisible(checkbox.isChecked() and checkbox.isEnabled())
            result = self.win.hotkey_results.get(key, HotkeyResult(False, "failed"))
            message = "已生效" if result else _hotkey_error_message(result)
            status.set_result(result, message)

    def _init_metric_pool(self):
        self.metric_pool = self.ui.float_metric_pool

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
        mode = self.win.view_options.color_mode
        self._applied_color_mode = mode
        dark = mode == "dark" or (mode == "system" and is_dark_theme())
        self.setStyleSheet(build_settings_stylesheet(dark, linux_fonts=sys.platform == "linux")
                          + "\n" + self._designer_stylesheet)
        self.setPalette(QApplication.palette() if mode == "system" else theme_palette(dark, QApplication.palette()))
        for component in (self.metric_pool, self.taskbar_settings.metric_pool, self.watchlist_editor):
            component.set_theme(dark)
        self._refresh_color_buttons()
        navigation = self.ui.settings_navigation
        for index in range(navigation.count()):
            navigation.item(index).setIcon(settings_navigation_icon(index, self.devicePixelRatioF(), self.palette()))

    def _on_settings_page_changed(self, _index):
        self.watchlist_editor.add_code_panel.hide()
        self._clear_watchlist_selection()
        self._clear_metric_selections()

    def _refresh_color_buttons(self):
        """用无描边圆形图标展示五个颜色按钮的当前色值。"""
        for button, attr, title in self._color_buttons:
            color = QColor(getattr(self.win, attr))
            set_color_button(button, color, title)

    def _connect_controls(self):
        """连接设置控件与浮窗状态，静态控件统一从 self.ui 访问。"""
        self._source_buttons = {
            "sina": self.ui.rb_sina,
            "eastmoney": self.ui.rb_em,
        }

        self._color_buttons = (
            (self.ui.btn_bg_color, "bg", "背景颜色"),
            (self.ui.btn_fg_color, "fg", "文字颜色"),
            (self.ui.btn_up_color, "up_color", "上涨颜色"),
            (self.ui.btn_down_color, "down_color", "下跌颜色"),
            (self.ui.btn_neutral_color, "neutral_color", "中性颜色"),
        )

        self.ui.sb_interval.valueChanged.connect(self.win.set_refresh_interval)
        for source, button in self._source_buttons.items():
            button.toggled.connect(partial(self._on_source_toggled, source))

        self.metric_pool.visible_metrics_changed.connect(
            self.win.set_visible_metrics
        )
        self.ui.cb_name_visible.toggled.connect(self._set_name_visible)
        self.ui.cmb_name_length.currentIndexChanged.connect(
            lambda _: self.win.set_name_length(self.ui.cmb_name_length.currentData()))
        self.ui.cb_code_visible.toggled.connect(self.win.set_code_visible)
        self.ui.cb_type_visible.toggled.connect(self.win.set_type_visible)
        self.ui.cmb_unit_mode.currentIndexChanged.connect(
            lambda _: self.win.set_unit_mode(self.ui.cmb_unit_mode.currentData()))
        for index, value in enumerate((-1, 1, 2, 3, 4)):
            self.ui.cmb_name_length.setItemData(index, value)
        for index, value in enumerate(("auto", "cn", "en")):
            self.ui.cmb_unit_mode.setItemData(index, value)
        for index, value in enumerate(("system", "light", "dark")):
            self.ui.cmb_color_mode.setItemData(index, value)
        self.ui.cmb_color_mode.currentIndexChanged.connect(
            lambda _: self.win.set_view_options(color_mode=self.ui.cmb_color_mode.currentData()))
        self.ui.cb_hide_tray_icon.toggled.connect(
            lambda hidden: self.win.set_view_options(hide_tray_icon=hidden))
        self.ui.cb_chart_enabled.toggled.connect(
            lambda enabled: self.win.set_view_options(chart_enabled=enabled))
        self.win.view_options_changed.connect(self._sync_common_options)

        self.ui.btn_check_update.clicked.connect(self._check_update_manually)
        self.ui.btn_open_cache_dir.clicked.connect(self._open_cache_dir)
        self.ui.btn_clear_watchlist.clicked.connect(self.watchlist_editor.clear_watchlist)
        self.ui.btn_reset_appearance.clicked.connect(self._reset_appearance)
        self.ui.btn_reset_settings.clicked.connect(self._reset_settings)

        self.ui.cb_unicolor.toggled.connect(self._on_unicolor_toggled)
        color_setters = {
            "bg": self.win.set_bg_rgb_keep_alpha,
            "fg": self.win.set_fg_color,
            "up_color": self.win.set_up_color,
            "down_color": self.win.set_down_color,
            "neutral_color": self.win.set_neutral_color,
        }
        for button, attr, title in self._color_buttons:
            button.clicked.connect(partial(self._pick_color, attr, title, color_setters[attr]))
        self.ui.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.ui.slider_all_alpha.valueChanged.connect(self.apply_win_opacity)

        self.ui.cmb_font.currentTextChanged.connect(self.win.set_font_family)
        self.ui.slider_font_size.valueChanged.connect(self.apply_font_size)
        self.ui.slider_line_interval.valueChanged.connect(self._on_line_changed)
        self.ui.keyseq_hide.editingFinished.connect(self._on_hotkey_changed)
        self.ui.keyseq_click_through.editingFinished.connect(self._on_click_through_hotkey_changed)
        self.ui.cb_auto_start.toggled.connect(self._on_start_on_boot_toggled)
        self.ui.cb_float_on_top.toggled.connect(self.win.set_float_on_top)
        self.ui.cb_force_top.toggled.connect(self.win.set_force_top)
        self.win.topmost_changed.connect(self._sync_topmost_from_win)
        self.ui.cb_click_through.toggled.connect(self.win.set_click_through)
        self.win.click_through_changed.connect(self._sync_click_through_from_win)
        # 浮窗右键菜单等外部途径修改显示指标时，同步设置窗口复选框
        self.win.display_flags_changed.connect(self._sync_display_flags_from_win)
        self.win.presentation_changed.connect(self._sync_metric_options)
        self.ui.cb_hotkey_hide.toggled.connect(self._on_hotkey_hide_enabled_toggled)
        self.ui.cb_hotkey_click_through.toggled.connect(self._on_click_through_hotkey_enabled_toggled)
        self.ui.cb_head.toggled.connect(self.win.set_header_visible)
        self.ui.cb_grid.toggled.connect(self.win.set_grid_visible)
        self.ui.cb_auto_hide.toggled.connect(
            lambda enabled: self.win.set_hide_options(auto_hide_enabled=enabled))
        self.ui.gb_scheduled_hide.toggled.connect(
            lambda enabled: self.win.set_hide_options(hide_enabled=enabled,
                                                      scheduled_hide_enabled=enabled))
        self.ui.btn_add_hide_time.clicked.connect(self._add_hide_time)
        self.ui.list_hide_times.remove_requested.connect(self._delete_hide_time)
        self.ui.hide_time_edit.timeChanged.connect(self._update_hide_time_buttons)
        self.ui.list_hide_times.currentRowChanged.connect(self._update_hide_time_buttons)
        self.win.hide_options_changed.connect(self._sync_hide_settings)
        self.ui.cb_boundary_check.toggled.connect(
            lambda enabled: self.win.set_position_options(boundary_check_enabled=enabled))
        self.ui.cb_edge_hide.toggled.connect(
            lambda enabled: self.win.set_position_options(edge_hide_enabled=enabled))
        self.win.position_options_changed.connect(self._sync_position_settings)

    def _init_view_settings(self):
        self.float_row_settings = self.ui.float_row_settings
        self.float_split_settings = self.ui.float_split_settings
        self.taskbar_settings = self.ui.taskbar_settings
        self._view_bindings = []
        for group, kind in ((self.float_row_settings, FloatRowSettings),
                            (self.float_split_settings, FloatSplitSettings),
                            (self.taskbar_settings, TaskbarSettings)):
            binding = kind(group)
            binding.bind(self.win)
            self._view_bindings.append(binding)
        self.taskbar_paging_settings = self.taskbar_settings.paging
        self.taskbar_style_settings = self.taskbar_settings.style
        self.taskbar_split_settings = self.taskbar_settings.split

    def _sync_view_settings(self):
        for binding in self._view_bindings:
            binding.sync()

    def _load_settings(self):
        with ExitStack() as stack:
            for widget in self.findChildren(QWidget):
                stack.enter_context(QSignalBlocker(widget))
            self.ui.sb_interval.setValue(self.win.refresh_seconds)
            self.metric_pool.set_visible_metrics(self.win.visible_metrics)
            self._sync_metric_options()
            self._sync_common_options()

            self._set_checked_blocked(self.ui.cb_unicolor, self.win.unicolor)
            self._update_direction_color_controls()
            self._refresh_color_buttons()
            self.ui.slider_bg_alpha.setValue(int(round(self.win.bg.alpha() / 2.55)))
            self.ui.label_bg_alpha.setText(f"{self.ui.slider_bg_alpha.value()}%")
            self.ui.slider_all_alpha.setValue(self.win.opacity_pct)
            self.ui.label_all_alpha.setText(f"{self.ui.slider_all_alpha.value()}%")

            self.ui.cmb_font.setCurrentText(self.win.font.family())
            self.ui.slider_font_size.setValue(self.win.font.pointSize())
            self.ui.label_current_font_size.setText(f"{self.win.font.pointSize()} pt")
            self.ui.slider_line_interval.setValue(self.win.line_extra_px)
            self.ui.label_current_line_interval.setText(f"+{self.ui.slider_line_interval.value()} px")

            self.ui.keyseq_hide.setKeySequence(QKeySequence(self.win.hotkey))
            self.ui.keyseq_hide.setEnabled(self.win.hotkey_enabled)
            self.ui.keyseq_click_through.setKeySequence(QKeySequence(self.win.hotkey_click_through))
            self.ui.keyseq_click_through.setEnabled(self.win.hotkey_click_through_enabled)
            self.ui.cb_hotkey_hide.setChecked(self.win.hotkey_enabled)
            self.ui.cb_hotkey_click_through.setChecked(self.win.hotkey_click_through_enabled)
            self.ui.cb_auto_start.setChecked(bool(self.win.start_on_boot))
            self._sync_topmost_from_win()
            self.ui.cb_click_through.setChecked(self.win.click_through)
            self.ui.cb_head.setChecked(self.win.header_visible)
            self.ui.cb_grid.setChecked(self.win.grid_visible)
            self._sync_hide_settings()
            self._sync_position_settings()

            self._apply_platform_limits()
            self._sync_view_settings()
            self._refresh_hotkey_status()
            self._setup_source_buttons()
            self._setup_about()
            self.refresh_data_state()

    def _sync_position_settings(self):
        self._set_checked_blocked(self.ui.cb_boundary_check, self.win.boundary_check_enabled)
        self._set_checked_blocked(self.ui.cb_edge_hide, self.win.edge_hide_enabled)
        self.ui.cb_edge_hide.setEnabled(self.win.boundary_check_enabled)

    def _sync_hide_settings(self):
        with QSignalBlocker(self.ui.cb_auto_hide), QSignalBlocker(self.ui.gb_scheduled_hide):
            self.ui.cb_auto_hide.setChecked(self.win.auto_hide_enabled)
            self.ui.gb_scheduled_hide.setChecked(self.win.hide_enabled)
        times_text = "，".join(self.win.scheduled_hide_times) if self.win.scheduled_hide_enabled else ""
        if not self.win.hide_enabled:
            status = "程序未启用自动隐藏"
        elif times_text and self.win.auto_hide_enabled:
            status = f"程序将在{times_text}及无数据更新时自动隐藏"
        elif times_text:
            status = f"程序将在{times_text}自动隐藏"
        elif self.win.auto_hide_enabled:
            status = "程序将在无数据更新时自动隐藏"
        else:
            status = "程序未启用自动隐藏"
        self.ui.label_hide_status.setText(status)
        times = self.ui.list_hide_times
        current = times.currentItem().text() if times.currentItem() else None
        if [times.item(i).text() for i in range(times.count())] != self.win.scheduled_hide_times:
            with QSignalBlocker(times):
                times.clear()
                times.addItems(self.win.scheduled_hide_times)
                if current in self.win.scheduled_hide_times:
                    times.setCurrentRow(self.win.scheduled_hide_times.index(current))
                elif times.count():
                    times.setCurrentRow(0)
        self._update_hide_time_buttons()

    def _update_hide_time_buttons(self, *_args):
        value = self.ui.hide_time_edit.time().toString("HH:mm")
        enabled = self.win.hide_enabled
        full = len(self.win.scheduled_hide_times) >= MAX_HIDE_TIMES
        duplicate = value in self.win.scheduled_hide_times
        self.ui.btn_add_hide_time.setEnabled(enabled
            and not full and not duplicate)
        self.ui.btn_add_hide_time.setToolTip("最多设置3个时间，请先删除一个" if full else
                                           "该时间已添加" if duplicate else "添加每日隐藏时间")

    def _add_hide_time(self):
        value = self.ui.hide_time_edit.time().toString("HH:mm")
        if (not self.win.hide_enabled or value in self.win.scheduled_hide_times
                or len(self.win.scheduled_hide_times) >= MAX_HIDE_TIMES):
            return
        self.win.set_hide_options(scheduled_hide_enabled=True,
                                 scheduled_hide_times=[*self.win.scheduled_hide_times, value])
        self.ui.list_hide_times.setCurrentRow(self.win.scheduled_hide_times.index(value))

    def _delete_hide_time(self, time):
        if self.win.hide_enabled:
            self.win.set_hide_options(scheduled_hide_times=[value for value in self.win.scheduled_hide_times
                                                          if value != time])

    def _reset_appearance(self):
        self.win.reset_appearance()
        self._clear_custom_icon()
        self._load_settings()

    def _reset_settings(self):
        self.win.reset_settings()
        self._on_start_on_boot_toggled(self.win.start_on_boot)
        self._load_settings()

    def _apply_platform_limits(self):
        """按当前平台禁用不支持的功能控件:
        - Wayland 下:全局快捷键、鼠标穿透、窗口整体透明度不可用。
        - Linux 下:强制置顶不可用(raise_ 受窗口管理器/合成器限制)。
        """
        self.ui.settings_navigation.item(5).setHidden(sys.platform != "win32")
        if not hotkeys_supported():
            for w in (self.ui.cb_hotkey_hide, self.ui.cb_hotkey_click_through,
                      self.ui.keyseq_hide, self.ui.keyseq_click_through):
                w.setEnabled(False)
                w.setToolTip(unsupported_tooltip("全局快捷键"))
        if not click_through_supported():
            self.ui.cb_click_through.setEnabled(False)
            self.ui.cb_click_through.setToolTip(unsupported_tooltip("鼠标穿透"))
            # 鼠标穿透不可用（如 macOS）时，其快捷键一并关闭
            for w in (self.ui.cb_hotkey_click_through, self.ui.keyseq_click_through):
                w.setEnabled(False)
                w.setToolTip(unsupported_tooltip("鼠标穿透"))
        if not opacity_supported():
            # 整体不透明度滑块:Wayland 平台插件不支持设置窗口透明度
            for w in (self.ui.slider_all_alpha, self.ui.label_all, self.ui.label_all_alpha):
                w.setEnabled(False)
            self.ui.slider_all_alpha.setToolTip(unsupported_tooltip("整体不透明度"))
        self._sync_topmost_from_win()
        if not start_on_boot_supported():
            self.ui.cb_auto_start.setEnabled(False)
            self.ui.cb_auto_start.setToolTip(unsupported_tooltip("开机自启"))

    def _sync_topmost_from_win(self):
        self._set_checked_blocked(self.ui.cb_float_on_top, self.win.float_on_top)
        self._set_checked_blocked(self.ui.cb_force_top, self.win.force_top)
        supported = force_top_supported()
        self.ui.cb_force_top.setEnabled(self.win.float_on_top and supported)
        self.ui.cb_force_top.setToolTip(
            unsupported_tooltip("强制置顶", suggest_x11=False) if not supported else
            "请先启用浮窗置顶" if not self.win.float_on_top else
            "每秒恢复浮窗置顶，避免被其他置顶窗口遮盖"
        )

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
        self.ui.label_data_state.setText(text)
        error = self.app.code_data_error() if self.app is not None else ""
        self.ui.label_data_state.setToolTip(
            f"{error}\n继续使用本地缓存，30 分钟后自动重试。" if isinstance(error, str) and error else ""
        )

    def refresh_code_search(self):
        self.watchlist_editor.refresh_code_search(self.win.codes_list)

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.Wheel and isinstance(obj, (QSlider, QComboBox)) and self._widget_in_dialog(obj):
            parent = obj.parentWidget()
            while parent and not isinstance(parent, QScrollArea):
                parent = parent.parentWidget()
            if parent:
                viewport = parent.viewport()
                forwarded = QWheelEvent(viewport.mapFromGlobal(ev.globalPosition()), ev.globalPosition(),
                                        ev.pixelDelta(), ev.angleDelta(), ev.buttons(), ev.modifiers(),
                                        ev.phase(), ev.inverted(), ev.source())
                QApplication.sendEvent(viewport, forwarded)
            return True
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
            for pool in self._metric_lists():
                if w is pool:
                    return pool
            w = w.parentWidget()
        return None

    def _metric_lists(self):
        return (self.metric_pool.available_pool, self.metric_pool.displayed_pool,
                self.taskbar_settings.metric_pool.available_pool, self.taskbar_settings.metric_pool.displayed_pool)

    def _clear_metric_selections(self, keep=None):
        for pool in self._metric_lists():
            if pool is not keep:
                pool.clearSelection()
                pool.setCurrentItem(None)
                pool.clearFocus()

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
            self._clear_metric_selections(keep=pool)
        elif self._inside_watchlist(obj):
            self._clear_metric_selections()
        else:
            self._clear_watchlist_selection()
            self._clear_metric_selections()
            if obj is self:
                self.setFocus()

    def _on_source_toggled(self, source: str, checked: bool):
        if checked:
            self.win.set_data_source(source)

    def _update_direction_color_controls(self):
        enabled = not self.win.unicolor
        for widget in (
            self.ui.btn_up_color,
            self.ui.btn_down_color,
            self.ui.btn_neutral_color,
        ):
            widget.setEnabled(enabled)

    def _on_unicolor_toggled(self, checked: bool):
        self.win.set_unicolor(bool(checked))
        self._update_direction_color_controls()
        self._refresh_color_buttons()

    def _set_name_visible(self, checked):
        self.win.set_name_length(self.ui.cmb_name_length.currentData() if checked else 0)

    def _sync_metric_options(self):
        controls = (self.ui.cb_name_visible, self.ui.cmb_name_length, self.ui.cb_code_visible,
                    self.ui.cb_type_visible, self.ui.cmb_unit_mode)
        with ExitStack() as stack:
            for control in controls:
                stack.enter_context(QSignalBlocker(control))
            self.ui.cb_name_visible.setChecked(self.win.name_length != 0)
            if self.win.name_length != 0:
                self.ui.cmb_name_length.setCurrentIndex(max(0, self.ui.cmb_name_length.findData(self.win.name_length)))
            self.ui.cmb_name_length.setEnabled(self.win.name_length != 0)
            self.ui.cb_code_visible.setChecked(self.win.code_visible)
            self.ui.cb_type_visible.setChecked(self.win.type_visible)
            self.ui.cmb_unit_mode.setCurrentIndex(self.ui.cmb_unit_mode.findData(self.win.unit_mode))

    def _sync_common_options(self):
        with QSignalBlocker(self.ui.cb_chart_enabled):
            self.ui.cb_chart_enabled.setChecked(self.win.view_options.chart_enabled)
        with QSignalBlocker(self.ui.cmb_color_mode), QSignalBlocker(self.ui.cb_hide_tray_icon):
            self.ui.cmb_color_mode.setCurrentIndex(self.ui.cmb_color_mode.findData(self.win.view_options.color_mode))
            self.ui.cb_hide_tray_icon.setChecked(self.win.view_options.hide_tray_icon)
        if self.win.view_options.color_mode != getattr(self, "_applied_color_mode", None):
            self._apply_theme_stylesheet()

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

    def _on_hotkey_changed(self):
        self.win.update_hotkey(self.ui.keyseq_hide.keySequence().toString())
        self._refresh_hotkey_status()

    def _on_click_through_hotkey_changed(self):
        self.win.update_click_through_hotkey(self.ui.keyseq_click_through.keySequence().toString())
        self._refresh_hotkey_status()

    def _on_start_on_boot_toggled(self, checked: bool):
        self.win.start_on_boot = bool(checked)
        if self.app is not None:
            self.app.set_start_on_boot(bool(checked))
            self.app.save_now()

    def _pick_color(self, attr, title, setter):
        base = QColor(getattr(self.win, attr))
        base.setAlpha(255)
        color = QColorDialog.getColor(base, self, f"选择{title}")
        if color.isValid():
            setter(color)
            self._refresh_color_buttons()

    def apply_bg_alpha(self, v: int):
        self.ui.label_bg_alpha.setText(f"{v}%")
        self.win.set_bg_alpha_percent(v)

    def apply_win_opacity(self, v: int):
        self.ui.label_all_alpha.setText(f"{v}%")
        self.win.set_window_opacity_percent(v)

    def apply_font_size(self, v: int):
        self.ui.label_current_font_size.setText(f"{v} pt")
        self.win.set_font_size(v)

    def _on_line_changed(self, v: int):
        self.ui.label_current_line_interval.setText(f"+{v} px")
        self.win.set_line_extra(v)

    def _sync_click_through_from_win(self, checked: bool):
        """浮窗鼠标穿透状态变化（如快捷键触发）时同步设置窗口复选框"""
        self._set_checked_blocked(self.ui.cb_click_through, checked)

    @staticmethod
    def _set_checked_blocked(widget, checked: bool):
        """设置可勾选控件的状态，并屏蔽其 toggled 信号，避免反向触发浮窗改动。"""
        with QSignalBlocker(widget):
            widget.setChecked(bool(checked))

    def _sync_display_flags_from_win(self):
        """浮窗右键菜单等外部途径修改显示状态时，同步指标池。"""
        self.metric_pool.set_visible_metrics(self.win.visible_metrics)
        self._sync_metric_options()
        self._set_checked_blocked(self.ui.cb_head, self.win.header_visible)
        self._set_checked_blocked(self.ui.cb_grid, self.win.grid_visible)
        self._set_checked_blocked(self.ui.cb_unicolor, self.win.unicolor)
        self._update_direction_color_controls()
        self._refresh_color_buttons()

    def _on_hotkey_hide_enabled_toggled(self, checked: bool):
        self.ui.keyseq_hide.setEnabled(bool(checked))
        self.win.set_hotkey_enabled(bool(checked))
        self._refresh_hotkey_status()

    def _on_click_through_hotkey_enabled_toggled(self, checked: bool):
        self.ui.keyseq_click_through.setEnabled(bool(checked))
        self.win.set_click_through_hotkey_enabled(bool(checked))
        self._refresh_hotkey_status()

    def _on_click_through_hotkey_changed(self):
        new_hotkey = self.ui.keyseq_click_through.keySequence().toString()
        self.win.update_click_through_hotkey(new_hotkey)
        self._refresh_hotkey_status()

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

        version_line = "📦 当前程序版本："
        if latest_version:
            version_line += link(
                github_links["releases"] + "/latest", f"有更新 (v{escape(str(app_version))} -> v{latest_version})"
            )
        else:
            version_line += f"最新 (v{escape(str(app_version))})"
        self.ui.label_version_state.setText(version_line)
        lines = [
            f'本项目基于 {link(links["license"], "Apache-2.0 License")} 开源',
            "Copyright 2026 sbr0574<br>",
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
        self.ui.btn_check_update.setEnabled(False)
        self.ui.btn_check_update.setText("检查中…")

        def _worker():
            errors = []
            try:
                has_update, latest_version = get_update_info(
                    self.app.app_version if self.app is not None else APP_VERSION,
                    errors=errors,
                )
            except Exception:
                has_update, latest_version = False, None
                errors.append("更新检查发生异常")
            result = (has_update, latest_version, "\n".join(errors))
            try:
                self.update_check_finished.emit(result)
            except RuntimeError:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_check_finished(self, result):
        self.ui.btn_check_update.setEnabled(True)
        self.ui.btn_check_update.setText("检查程序更新")
        has_update, latest_version = result[:2]
        error = result[2] if len(result) > 2 else ""
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
                self, "检查更新", "检查更新失败。\n" + (error or "未获取到版本信息，请稍后重试。")
            )

    def _open_cache_dir(self):
        """打开配置/缓存目录所在的文件夹。"""
        path = config_paths()
        os.makedirs(path, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def closeEvent(self, event):
        self.watchlist_editor.close()
        super().closeEvent(event)
