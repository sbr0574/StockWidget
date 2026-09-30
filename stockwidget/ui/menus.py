# -*- coding: utf-8 -*-
"""系统托盘封装：统一 Windows / macOS / Linux 的菜单与点击行为。

平台差异：
- Windows：左键单击/双击切换显示隐藏，右键弹出菜单。
- macOS / Linux：单击（不区分左右键）直接弹出菜单，无左键切换逻辑。
"""

from functools import partial
import sys

from PySide6.QtCore import Qt, QSignalBlocker
from PySide6.QtGui import QAction, QActionGroup
from PySide6.QtWidgets import QMenu, QSystemTrayIcon

from stockwidget.core.quote_presentation import METRIC_SPECS, SORTABLE_HEADERS, metric_headers
from stockwidget.core.view_options import DISPLAY_MODES
from stockwidget.platform.capabilities import click_through_supported, tray_click_toggles


def display_mode_actions(menu, on_change):
    """托盘和行情菜单使用同一组显示位置选项。"""
    actions = {}
    group = QActionGroup(menu)
    for mode, label in DISPLAY_MODES:
        action = menu.addAction(label)
        action.setCheckable(True)
        group.addAction(action)
        action.triggered.connect(lambda checked=False, value=mode: on_change(value))
        actions[mode] = action
    return actions


def sync_display_modes(actions, current, taskbar_enabled):
    for mode, action in actions.items():
        with QSignalBlocker(action):
            action.setChecked(current == mode)
        enabled = mode == "float" or taskbar_enabled
        action.setEnabled(enabled)
        action.setToolTip("" if enabled else "请先在设置的任务栏页启用任务栏模式")


class TrayIcon(QSystemTrayIcon):
    """托盘图标 + 右键菜单 + 平台感知的点击行为。"""

    def __init__(self, icon, app_name, *,
                 on_toggle, on_open_settings, on_quit,
                 on_click_through, click_through_getter,
                 on_display_mode=None, display_mode_getter=None, taskbar_enabled_getter=None):
        super().__init__(icon)
        self._on_toggle = on_toggle
        self._on_click_through = on_click_through
        self._click_through_getter = click_through_getter

        self.setToolTip(app_name)

        menu = QMenu()
        menu.addAction(QAction("显示/隐藏", self, triggered=self._on_toggle))
        self._display_mode_getter = display_mode_getter
        self._taskbar_enabled_getter = taskbar_enabled_getter
        self._mode_actions = {}
        if sys.platform == "win32" and on_display_mode and display_mode_getter:
            self._mode_actions = display_mode_actions(menu.addMenu("显示方式"), on_display_mode)

        self.act_click_through = QAction("鼠标穿透", self, checkable=True)
        self.act_click_through.setChecked(bool(self._click_through_getter()))
        self._sync_display_modes()
        self.act_click_through.toggled.connect(self._on_click_through)
        if not click_through_supported():
            # 当前平台（如 Wayland）不支持鼠标穿透，置为不可点按
            self.act_click_through.setEnabled(False)
            self.act_click_through.setToolTip("当前会话不支持鼠标穿透")
        menu.addAction(self.act_click_through)

        menu.addAction(QAction("设置…", self, triggered=on_open_settings))
        menu.addSeparator()
        menu.addAction(QAction("退出", self, triggered=on_quit))
        menu.aboutToShow.connect(self.sync_click_through)
        self.setContextMenu(menu)

        self.activated.connect(self._on_activated)

    def sync_click_through(self):
        """菜单显示前，用浮窗当前状态同步「鼠标穿透」勾选。"""
        self.act_click_through.setChecked(bool(self._click_through_getter()))
        self._sync_display_modes()

    def _sync_display_modes(self):
        if self._display_mode_getter:
            sync_display_modes(self._mode_actions, self._display_mode_getter(),
                               not self._taskbar_enabled_getter or self._taskbar_enabled_getter())

    def _on_activated(self, reason):
        # Windows 左键切换；macOS/Linux 单击即弹菜单，无切换逻辑
        if tray_click_toggles() and reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._on_toggle()


def build_quote_menu(source, surface="float"):
    menu = QMenu(source)
    sub_cols = QMenu("显示指标", menu)
    metrics = source.get_surface_metrics(surface)
    for spec in METRIC_SPECS:
        action = QAction(spec.label, sub_cols, checkable=True)
        action.setChecked(spec.metric_id in metrics)
        action.toggled.connect(partial(source.set_surface_metric_visible, surface, spec.metric_id))
        sub_cols.addAction(action)
    menu.addMenu(sub_cols)

    sort_menu = QMenu("排序", menu)
    visible_headers = set(metric_headers(metrics))
    for header_name in SORTABLE_HEADERS:
        metric_menu = QMenu(header_name, sort_menu)
        metric_menu.setEnabled(header_name in visible_headers)
        asc = QAction("升序", metric_menu, checkable=True)
        desc = QAction("降序", metric_menu, checkable=True)
        asc.setChecked(
            source.quotes.sort_header == header_name
            and source.quotes.sort_order == Qt.SortOrder.AscendingOrder
        )
        desc.setChecked(
            source.quotes.sort_header == header_name
            and source.quotes.sort_order == Qt.SortOrder.DescendingOrder
        )
        asc.triggered.connect(
            partial(source.quotes.set_sort, header_name, Qt.SortOrder.AscendingOrder)
        )
        desc.triggered.connect(
            partial(source.quotes.set_sort, header_name, Qt.SortOrder.DescendingOrder)
        )
        metric_menu.addAction(asc)
        metric_menu.addAction(desc)
        sort_menu.addMenu(metric_menu)
    sort_menu.addSeparator()
    clear_action = QAction("恢复自选顺序", sort_menu)
    clear_action.setEnabled(source.quotes.sort_header is not None)
    clear_action.triggered.connect(source.quotes.clear_sort)
    sort_menu.addAction(clear_action)
    menu.addMenu(sort_menu)

    act_header = QAction("显示表头", menu, checkable=True)
    act_header.setChecked(source.header_visible)
    act_header.toggled.connect(source.set_header_visible)
    menu.addAction(act_header)

    act_grid = QAction("显示网格",menu, checkable=True)
    act_grid.setChecked(source.grid_visible)
    act_grid.toggled.connect(source.set_grid_visible)
    menu.addAction(act_grid)

    act_color = QAction("统一颜色", menu, checkable=True)
    act_color.setChecked(source.unicolor)
    act_color.toggled.connect(source.set_unicolor)
    menu.addAction(act_color)

    if sys.platform == "win32":
        actions = display_mode_actions(menu.addMenu("显示位置"), source.set_display_mode)
        sync_display_modes(actions, source.display_mode, source.view_options.taskbar_enabled)

    menu.addSeparator()
    act_open_settings = QAction("设置…", menu)
    if callable(source._open_settings_cb):
        act_open_settings.triggered.connect(source._open_settings_cb)
    else:
        act_open_settings.setEnabled(False)
    menu.addAction(act_open_settings)

    menu.addSeparator()
    menu.addAction(QAction("隐藏", menu, triggered=source.hide_widget))
    return menu
