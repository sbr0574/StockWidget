# -*- coding: utf-8 -*-
"""系统托盘封装：统一 Windows / macOS / Linux 的菜单与点击行为。

平台差异：
- Windows：左键单击/双击切换显示隐藏，右键弹出菜单。
- macOS / Linux：单击（不区分左右键）直接弹出菜单，无左键切换逻辑。
"""

from functools import partial
import sys

from PySide6.QtCore import Qt, QSignalBlocker, QUrl
from PySide6.QtGui import QAction, QDesktopServices
from PySide6.QtWidgets import QApplication, QMenu, QSystemTrayIcon

from stockwidget.core.quote_presentation import METRIC_SPECS, SORTABLE_HEADERS, metric_headers
from stockwidget.data.update_check import project_links
from stockwidget.platform.capabilities import click_through_supported, tray_click_toggles


def _sync_toggles(toggles):
    for action, getter, tooltip in toggles:
        with QSignalBlocker(action):
            action.setChecked(bool(getter()))
        if tooltip:
            action.setToolTip(tooltip() if callable(tooltip) else tooltip)


def _surface_scope(source, surface, sync_option):
    """Describe the effective setting without changing its saved sync preference."""
    current = surface() if callable(surface) else surface
    if (sys.platform == "win32" and getattr(source.view_options, sync_option)
            and (current == "taskbar" or source.view_options.taskbar_enabled)):
        return "浮窗和任务栏已同步此设置，修改会同时影响两处。"
    return "仅影响任务栏行情。" if current == "taskbar" else "仅影响浮窗行情。"


def _view_toggles(menu, source, surface="float", *, include_taskbar=False):
    """Bind both menus to the same saved options as the settings page."""
    toggles = []

    def add(text, getter, setter, tooltip=None):
        action = menu.addAction(text)
        action.setCheckable(True)
        action.toggled.connect(setter)
        toggles.append((action, getter, tooltip))
        return action

    if include_taskbar and sys.platform == "win32":
        add("任务栏行情", lambda: source.view_options.taskbar_enabled,
            lambda enabled: source.set_view_options(taskbar_enabled=enabled))

    def split_key():
        current = surface() if callable(surface) else surface
        return "float_split_enabled" if current == "float" or source.view_options.taskbar_sync_split else "taskbar_split_enabled"

    add("分栏", lambda: getattr(source.view_options, split_key()),
        lambda enabled: source.set_view_options(**{split_key(): enabled}),
        lambda: _surface_scope(source, surface, "taskbar_sync_split"))
    through = add("鼠标穿透", lambda: source.click_through, source.set_click_through,
                  "可从托盘菜单关闭鼠标穿透。" if click_through_supported()
                  else "当前会话不支持鼠标穿透")
    if not click_through_supported():
        through.setEnabled(False)
    _sync_toggles(toggles)
    return toggles


class TrayIcon(QSystemTrayIcon):
    """托盘图标 + 右键菜单 + 平台感知的点击行为。"""

    def __init__(self, icon, app_name, *,
                 on_toggle, on_open_settings, on_quit,
                 source):
        super().__init__(icon)
        self._on_toggle = on_toggle
        self._source = source

        self.setToolTip(app_name)

        menu = QMenu()
        menu.setToolTipsVisible(True)
        self._visibility_action = QAction("显示行情", self, triggered=self._on_toggle)
        menu.addAction(self._visibility_action)
        menu.addSeparator()
        self._toggles = _view_toggles(
            menu, source, lambda: "taskbar" if source.display_mode == "taskbar" else "float",
            include_taskbar=True,
        )
        menu.addSeparator()
        menu.addAction(QAction("设置…", self, triggered=on_open_settings))
        menu.addSeparator()
        for title, key in (("使用帮助", "readme"), ("问题反馈", "issues")):
            menu.addAction(QAction(title, self, triggered=lambda _checked=False, link=key:
                                  QDesktopServices.openUrl(QUrl(project_links()[link]))))
        menu.addSeparator()
        menu.addAction(QAction("退出", self, triggered=on_quit))
        menu.aboutToShow.connect(self.sync_settings)
        self.setContextMenu(menu)
        self.sync_settings()

        self.activated.connect(self._on_activated)

    def sync_settings(self):
        self._visibility_action.setText("隐藏行情" if self._source.widget_visible else "显示行情")
        _sync_toggles(self._toggles)

    def _on_activated(self, reason):
        # Windows 左键切换；macOS/Linux 单击即弹菜单，无切换逻辑
        if tray_click_toggles() and reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._on_toggle()


def build_quote_menu(source, surface="float"):
    menu = QMenu(source)
    menu.setToolTipsVisible(True)
    sub_cols = QMenu("显示指标", menu)
    sub_cols.setToolTipsVisible(True)
    metric_scope = _surface_scope(source, surface, "taskbar_sync_metrics")
    sub_cols.menuAction().setToolTip(metric_scope)
    metrics = source.get_surface_metrics(surface)
    for spec in METRIC_SPECS:
        action = QAction(spec.label, sub_cols, checkable=True)
        action.setChecked(spec.metric_id in metrics)
        action.setToolTip(metric_scope)
        action.toggled.connect(partial(source.set_surface_metric_visible, surface, spec.metric_id))
        sub_cols.addAction(action)
    menu.addMenu(sub_cols)

    sort_menu = QMenu("排序", menu)
    sort_menu.setToolTipsVisible(True)
    sort_menu.menuAction().setToolTip(
        "浮窗与任务栏共用排序；恢复自选顺序不会修改自选列表。"
        if sys.platform == "win32" else "恢复自选顺序不会修改自选列表。")
    visible_headers = metric_headers(metrics)
    sort_headers = [header for header in visible_headers if header in SORTABLE_HEADERS]
    current_sort = source.quotes.sort_header
    if current_sort in SORTABLE_HEADERS and current_sort not in sort_headers:
        sort_headers.append(current_sort)
    for header_name in sort_headers:
        title = header_name if header_name in visible_headers else f"{header_name}（当前排序，未显示）"
        metric_menu = QMenu(title, sort_menu)
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
    if sort_headers:
        sort_menu.addSeparator()
    clear_action = QAction("恢复自选顺序", sort_menu)
    clear_action.setEnabled(source.quotes.sort_header is not None)
    clear_action.triggered.connect(source.quotes.clear_sort)
    sort_menu.addAction(clear_action)
    if not sort_headers:
        sort_menu.setEnabled(False)
        sort_menu.menuAction().setToolTip("请先在“显示指标”中启用可排序的指标。")
    menu.addMenu(sort_menu)
    menu.addSeparator()

    if surface == "float":
        act_header = QAction("显示表头", menu, checkable=True)
        act_header.setChecked(source.header_visible)
        act_header.toggled.connect(source.set_header_visible)
        menu.addAction(act_header)

    act_grid = QAction("显示网格", menu, checkable=True)
    act_grid.setChecked(source.grid_visible)
    act_grid.setToolTip("同时应用于浮窗和任务栏。" if sys.platform == "win32" else "显示浮窗单元格边框。")
    act_grid.toggled.connect(source.set_grid_visible)
    menu.addAction(act_grid)

    act_color = QAction("统一颜色", menu, checkable=True)
    independent_color = surface == "taskbar" and not source.view_options.taskbar_sync_appearance
    if surface == "taskbar":
        auto_color = QAction("自动配色", menu, checkable=True)
        auto_color.setChecked(source.view_options.taskbar_auto_color)
        auto_color.setEnabled(independent_color)
        auto_color.setToolTip("跟随系统任务栏深浅色自动选择文字颜色。" if independent_color else
                             "任务栏正同步浮窗外观；在设置中关闭外观同步后可启用自动配色。")
        auto_color.toggled.connect(lambda enabled: source.set_view_options(taskbar_auto_color=enabled))
        menu.addAction(auto_color)
    act_color.setChecked(source.get_taskbar_appearance()[3] if surface == "taskbar" else source.unicolor)
    act_color.setToolTip(_surface_scope(source, surface, "taskbar_sync_appearance"))
    if independent_color:
        act_color.setEnabled(not source.view_options.taskbar_auto_color)
        if source.view_options.taskbar_auto_color:
            act_color.setToolTip("任务栏已开启自动配色；关闭自动配色后可修改统一颜色。")
        act_color.toggled.connect(lambda enabled: source.set_view_options(taskbar_unicolor=enabled))
    else:
        act_color.toggled.connect(source.set_unicolor)
    menu.addAction(act_color)

    _view_toggles(menu, source, surface)

    menu.addSeparator()
    act_open_settings = QAction("设置…", menu)
    if callable(source._open_settings_cb):
        act_open_settings.triggered.connect(source._open_settings_cb)
    else:
        act_open_settings.setEnabled(False)
    menu.addAction(act_open_settings)

    menu.addSeparator()
    menu.addAction(QAction("隐藏", menu, triggered=source.hide_widget))
    menu.addSeparator()
    menu.addAction(QAction("退出", menu, triggered=source._quit_cb or QApplication.instance().quit))
    return menu
