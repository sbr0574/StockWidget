# -*- coding: utf-8 -*-
"""系统托盘封装：统一 Windows / macOS / Linux 的菜单与点击行为。

平台差异：
- Windows：左键单击/双击切换显示隐藏，右键弹出菜单。
- macOS / Linux：单击（不区分左右键）直接弹出菜单，无左键切换逻辑。
"""

from PySide6.QtGui import QAction
from PySide6.QtWidgets import QMenu, QSystemTrayIcon

from stockwidget.platform.capabilities import click_through_supported, tray_click_toggles


class TrayIcon(QSystemTrayIcon):
    """托盘图标 + 右键菜单 + 平台感知的点击行为。"""

    def __init__(self, icon, app_name, *,
                 on_toggle, on_open_settings, on_quit,
                 on_click_through, click_through_getter):
        super().__init__(icon)
        self._on_toggle = on_toggle
        self._on_click_through = on_click_through
        self._click_through_getter = click_through_getter

        self.setToolTip(app_name)

        menu = QMenu()
        menu.addAction(QAction("显示/隐藏 浮窗", self, triggered=self._on_toggle))

        self.act_click_through = QAction("鼠标穿透", self, checkable=True)
        self.act_click_through.setChecked(bool(self._click_through_getter()))
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

    def _on_activated(self, reason):
        # Windows 左键切换；macOS/Linux 单击即弹菜单，无切换逻辑
        if tray_click_toggles() and reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._on_toggle()
