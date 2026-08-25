# -*- coding: utf-8 -*-
"""应用装配层：创建各层对象、连接信号、处理托盘与后台任务。

本模块只负责“组装”各层，不含界面细节（见 ui/）与数据/网络细节（见 data/）。
"""

import os
import sys
import threading

from PySide6.QtCore import Qt, QPoint, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QStyle, QMessageBox

import resources.resources_rc  # noqa: F401  加载 Qt 内嵌资源（图标、内置代码列表）

from stockwidget.constants import APP_NAME, APP_VERSION, CONFIG_FILE
from stockwidget.core.config_store import load_file, save_file
from stockwidget.data.code_lists import (
    all_codes_fresh,
    code_data_state,
    download_codes,
    load_local_codes,
    load_resource_codes,
)
from stockwidget.data.update_check import GITEE, GITHUB, get_update_info, github_available
from stockwidget.platform.autostart import set_start_on_boot
from stockwidget.ui.settings_dialog import SettingsDialog
from stockwidget.ui.tray import TrayIcon
from stockwidget.ui.widget import FloatLabel


class App(QApplication):
    network_finished = Signal(object)

    def __init__(self, argv):
        super().__init__(argv)
        self.app_name = APP_NAME
        self.app_version = APP_VERSION
        self._remote_source = GITHUB

        self.setQuitOnLastWindowClosed(False)
        cfg = load_file(self.app_name, CONFIG_FILE)

        # 加载市场代码列表：更新日期为今天则直接读取，否则先用资源缓存启动，后台再更新
        self._need_background_refresh = False
        local_codes = load_local_codes(self.app_name)
        if local_codes and all_codes_fresh(self.app_name):
            codes_list = local_codes
        else:
            # 本地缺失/过期：先用资源内嵌兜底显示，启动后从可用托管源下载
            codes_list = load_resource_codes() or local_codes
            self._need_background_refresh = True

        # 加载图标
        self._icon_choice = cfg.get('app_icon')
        app_icon = self.find_icon(self._icon_choice)
        self.setWindowIcon(app_icon)

        # 初始化浮窗
        self.win = FloatLabel(cfg, codes_list)
        self._start_on_boot = bool(cfg.get("start_on_boot", False))
        self.set_start_on_boot(self._start_on_boot)  # 应用配置中的开机自启
        self.win.set_on_change(self.save_now)
        self.win.set_open_settings_callback(self.open_settings)

        # 初始化托盘（菜单与点击行为封装在 ui/tray.py，按平台区分）
        self.tray = TrayIcon(
            app_icon,
            APP_NAME,
            on_toggle=self.toggle_win,
            on_open_settings=self.open_settings,
            on_quit=self.quit_app,
            on_click_through=self.win.set_click_through,
            click_through_getter=lambda: self.win.click_through,
        )
        self.tray.show()

        # 启动浮窗
        self.settings_dlg = None
        self.win.show()
        self.win.raise_()
        self.win.activateWindow()
        self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

        # 启动时在同一后台任务中探测代码托管源、检查更新并按需刷新代码表
        self._has_update = False
        self._latest_version = None
        self._latest_release_url = None
        self.network_finished.connect(self._on_network_finished)
        self._start_network_tasks()

    def find_icon(self, choice: str) -> QIcon:
        if choice == 'lightG':
            return QIcon(":/LightGlass.ico")
        if choice == 'darkG':
            return QIcon(":/DarkGlass.ico")
        if choice == 'dark':
            return QIcon(":/DarkSW.ico")
        return QIcon(":/StockWidget.ico")

    def toggle_win(self):
        if self.win.isVisible():
            self.win.hide()
        else:
            self.win.show()
            self.win.raise_()
            self.win.activateWindow()
            self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

    def open_settings(self):
        if self.settings_dlg and self.settings_dlg.isVisible():
            self.settings_dlg.raise_()
            self.settings_dlg.activateWindow()
            return
        try:
            self.settings_dlg = SettingsDialog(self.win, self.win, app=self)
        except Exception as exc:
            QMessageBox.critical(None, "设置窗口错误", f"无法打开设置窗口：\n{exc}")
            return
        # 将设置窗口放在屏幕正中
        screen = QApplication.primaryScreen().availableGeometry()
        self.settings_dlg.adjustSize()
        cx = screen.left() + (screen.width() - self.settings_dlg.width()) // 2
        cy = screen.top() + (screen.height() - self.settings_dlg.height()) // 2
        self.settings_dlg.move(QPoint(cx, cy))
        self.settings_dlg.show()
        self.settings_dlg.raise_()
        self.settings_dlg.activateWindow()

    def _start_network_tasks(self):
        """后台选择 GitHub/Gitee，并统一执行代码下载与版本检查。"""
        if self._need_background_refresh:
            self.win.set_index_updating(True)
            self.win._show_message("正在更新市场代码数据…")

        def _worker():
            source = GITHUB if github_available() else GITEE
            codes = None
            if self._need_background_refresh:
                try:
                    codes = download_codes(self.app_name, source)
                except Exception:
                    codes = None
                if not codes and source == GITHUB:
                    source = GITEE
                    try:
                        codes = download_codes(self.app_name, source)
                    except Exception:
                        codes = None
            try:
                has_update, latest_version, latest_url = get_update_info(self.app_version, source)
            except Exception:
                has_update, latest_version, latest_url = False, None, None
            self.network_finished.emit({
                "source": source,
                "codes": codes or {},
                "has_update": bool(has_update),
                "latest_version": latest_version,
                "latest_url": latest_url,
            })

        threading.Thread(target=_worker, daemon=True).start()

    def _on_network_finished(self, result: dict):
        self._remote_source = result.get("source") or GITHUB
        self._has_update = bool(result.get("has_update"))
        self._latest_version = result.get("latest_version") if self._has_update else None
        self._latest_release_url = result.get("latest_url") if self._has_update else None
        codes = result.get("codes") or {}
        if codes:
            self.win.set_codes_list(codes)
        if self._need_background_refresh:
            self.win.set_index_updating(False)
            self.win._clear_message()
        if self.settings_dlg is not None and self.settings_dlg.isVisible():
            try:
                self.settings_dlg.refresh_data_state()
                self.settings_dlg.refresh_about()
            except Exception:
                pass

    def code_data_state(self) -> tuple:
        """市场代码数据状态与更新日期：('online'|'cached'|'offline', 'YYYY-MM-DD')。"""
        return code_data_state(self.app_name)

    def quit_app(self):
        self.tray.hide()
        self.save_now()
        sys.exit(0)

    def save_now(self):
        cfg = self.win.current_config()
        cfg['app_icon'] = self._icon_choice
        cfg['start_on_boot'] = self._start_on_boot
        save_file(cfg, self.app_name, CONFIG_FILE)

    def set_app_icon(self, choice):
        """Set application and tray icon."""
        self._icon_choice = choice
        app_icon = self.find_icon(self._icon_choice)
        self.setWindowIcon(app_icon)
        self.tray.setIcon(app_icon)

    def set_start_on_boot(self, enabled: bool):
        """启用或禁用开机自启（Windows/Linux/macOS），由平台支持层统一实现。"""
        self._start_on_boot = enabled
        set_start_on_boot(enabled, APP_NAME)
