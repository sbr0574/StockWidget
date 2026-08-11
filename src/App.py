import sys, os, threading
from datetime import datetime

# 先判断系统类型,再按需 import 平台相关模块
if sys.platform == "win32":
    import winreg

from PySide6.QtCore import Qt, QPoint, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QStyle, QMessageBox
import resources.resources_rc
from src.WidgetPanel import FloatLabel
from src.SettingPanel import SettingsDialog
from services.code_index import refresh_index_from_akshare
from services.update_check import check_for_update
from src.utils import load_file, save_file, load_json_from_resource
from src.platform_support import click_through_supported, start_on_boot_supported

APP_NAME = "StockWidget"
APP_VERSION = "1.3.1"
CONFIG_FILE = "stock_widget_config.json"
LIST_FILE = "stock_codes_list.json"

class App(QApplication):
    update_checked = Signal(bool)
    index_progress = Signal(int)
    index_finished = Signal(dict)

    def __init__(self, argv):
        super().__init__(argv)
        self.app_name = APP_NAME
        self.app_version = APP_VERSION

        self.setQuitOnLastWindowClosed(False)
        cfg = load_file(self.app_name, CONFIG_FILE)

        # 加载市场代码列表：更新日期为今天则直接读取，否则先用资源缓存启动，后台再更新
        self._need_background_refresh = False
        codes_list = None
        list_file = load_file(self.app_name, LIST_FILE)
        list_last_update = list_file.get("last_update")
        if list_last_update:
            try:
                stale = datetime.now().date() > datetime.strptime(list_last_update, "%Y-%m-%d").date()
            except (ValueError, TypeError):
                stale = True
            if not stale:
                codes_list = list_file.get("codes")
        if not codes_list:
            resource_list = load_json_from_resource(":/stock_codes_list.json")
            codes_list = (resource_list or {}).get("codes")
            self._need_background_refresh = True

        # 加载图标
        self._icon_choice = cfg.get('app_icon')
        app_icon = self.find_icon(self._icon_choice)
        self.setWindowIcon(app_icon)

        # 初始化浮窗
        self.win = FloatLabel(cfg, codes_list)
        self._start_on_boot = bool(cfg.get("start_on_boot", False))
        self.set_start_on_boot(self._start_on_boot) # Apply start-on-boot setting from config
        self.win.set_on_change(self.save_now)
        self.win.set_open_settings_callback(self.open_settings)

        # 初始化托盘
        self.tray = QSystemTrayIcon(app_icon, self)
        self.tray.setToolTip(APP_NAME)
        menu = QMenu()
        menu.addAction(QAction("显示/隐藏 浮窗", self, triggered=self.toggle_win))
        self.act_click_through = QAction("鼠标穿透", self, checkable=True)
        self.act_click_through.setChecked(self.win.click_through)
        self.act_click_through.toggled.connect(self.win.set_click_through)
        if not click_through_supported():
            # 当前平台(如 Wayland)不支持鼠标穿透,置为不可点按
            self.act_click_through.setEnabled(False)
            self.act_click_through.setToolTip("当前会话不支持鼠标穿透")
        menu.addAction(self.act_click_through)
        menu.addAction(QAction("设置…", self, triggered=self.open_settings))
        menu.addSeparator()
        menu.addAction(QAction("退出", self, triggered=self.quit_app))
        menu.aboutToShow.connect(self._sync_tray_click_through)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()

        # 启动浮窗
        self.settings_dlg = None
        self.win.show()
        self.win.raise_()
        self.win.activateWindow()
        self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

        # 启动时后台检查更新
        self._has_update = False
        self.update_checked.connect(self._on_update_checked)
        self._start_update_check()

        # 启动时后台更新市场代码列表（不阻塞启动）
        if self._need_background_refresh:
            self._start_refresh_index()

    def find_icon(self, choice: str) -> QIcon:
        if not choice or choice == 'default':
            return QIcon(":/StockWidget.ico")
        if isinstance(choice, str) and choice.startswith('std:'):
            key = choice.split(':',1)[1]
            mapping = {
                'computer': QStyle.SP_ComputerIcon,
                'network': QStyle.SP_DriveNetIcon,
                'folder': QStyle.SP_DirIcon,
                'file': QStyle.SP_FileIcon,
                'trash': QStyle.SP_TrashIcon,
                'desktop': QStyle.SP_DesktopIcon,
            }
            sp = mapping.get(key, QStyle.SP_ComputerIcon)
            return self.style().standardIcon(sp)
        try:
            if os.path.exists(choice):
                return QIcon(choice)
        except Exception:
            return QIcon(":/StockWidget.ico")

    def on_tray_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick): self.toggle_win()

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

    def _sync_tray_click_through(self):
        if hasattr(self, "act_click_through") and hasattr(self, "win"):
            self.act_click_through.setChecked(self.win.click_through)

    def _start_update_check(self):
        """后台线程检查更新，结果通过信号回传到主线程显示"""
        def _worker():
            try:
                has_update = check_for_update(self.app_version)
            except Exception:
                has_update = False
            self.update_checked.emit(bool(has_update))

        threading.Thread(target=_worker, daemon=True).start()

    def _start_refresh_index(self):
        """后台线程拉取市场代码，进度通过信号回传到主线程显示"""
        self.index_progress.connect(self._on_index_progress)
        self.index_finished.connect(self._on_index_finished)
        self.win.set_index_updating(True)
        self.win._show_message("正在更新市场代码数据 ( 0% )")

        def _worker():
            try:
                result = refresh_index_from_akshare(progress_cb=self.index_progress.emit)
                save_file(result, self.app_name, LIST_FILE)
            except Exception:
                result = None
            codes = result.get("codes") if result is not None else None
            self.index_finished.emit(codes or {})

        threading.Thread(target=_worker, daemon=True).start()

    def _on_index_progress(self, percent: int):
        self.win._show_message(f"正在更新市场代码数据 ( {percent}% )")

    def _on_index_finished(self, codes: dict):
        self.win.set_index_updating(False)
        if codes:
            self.win.codes_list = codes
        self.win._clear_message()

    def _on_update_checked(self, has_update: bool):
        self._has_update = has_update
        if self.settings_dlg is not None and self.settings_dlg.isVisible():
            try:
                self.settings_dlg.refresh_about()
            except Exception:
                pass

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
        """Set application and tray icon. `choice` can be None/'default', 'std:KEY' or a file path."""
        self._icon_choice = choice
        app_icon = self.find_icon(self._icon_choice)
        self.setWindowIcon(app_icon)
        self.tray.setIcon(app_icon)

    def set_start_on_boot(self, enabled: bool):
        """Enable or disable auto-start on login.
        - Windows: 写/删 HKCU Run 注册表键。
        - Linux:   写/删 XDG autostart .desktop 文件(桌面环境通用)。
        - macOS:   写/删 ~/Library/LaunchAgents 下的 LaunchAgent plist。
        """
        self._start_on_boot = enabled
        system = sys.platform
        if system == "win32":
            self._set_start_on_boot_windows(enabled)
        elif system == "linux":
            self._set_start_on_boot_linux(enabled)
        elif system == "darwin":
            self._set_start_on_boot_macos(enabled)

    def _set_start_on_boot_windows(self, enabled: bool):
        """Windows:通过 HKCU Run 注册表键实现开机自启。"""
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            name = APP_NAME
            if enabled:
                if getattr(sys, 'frozen', False):
                    cmd = f'"{sys.executable}"'
                else:
                    cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, cmd)
            else:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, name)
        except Exception:
            pass

    def _set_start_on_boot_linux(self, enabled: bool):
        """Linux:通过 XDG autostart 目录(~/.config/autostart)下的 .desktop 文件实现开机自启。
        GNOME/KDE/XFCE 等主流桌面环境均支持该机制。
        """
        try:
            autostart_dir = os.path.join(os.path.expanduser("~"), ".config", "autostart")
            desktop_file = os.path.join(autostart_dir, f"{APP_NAME}.desktop")
            if not enabled:
                if os.path.exists(desktop_file):
                    os.remove(desktop_file)
                return
            if getattr(sys, 'frozen', False):
                exec_cmd = f'"{sys.executable}"'
            else:
                exec_cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
            content = (
                "[Desktop Entry]\n"
                "Type=Application\n"
                f"Name={APP_NAME}\n"
                f"Comment={APP_NAME} 透明盯盘浮窗\n"
                f"Exec={exec_cmd}\n"
                "Terminal=false\n"
                "X-GNOME-Autostart-enabled=true\n"
            )
            os.makedirs(autostart_dir, exist_ok=True)
            with open(desktop_file, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception:
            pass

    def _set_start_on_boot_macos(self, enabled: bool):
        """macOS:通过 LaunchAgent(~/Library/LaunchAgents)实现开机自启。

        写入 com.sbr0574.StockWidget.plist(RunAtLoad = true),登录时由 launchd
        在用户图形会话中自动拉起;并调用 launchctl 立即加载/卸载。
        """
        try:
            import plistlib
            import subprocess
            label = "com.sbr0574.StockWidget"
            launch_dir = os.path.join(os.path.expanduser("~"), "Library", "LaunchAgents")
            plist_path = os.path.join(launch_dir, f"{label}.plist")
            uid = str(os.getuid())
            if not enabled:
                if os.path.exists(plist_path):
                    os.remove(plist_path)
                # 立即卸载(失败忽略,注销后也会自然消失)
                subprocess.run(
                    ["launchctl", "bootout", f"gui/{uid}", plist_path],
                    capture_output=True, timeout=10,
                )
                return
            if getattr(sys, 'frozen', False):
                args = [sys.executable]
            else:
                args = [sys.executable, os.path.abspath(sys.argv[0])]
            payload = {
                "Label": label,
                "ProgramArguments": args,
                "WorkingDirectory": os.path.dirname(os.path.abspath(sys.argv[0])),
                "RunAtLoad": True,
                "ProcessType": "Interactive",
            }
            os.makedirs(launch_dir, exist_ok=True)
            with open(plist_path, "wb") as f:
                plistlib.dump(payload, f)
            # 立即加载(若已在运行,launchctl 会报错,忽略即可)
            subprocess.run(
                ["launchctl", "bootstrap", f"gui/{uid}", plist_path],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass
