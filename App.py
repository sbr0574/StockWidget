import importlib
import sys, os, platform

SYSTEM = platform.system()
if SYSTEM == "Windows":
    import winreg
    keyboard = importlib.import_module("keyboard")
elif SYSTEM == "Darwin":
    keyboard = None
else:
    keyboard = None

from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QStyle
from WidgetPanel import FloatLabel
from SettingPanel import SettingsDialog
from config_store import load_config, save_config
from code_index import refresh_index_from_akshare

# ----- 程序与资源 -----
APP_NAME = "StockWidget"
APP_ICON_FILE = "StockWidget.ico"

def resource_path(rel_path):
    base = getattr(sys, "_MEIPASS", "")
    return os.path.join(base, rel_path)

class App(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.system = SYSTEM
        self.setQuitOnLastWindowClosed(False)
        icon_path = resource_path(APP_ICON_FILE)
        cfg = load_config(APP_NAME)
        self.code_index = refresh_index_from_akshare(APP_NAME)
        icon_choice = cfg.get('app_icon')
        self._app_icon_choice = icon_choice
        def _resolve_icon(choice):
            # choice can be None, 'default', 'std:NAME' or a file path
            if not choice or choice == 'default':
                p = resource_path(APP_ICON_FILE)
                if os.path.exists(p):
                    return QIcon(p)
                return self.style().standardIcon(QStyle.SP_ComputerIcon)
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
            # assume it's a file path
            try:
                if os.path.exists(choice):
                    return QIcon(choice)
            except Exception:
                pass
            return self.style().standardIcon(QStyle.SP_ComputerIcon)

        app_icon = _resolve_icon(icon_choice)
        self.setWindowIcon(app_icon)

        self.win = FloatLabel(cfg)
        # Apply start-on-boot setting from config
        try:
            self.set_start_on_boot(bool(cfg.get("start_on_boot", False)))
        except Exception:
            pass
        self.win.set_on_change(self.save_now)
        self.win.set_open_settings_callback(self.open_settings)

        self.tray = QSystemTrayIcon(app_icon, self)
        self.tray.setToolTip(APP_NAME)
        menu = QMenu()
        menu.addAction(QAction("显示/隐藏 浮窗", self, triggered=self.toggle_win))
        menu.addAction(QAction("设置…", self, triggered=self.open_settings))
        menu.addSeparator()
        menu.addAction(QAction("退出", self, triggered=self.quit_app))
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()

        self.settings_dlg = None
        self.win.show()
        self.win.raise_()
        self.win.activateWindow()
        self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

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
            from PySide6.QtWidgets import QMessageBox
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

    def quit_app(self):
        self.tray.hide()
        self.save_now()
        if keyboard is not None and hasattr(keyboard, "unhook_all_hotkeys"):
            try:
                keyboard.unhook_all_hotkeys()
            except Exception:
                pass
        sys.exit(0)

    def save_now(self):
        cfg = self.win.current_config()
        # persist selected app icon
        try:
            cfg['app_icon'] = getattr(self, '_app_icon_choice', None)
        except Exception:
            pass
        save_config(APP_NAME, cfg)

    def set_app_icon(self, choice):
        """Set application and tray icon. `choice` can be None/'default', 'std:KEY' or a file path."""
        self._app_icon_choice = choice
        # resolve to QIcon
        def _resolve_icon(choice):
            if not choice or choice == 'default':
                p = resource_path(APP_ICON_FILE)
                if os.path.exists(p):
                    return QIcon(p)
                return self.style().standardIcon(QStyle.SP_ComputerIcon)
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
                pass
            return self.style().standardIcon(QStyle.SP_ComputerIcon)

        icon = _resolve_icon(choice)
        try:
            self.setWindowIcon(icon)
        except Exception:
            pass
        try:
            if hasattr(self, 'tray') and self.tray is not None:
                self.tray.setIcon(icon)
        except Exception:
            pass

    def set_start_on_boot(self, enabled: bool):
        """Enable or disable Windows startup by writing/removing Run key in HKCU.
        On macOS/Linux this is ignored gracefully.
        """
        if self.system != "win32":
            return
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
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.DeleteValue(key, name)
                except OSError:
                    pass
        except Exception:
            pass
