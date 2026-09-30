"""共用 Qt 场景夹具；隔离后台请求与原生操作，并确定性释放窗口。"""

from unittest.mock import Mock, patch
import os
import threading
import unittest

from PySide6.QtCore import Qt, QPointF
from PySide6.QtGui import QMouseEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QAbstractItemDelegate
from shiboken6 import delete

from stockwidget.platform.taskbar import TaskbarArea
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import TaskbarController
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.settings.dialog import SettingsDialog
from stockwidget.ui.watchlist.editor import CodeSearchEditor


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


CODES = {
    "sh600519": {
        "code": "600519",
        "market": "sh",
        "name": "贵州茅台",
        "type": "沪",
        "py": "guizhoumaotai",
        "abbr": "gzmt",
    },
    "sh501001": {
        "code": "501001",
        "market": "sh",
        "name": "财通精选混合LOF",
        "type": "基",
        "py": "caitongjingxuanhunhelof",
        "abbr": "ctjxhhlof",
    },
    "sh000001": {
        "code": "000001",
        "market": "sh",
        "name": "上证指数",
        "type": "指",
        "py": "shangzhengzhishu",
        "abbr": "szzs",
    },
    "ad0": {
        "code": "ad0",
        "market": "",
        "name": "铸造铝合金连续",
        "type": "期",
        "py": "zhuzaolvhejinlianxu",
        "abbr": "zzlhjlx",
    },
    "longname": {
        "code": "longname",
        "market": "us",
        "name": "用于验证候选列表根据条目内容自动扩展宽度的特别长证券名称",
        "type": "美",
        "py": "yongyuyanzhengchaotezhengquanmingcheng",
        "abbr": "cctm",
    },
}


class SignalRecorder:
    def __init__(self):
        self.event = threading.Event()
        self.value = None

    def emit(self, value):
        self.value = value
        self.event.set()


class QtTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = cls.qt_app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)


class SettingsTestCase(QtTestCase):

    def setUp(self):
        self._windows = []
        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.floating.widget.apply_click_through"))
        self.enterContext(patch("stockwidget.ui.floating.widget.ensure_topmost"))

    def tearDown(self):
        for dialog, window in reversed(self._windows):
            for editor in dialog.watchlist_editor.list_codes.findChildren(CodeSearchEditor):
                editor.setProperty("_code_editor_committed", True)
                editor._code_completer.popup().hide()
                dialog.watchlist_editor.list_codes.closeEditor(
                    editor, QAbstractItemDelegate.EndEditHint.NoHint
                )
            self.qt_app.processEvents()
            dialog.close()
            window.close()
            delete(dialog)
            delete(window)
        self.qt_app.processEvents()

    def _make_dialog(self, watchlist=None, codes=None, app=None, **settings):
        cfg = {"watchlist": watchlist or {}, **settings}
        window = FloatLabel(cfg, CODES if codes is None else codes)
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(window, window, app=app)
        self._windows.append((dialog, window))
        return dialog, window

    def _click_info(self, dialog, pool, item):
        dialog.show()
        self.qt_app.processEvents()
        pool.scrollToItem(item)
        QTest.mouseClick(pool.viewport(), Qt.LeftButton, pos=pool.info_rect(item).center())

    def _start_code_editor(self, dialog):
        dialog.watchlist_editor._start_quick_add()
        self.qt_app.processEvents()
        editor = dialog.watchlist_editor.list_codes.findChild(CodeSearchEditor)
        self.assertIsNotNone(editor)
        return editor


class PagingTestCase(QtTestCase):

    def setUp(self):
        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.floating.widget.apply_click_through"))
        self.win = FloatLabel({"taskbar_enabled": True, "taskbar_sync_metrics": False,
                               "float_paging_enabled": True, "taskbar_sync_paging": False}, {})
        self.win._clear_message()
        self.controller = None
        self.populate()

    def populate(self, total=5):
        self.win.quotes._last_full_rows = [dict(zip(("名称", "现价", "涨幅", "涨跌"),
                                                   (f"标的{i}", str(i * 10), f"+{i}%", str(i)))) for i in range(1, total + 1)]
        self.win.quotes._last_color_roles = [{h: "up" for h in row} for row in self.win.quotes._last_full_rows]
        self.win.quotes._last_sort_values = [{"现价": i * 10} for i in range(1, total + 1)]
        self.win.quotes.reproject()

    def tearDown(self):
        if self.controller:
            self.controller.close()
            delete(self.controller)
        delete(self.win)
        self.app.processEvents()

    def mouse(self, target, kind, global_pos, button, buttons):
        local = target.mapFromGlobal(global_pos)
        event = QMouseEvent(kind, QPointF(local), QPointF(global_pos), button, buttons, Qt.NoModifier)
        QApplication.sendEvent(target, event)

    def make_controller(self):
        self.enterContext(patch("stockwidget.ui.floating.taskbar.QApplication.platformName", return_value="windows"))
        self.enterContext(patch("stockwidget.ui.floating.widget.sys.platform", "win32"))
        self.enterContext(patch("stockwidget.ui.floating.taskbar.find_taskbar", return_value=TaskbarArea(1, 1280, 48, 900)))
        self.native = Mock()
        self.native.poll_pointer.return_value = None
        self.enterContext(patch("stockwidget.ui.floating.taskbar.NativeTaskbarWindow", return_value=self.native))
        self.over = self.enterContext(patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=False))
        self.controller = TaskbarController(self.win, Mock())
        self.controller.apply_mode()
        return self.controller
