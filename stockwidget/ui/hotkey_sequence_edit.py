"""在全局快捷键已注册时，仍保留输入框中显示的组合键。"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QKeySequenceEdit


class HotkeySequenceEdit(QKeySequenceEdit):
    # 系统可能截获已注册组合的主键，只把 Ctrl/Alt 等事件交给输入框。
    # 等收到实际主键再开始编辑，避免仅按修饰键就清空已有组合。
    _MODIFIERS = {Qt.Key_Control, Qt.Key_Shift, Qt.Key_Alt, Qt.Key_Meta, Qt.Key_AltGr}

    def keyPressEvent(self, event):
        if event.key() in self._MODIFIERS:
            event.accept()
            return
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        if event.key() in self._MODIFIERS:
            event.accept()
            return
        super().keyReleaseEvent(event)
