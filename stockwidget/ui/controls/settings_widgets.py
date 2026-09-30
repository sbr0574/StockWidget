"""Designer 设置控件：快捷键输入与状态、自定义图标选择。"""

from PySide6.QtCore import Qt, QRectF, QPointF, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QIcon
from PySide6.QtWidgets import QKeySequenceEdit, QWidget, QPushButton


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


class HotkeyStatus(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)
        self.active = False

    def set_result(self, result, message):
        self.active = bool(result)
        self.setToolTip(message)
        self.setAccessibleName(message)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor("#28a745" if self.active else "#dc3545"))
        painter.drawEllipse(QRectF(1, 1, 18, 18))
        painter.setPen(QPen(Qt.white, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        if self.active:
            painter.drawLine(QPointF(5.5, 10), QPointF(8.5, 13))
            painter.drawLine(QPointF(8.5, 13), QPointF(14.5, 7))
        else:
            painter.drawLine(QPointF(6.5, 6.5), QPointF(13.5, 13.5))
            painter.drawLine(QPointF(13.5, 6.5), QPointF(6.5, 13.5))


class CustomIconButton(QPushButton):
    """显示自定义图标，并在悬停时提供右上角删除热区。"""

    deleteRequested = Signal()
    _DELETE_SIZE = 16

    def __init__(self, parent=None):
        super().__init__(parent)
        self._has_custom_icon = False
        self._delete_hovered = False
        self._delete_pressed = False
        self.setMouseTracking(True)

    def set_custom_icon(self, icon: QIcon):
        self._has_custom_icon = not icon.isNull()
        self.setIcon(icon if self._has_custom_icon else QIcon())
        self.setText("" if self._has_custom_icon else "+")
        self.update()

    def clear_custom_icon(self):
        self.set_custom_icon(QIcon())

    def has_custom_icon(self) -> bool:
        return self._has_custom_icon

    def delete_rect(self) -> QRectF:
        size = self._DELETE_SIZE
        return QRectF(self.width() - size, 0, size, size)

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._has_custom_icon or not self.underMouse():
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        circle = self.delete_rect().adjusted(0.75, 0.75, -0.75, -0.75)
        if self._delete_pressed:
            fill = QColor(190, 38, 32)
        elif self._delete_hovered:
            fill = QColor(255, 69, 58)
        else:
            fill = QColor(225, 52, 45)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(fill)
        painter.drawEllipse(circle)

        inset = 4.5
        cross = circle.adjusted(inset, inset, -inset, -inset)
        painter.setPen(QPen(Qt.GlobalColor.white, 1.5, Qt.PenStyle.SolidLine,
                            Qt.PenCapStyle.RoundCap))
        painter.drawLine(cross.topLeft(), cross.bottomRight())
        painter.drawLine(cross.topRight(), cross.bottomLeft())
        painter.end()

    def mouseMoveEvent(self, event):
        hovered = self._has_custom_icon and self.delete_rect().contains(event.position())
        if hovered != self._delete_hovered:
            self._delete_hovered = hovered
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        self._delete_hovered = False
        self._delete_pressed = False
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if (self._has_custom_icon
                and event.button() == Qt.MouseButton.LeftButton
                and self.delete_rect().contains(event.position())):
            self._delete_pressed = True
            self.update()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._delete_pressed:
            should_delete = (
                event.button() == Qt.MouseButton.LeftButton
                and self.delete_rect().contains(event.position())
            )
            self._delete_pressed = False
            self.update()
            event.accept()
            if should_delete:
                self.deleteRequested.emit()
            return
        super().mouseReleaseEvent(event)
