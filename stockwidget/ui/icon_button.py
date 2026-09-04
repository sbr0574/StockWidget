"""设置页自定义图标选择按钮。"""

from PySide6.QtCore import QRectF, Qt, Signal
from PySide6.QtGui import QColor, QIcon, QPainter, QPen
from PySide6.QtWidgets import QPushButton


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
