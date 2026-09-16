"""快捷键注册状态：无描边圆形底色和白色勾/叉。"""

from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QWidget


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
