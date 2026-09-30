"""Shared vertical page controls for the float and native taskbar renderer."""

from PySide6.QtCore import QRect, Qt, Signal, QSize
from PySide6.QtGui import QColor, QFont, QFontMetrics, QPainter
from PySide6.QtWidgets import QWidget


def pager_regions(rect):
    h = max(1, rect.height() // 3)
    return (QRect(rect.x(), rect.y(), rect.width(), h),
            QRect(rect.x(), rect.y() + h, rect.width(), rect.height() - 2 * h),
            QRect(rect.x(), rect.bottom() - h + 1, rect.width(), h))


def pager_hit(rect, x, y):
    previous, _, following = pager_regions(rect)
    if previous.contains(x, y):
        return -1
    if following.contains(x, y):
        return 1
    return 0


def paint_pager(painter, rect, page, color, font):
    painter.save()
    painter.setClipRect(rect)
    f = QFont(font)
    pixels = font.pixelSize() if font.pixelSize() > 0 else QFontMetrics(font).height()
    f.setPixelSize(max(7, min(round(rect.height() / 4.2), pixels)))
    label = f"{page.index + 1}/{page.count}"
    while f.pixelSize() > 7 and QFontMetrics(f).horizontalAdvance(label) > rect.width() - 2:
        f.setPixelSize(f.pixelSize() - 1)
    painter.setFont(f)
    painter.setPen(color)
    for area, text in zip(pager_regions(rect), ("︿", label, "﹀")):
        painter.drawText(area, Qt.AlignCenter, text)
    painter.restore()


class PagerWidget(QWidget):
    page_requested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.page = None
        self.color = QColor("white")
        self.setFixedWidth(38)
        self.setMinimumHeight(42)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("单击上方 / 下方翻页；按住任意位置拖动；双击隐藏")
        self.hide()

    def sizeHint(self):
        return QSize(38, 42)

    def set_page(self, page, color, font):
        self.page, self.color = page, QColor(color)
        self.setFont(font)
        self.setVisible(page.controls)
        self.update()

    def paintEvent(self, event):
        if self.page:
            painter = QPainter(self)
            paint_pager(painter, self.rect(), self.page, self.color, self.font())

    def activate_at(self, position):
        """由宿主的共享拖拽过滤器确认未拖动后调用。"""
        delta = pager_hit(self.rect(), position.x(), position.y())
        if delta:
            self.page_requested.emit(delta)
