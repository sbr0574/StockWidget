"""浮窗表头的排序箭头，颜色跟随表头文字，不受系统主题影响。"""

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QPainter, QPalette, QPolygonF
from PySide6.QtWidgets import QApplication, QProxyStyle, QStyle, QStyleOptionHeader


class SortIndicatorStyle(QProxyStyle):
    def __init__(self, header):
        # 按名称创建独立样式，避免代理接管 QApplication 共享样式的所有权。
        super().__init__(QApplication.style().objectName())
        self.setParent(header)

    def sizeFromContents(self, contents_type, option, size, widget=None):
        if contents_type == QStyle.CT_HeaderSection:
            # QHeaderView 会为所有列附加排序标记；尺寸计算忽略它，避免整表变宽。
            option = QStyleOptionHeader(option)
            option.sortIndicator = QStyleOptionHeader.SortIndicator.None_
        return super().sizeFromContents(contents_type, option, size, widget)

    def drawPrimitive(self, element, option, painter, widget=None):
        if element != QStyle.PE_IndicatorHeaderArrow:
            return super().drawPrimitive(element, option, painter, widget)

        rect = QRectF(option.rect)
        center = rect.center()
        half_width = min(8.0, rect.width()) / 2
        half_height = min(5.0, rect.height()) / 2
        # Qt 的 SortUp 表示降序，箭头尖端应朝下。
        direction = 1 if option.sortIndicator == QStyleOptionHeader.SortUp else -1
        base_y = center.y() - direction * half_height
        arrow = QPolygonF([
            QPointF(center.x() - half_width, base_y),
            QPointF(center.x() + half_width, base_y),
            QPointF(center.x(), center.y() + direction * half_height),
        ])
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(option.palette.brush(QPalette.ButtonText))
        painter.drawPolygon(arrow)
        painter.restore()
