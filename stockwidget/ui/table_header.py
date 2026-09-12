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
        # 原生箭头锚定列边界；改在文字绘制后按实际文字位置绘制。
        if element != QStyle.PE_IndicatorHeaderArrow:
            return super().drawPrimitive(element, option, painter, widget)

    def drawItemText(self, painter, rect, flags, palette, enabled, text, text_role=QPalette.NoRole):
        super().drawItemText(painter, rect, flags, palette, enabled, text, text_role)
        header = self.parent()
        if (
            text_role != QPalette.ButtonText
            or not header.isSortIndicatorShown()
            or header.logicalIndexAt(rect.center()) != header.sortIndicatorSection()
        ):
            return

        # 使用样式表已解析的字体与对齐方式；列变宽时，箭头仍紧跟文字。
        text_rect = self.itemTextRect(painter.fontMetrics(), rect, flags, enabled, text)
        arrow_rect = QRectF(text_rect.right() + 2, text_rect.center().y() - 2, 5, 4)
        center = arrow_rect.center()
        half_width = arrow_rect.width() / 2
        half_height = arrow_rect.height() / 2
        direction = 1 if header.sortIndicatorOrder() == Qt.DescendingOrder else -1
        base_y = center.y() - direction * half_height
        arrow = QPolygonF([
            QPointF(center.x() - half_width, base_y),
            QPointF(center.x() + half_width, base_y),
            QPointF(center.x(), center.y() + direction * half_height),
        ])
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(palette.brush(text_role))
        painter.drawPolygon(arrow)
        painter.restore()
