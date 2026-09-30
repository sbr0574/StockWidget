"""A single grid layer for the whole quote table, including its header and candles."""

from PySide6.QtCore import QPoint, QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QTableView, QWidget


class _TableGrid(QWidget):
    def __init__(self, table):
        super().__init__(table)
        self.table = table
        self.color = QColor()
        # The grid is only paint; hit-testing and dragging reach the usual regions.
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setFocusPolicy(Qt.NoFocus)
        self.hide()

    def paintEvent(self, event):
        table, model = self.table, self.table.model()
        if model is None or self.width() < 2 or self.height() < 2:
            return
        bounds = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        path = QPainterPath()
        path.addRoundedRect(bounds, 3, 3)
        viewport = table.viewport()
        origin = viewport.mapTo(table, QPoint())
        header = table.horizontalHeader()
        for c in range(model.columnCount()):
            if table.isColumnHidden(c):
                continue
            x = origin.x() + table.columnViewportPosition(c) + table.columnWidth(c) - 0.5
            if bounds.left() < x < bounds.right():
                path.moveTo(x, bounds.top())
                path.lineTo(x, bounds.bottom())
        if header.isVisible() and bounds.top() < origin.y() - 0.5 < bounds.bottom():
            path.moveTo(bounds.left(), origin.y() - 0.5)
            path.lineTo(bounds.right(), origin.y() - 0.5)
        for r in range(model.rowCount()):
            if table.isRowHidden(r):
                continue
            y = origin.y() + table.rowViewportPosition(r) + table.rowHeight(r) - 0.5
            if bounds.top() < y < bounds.bottom():
                path.moveTo(bounds.left(), y)
                path.lineTo(bounds.right(), y)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        pen = QPen(self.color, 1)
        pen.setCapStyle(Qt.FlatCap)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        # One stroke prevents alpha accumulating where cell and frame edges meet.
        painter.drawPath(path)
        painter.end()


class QuoteTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.grid_layer = _TableGrid(self)
        for header in (self.horizontalHeader(), self.verticalHeader()):
            header.sectionResized.connect(self.grid_layer.update)
            header.sectionMoved.connect(self.grid_layer.update)
            header.geometriesChanged.connect(self.grid_layer.update)
        for scrollbar in (self.horizontalScrollBar(), self.verticalScrollBar()):
            scrollbar.valueChanged.connect(self.grid_layer.update)

    def setModel(self, model):
        old = self.model()
        if old is not None:
            old.modelReset.disconnect(self.grid_layer.update)
            old.layoutChanged.disconnect(self.grid_layer.update)
        super().setModel(model)
        if model is not None:
            model.modelReset.connect(self.grid_layer.update)
            model.layoutChanged.connect(self.grid_layer.update)
        self.grid_layer.update()

    def set_grid_appearance(self, enabled, color):
        self.setShowGrid(False)
        self.grid_layer.color = QColor(color)
        self.grid_layer.setGeometry(self.rect())
        self.grid_layer.setVisible(enabled)
        self.grid_layer.raise_()
        self.grid_layer.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.grid_layer.setGeometry(self.rect())
        self.grid_layer.raise_()
