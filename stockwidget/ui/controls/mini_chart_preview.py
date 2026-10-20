"""固定示例行情，让设置页直接比较三种迷你图的表现。"""

from PySide6.QtCore import QRect, QSize, Qt
from PySide6.QtGui import QFontInfo, QFontMetrics, QPainter, QPalette, QPen
from PySide6.QtWidgets import QPushButton, QStyle, QStyleOptionButton, QStyleOptionViewItem, QStyledItemDelegate

from stockwidget.core.quote_presentation import MiniChartCell
from stockwidget.ui.controls.mini_chart import (
    MiniChartDelegate,
    mini_chart_width,
    paint_amplitude_axis,
)
from stockwidget.ui.controls.quote_view import SimpleTableModel


# 同一组固定行情贯穿所有类型；幅值图共用百分比范围，便于横向比较。
_PREVIEW_QUOTES = (
    ("股A", (100.8, 104.2, 107.0, 98.0, 100.0),
     (100.8, 99.5, 98.0, 100.2, 102.6, 101.8, 104.8, 107.0, 105.4, 104.2)),
    ("股B", (50.3, 48.7, 51.5, 47.5, 50.0),
     (50.3, 51.5, 50.8, 49.7, 48.8, 49.4, 48.0, 47.5, 48.2, 48.7)),
    ("股C", (199.0, 200.8, 202.4, 198.4, 200.0),
     (199.0, 198.4, 199.6, 201.0, 200.1, 202.4, 201.4, 200.4, 201.2, 200.8)),
)
_LIMITS = (-5.0, 7.0)


class MiniChartOption(QPushButton):
    """整块可点击的固定示例按钮，沿用行情模型、绘制代理和共用按钮主题。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.kind = "kline"
        self.model = SimpleTableModel(parent=self)
        self.delegate = MiniChartDelegate(self)
        self._text_delegate = QStyledItemDelegate(self)
        self._set_rows()

    def sizeHint(self):
        return QSize(150, 112)

    def set_kind(self, kind):
        if self.kind != kind:
            self.kind = kind
            self._set_rows()

    def _set_rows(self):
        rows, roles = [], []
        for name, candle, prices in _PREVIEW_QUOTES:
            change = (candle[1] / candle[4] - 1) * 100
            points = tuple((index / (len(prices) - 1), price) for index, price in enumerate(prices))
            chart = MiniChartCell(kind=self.kind, candle=candle, points=points, limits=_LIMITS)
            rows.append([name, chart, f"{change:+.2f}%"])
            role = "up" if change > 0 else "down" if change < 0 else "neutral"
            roles.append(["text", "text", role])
        self.model.set_rows_headers(rows, ["名称", "迷你图", "涨幅"], roles)
        self.model.set_align_right_cols((2,))

    def sync(self, source):
        colors = (source.unicolor, self.palette().color(QPalette.Text),
                  source.up_color, source.down_color, source.neutral_color)
        self.model.set_colors(*colors)
        self.delegate.set_colors(*colors)
        self.delegate.set_point_size(QFontInfo(self.font()).pointSizeF())
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        button = QStyleOptionButton()
        self.initStyleOption(button)
        button.text = ""
        self.style().drawControl(QStyle.CE_PushButton, button, painter, self)
        fm = QFontMetrics(self.font())
        left = 6
        name_width = fm.horizontalAdvance("名称") + 4
        percent_width = fm.horizontalAdvance("+4.20%") + 10
        chart_width = max(1, self.width() - 2 * left - name_width - percent_width)
        plot_width = min(chart_width, mini_chart_width(self.font(), self.kind)) if self.kind == "kline" else chart_width
        widths = (name_width, chart_width, percent_width)
        row_height = max(1, (self.height() - 36) // 4)
        top = 30
        chart_left = left + name_width
        painter.setFont(self.font())
        if self.kind == "amplitude":
            paint_amplitude_axis(painter, QRect(chart_left, top + row_height, chart_width, row_height * 3), _LIMITS)
        x = left
        painter.setPen(self.palette().color(QPalette.Text))
        title_font = self.font()
        title_font.setBold(True)
        painter.setFont(title_font)
        painter.drawText(QRect(left, 6, self.width() - left * 2, 20), Qt.AlignCenter, self.text())
        painter.setFont(self.font())
        for column, width in enumerate(widths):
            alignment = Qt.AlignRight if column >= 2 else Qt.AlignCenter if column == 1 else Qt.AlignLeft
            painter.drawText(QRect(x + 2, top, width - 4, row_height), alignment | Qt.AlignVCenter,
                             self.model.headerData(column, Qt.Horizontal))
            x += width
        painter.setPen(QPen(self.palette().color(QPalette.Mid), 1))
        painter.drawLine(left, top + row_height - 1, self.width() - left, top + row_height - 1)
        for row in range(3):
            x = left
            for column, width in enumerate(widths):
                option = QStyleOptionViewItem()
                option.initFrom(self)
                option.font = self.font()
                option.palette = self.palette()
                option.rect = QRect(x, top + (row + 1) * row_height, width, row_height)
                index = self.model.index(row, column)
                if column == 1:
                    plot = option.rect.adjusted((chart_width - plot_width) // 2, 0,
                                                -(chart_width - plot_width) // 2, 0)
                    self.delegate.paint_cell(painter, plot, index.data(Qt.UserRole), draw_axis=False)
                else:
                    self._text_delegate.paint(painter, option, index)
                x += width
