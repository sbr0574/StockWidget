"""History plotting with QtGui only; no extra charting/runtime dependency."""

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QWidget, QToolTip

from stockwidget.data.bars import MA_PERIODS, moving_average, trading_date


MA_COLORS = ("#e3ac28", "#ba6ee0", "#2a9fd6", "#d97839", "#6fa84c")


class HistoryChart(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 280)
        self.setMouseTracking(True)
        self.bars = ()
        self.averages = {}
        self.periods = set(MA_PERIODS)
        self.view = "intraday"
        self.instrument = {}
        self.up_color, self.down_color = QColor("#ff4444"), QColor("#00b060")
        self._hover = None
        self._plot = QRectF()

    def set_data(self, bars, view, instrument):
        self.view, self.instrument = view, instrument
        self.bars = tuple(bars[-30:] if view == "daily" else bars)
        self.averages = {period: moving_average(bars, period)[-30:] for period in MA_PERIODS} if view == "daily" else {}
        self._hover = None
        self.update()

    def set_periods(self, periods):
        self.periods = set(periods)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        palette = self.palette()
        painter.fillRect(self.rect(), palette.base())
        if not self.bars:
            painter.setPen(palette.text().color())
            painter.drawText(self.rect(), Qt.AlignCenter, "暂无可用历史数据")
            return
        plot = self._plot = QRectF(68, 20, max(1, self.width() - 88), max(1, self.height() * .68 - 30))
        volume_rect = QRectF(plot.left(), plot.bottom() + 28, plot.width(), max(1, self.height() - plot.bottom() - 64))
        values = [price for bar in self.bars for price in ((bar.low, bar.high) if self.view == "daily" else (bar.close,))]
        if self.view == "daily":
            values += [value for period in self.periods for value in self.averages[period] if value is not None]
        else:
            values += [bar.average for bar in self.bars if bar.average is not None and bar.average > 0]
        low, high = min(values), max(values)
        padding = max((high - low) * .08, abs(high) * .002, .001)
        low, high = low - padding, high + padding
        step = plot.width() / len(self.bars)
        x_at = lambda i: plot.left() + step * (i + .5)
        y_at = lambda price: plot.bottom() - (price - low) / (high - low) * plot.height()
        text_color = palette.text().color()
        grid_color = palette.mid().color()
        grid_color.setAlpha(80)
        painter.setPen(QPen(grid_color, 1))
        for i in range(5):
            y = plot.top() + plot.height() * i / 4
            painter.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
            painter.setPen(text_color)
            painter.drawText(QRectF(0, y - 10, 60, 20), Qt.AlignRight | Qt.AlignVCenter,
                             f"{high - (high - low) * i / 4:.3f}".rstrip("0").rstrip("."))
            painter.setPen(QPen(grid_color, 1))
        dates = [trading_date(bar, self.instrument) for bar in self.bars]
        for i in range(1, len(dates)):
            if dates[i] != dates[i - 1] and self.view != "daily":
                painter.drawLine(QPointF(x_at(i) - step / 2, plot.top()),
                                 QPointF(x_at(i) - step / 2, volume_rect.bottom()))
        if self.view == "daily":
            for i, bar in enumerate(self.bars):
                color = self.up_color if bar.close >= bar.open else self.down_color
                painter.setPen(QPen(color, 1))
                x = x_at(i)
                painter.drawLine(QPointF(x, y_at(bar.high)), QPointF(x, y_at(bar.low)))
                top, bottom = sorted((y_at(bar.open), y_at(bar.close)))
                painter.fillRect(QRectF(x - step * .3, top, step * .6, max(1, bottom - top)), color)
            for period, color in zip(MA_PERIODS, MA_COLORS):
                if period in self.periods:
                    self._line(painter, self.averages[period], x_at, y_at, QColor(color))
        else:
            self._line(painter, [bar.close for bar in self.bars], x_at, y_at, palette.highlight().color(), dates)
            self._line(painter, [bar.average if bar.average and bar.average > 0 else None for bar in self.bars],
                       x_at, y_at, QColor(MA_COLORS[0]), dates)
        max_volume = max(1, max(bar.volume for bar in self.bars))
        for i, bar in enumerate(self.bars):
            color = self.up_color if bar.close >= bar.open else self.down_color
            painter.fillRect(QRectF(x_at(i) - step * .3, volume_rect.bottom() - bar.volume / max_volume * volume_rect.height(),
                                   max(1, step * .6), bar.volume / max_volume * volume_rect.height()), color)
        painter.setPen(text_color)
        painter.drawText(QRectF(0, volume_rect.top(), 60, 20), Qt.AlignRight, "成交量")
        for rect, label in self._time_labels():
            painter.drawText(rect, Qt.AlignCenter, label)
        if self._hover is not None:
            painter.setPen(QPen(palette.highlight().color(), 1, Qt.DashLine))
            painter.drawLine(QPointF(x_at(self._hover), plot.top()), QPointF(x_at(self._hover), volume_rect.bottom()))

    def _time_labels(self):
        """Keep both endpoints and omit intermediate ticks that would overlap."""
        if not self.bars:
            return []
        step = self._plot.width() / len(self.bars)
        labels = []
        for i in sorted({0, len(self.bars) // 4, len(self.bars) // 2, 3 * len(self.bars) // 4, len(self.bars) - 1}):
            label = self.bars[i].time[5:10] if self.view == "daily" else self.bars[i].time[5:16]
            width = self.fontMetrics().horizontalAdvance(label) + 4
            x = min(self.width() - width - 8, max(self._plot.left(),
                                                self._plot.left() + step * (i + .5) - width / 2))
            labels.append((QRectF(x, self.height() - 28, width, 22), label))
        visible = labels[:1]
        for rect, label in labels[1:-1]:
            if rect.left() >= visible[-1][0].right() + 8 and rect.right() + 8 <= labels[-1][0].left():
                visible.append((rect, label))
        return visible + labels[-1:] if len(labels) > 1 else visible

    @staticmethod
    def _line(painter, values, x_at, y_at, color, dates=None):
        path, started = QPainterPath(), False
        for i, value in enumerate(values):
            if value is None:
                started = False
                continue
            if dates and i and dates[i] != dates[i - 1]:
                started = False
            point = QPointF(x_at(i), y_at(value))
            if started:
                path.lineTo(point)
            else:
                path.moveTo(point)
                started = True
        painter.setPen(QPen(color, 1.5))
        painter.drawPath(path)

    def mouseMoveEvent(self, event):
        if not self.bars or not self._plot.contains(event.position()):
            QToolTip.hideText()
            return
        i = min(len(self.bars) - 1, max(0, int((event.position().x() - self._plot.left()) / self._plot.width() * len(self.bars))))
        bar = self.bars[i]
        text = f"{bar.time}\n开 {bar.open:g}  高 {bar.high:g}  低 {bar.low:g}  收 {bar.close:g}\n成交量 {bar.volume:g}"
        for period in sorted(self.periods):
            value = self.averages.get(period, [])[i] if self.view == "daily" else None
            if value is not None:
                text += f"\nMA{period}: {value:.3f}"
        self._hover = i
        QToolTip.showText(event.globalPosition().toPoint(), text, self)
        self.update()

    def leaveEvent(self, event):
        self._hover = None
        QToolTip.hideText()
        self.update()
        super().leaveEvent(event)
