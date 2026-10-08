"""History plotting with QtGui only; no extra charting/runtime dependency."""

import math
from datetime import datetime

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QWidget, QToolTip

from stockwidget.core.quote_presentation import format_value, price_precision, should_use_english_units, volume_lot_size
from stockwidget.data.bars import MA_PERIODS, intraday_average, intraday_timeline, moving_average, trading_date


MA_COLORS = ("#e3ac28", "#ba6ee0", "#2a9fd6", "#d97839", "#6fa84c")


class HistoryChart(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 280)
        self.setMouseTracking(True)
        self.bars = ()
        self.averages = {}
        self.minute_averages = []
        self.periods = set(MA_PERIODS)
        self.show_average = self.show_volume = True
        self.view = "intraday"
        self.instrument = {}
        self.unit_mode = "auto"
        self.reference_price = None
        self.up_color, self.down_color = QColor("#ff4444"), QColor("#00b060")
        self._hover = None
        self._plot = QRectF()
        self._volume_plot = QRectF()
        self._reference_y = None
        self._time_axis = ()
        self._positions = ()
        self._indices = {}

    def set_data(self, bars, view, instrument, *, reference_price=None, now=None):
        self.view, self.instrument = view, instrument
        self.bars = tuple(bars[-30:] if view == "daily" else bars)
        self.averages = {period: moving_average(bars, period)[-30:] for period in MA_PERIODS} if view == "daily" else {}
        self.minute_averages = intraday_average(self.bars, instrument) if view != "daily" else []
        self.reference_price = reference_price if reference_price and math.isfinite(reference_price) and reference_price > 0 else None
        timeline = intraday_timeline(self.bars, instrument, now) if view == "intraday" else ()
        self._time_axis = timeline or tuple(bar.time for bar in self.bars)
        slots = {stamp: i for i, stamp in enumerate(timeline)}
        self._positions = tuple(slots[datetime.fromisoformat(bar.time).strftime("%Y-%m-%d %H:%M:00")]
                                for bar in self.bars) if timeline else tuple(range(len(self.bars)))
        self._indices = {slot: i for i, slot in enumerate(self._positions)}
        self._hover = None
        self.update()

    def set_periods(self, periods):
        self.periods = set(periods)
        self.update()

    def set_options(self, periods, average, volume, unit_mode="auto"):
        self.show_average, self.show_volume = average, volume
        self.unit_mode = unit_mode
        self.set_periods(periods)

    def _price(self, value):
        precision = price_precision(self.instrument.get("type"), self.instrument.get("market", ""))
        return f"{value:.{precision}f}"

    def _indicator_width(self):
        return max(.65, min(1.5, 1.5 * min(self.width() / 740, self.height() / 420)))

    def _volume(self, value):
        market, security_type = self.instrument.get("market", ""), self.instrument.get("type")
        lot = volume_lot_size(security_type, market)
        unit = "手" if lot == 100 else "合约" if security_type == "期" else "股"
        return format_value(value, lot, not should_use_english_units(self.unit_mode, market)) + unit

    def _percent(self, price):
        change = (price / self.reference_price - 1) * 100
        return "0%" if abs(change) < .005 else f"{change:+.2f}%"

    def price_range(self):
        values = [price for bar in self.bars for price in ((bar.low, bar.high) if self.view == "daily" else (bar.close,))]
        if self.view == "daily":
            values += [value for period in self.periods for value in self.averages[period] if value is not None]
        else:
            if self.show_average:
                values += self.minute_averages
            if self.reference_price:
                values.append(self.reference_price)
        low, high = min(values), max(values)
        padding = max((high - low) * .08, abs(high) * .002, .001)
        return (max(0, low - padding) if low > 0 else low - padding), high + padding

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        palette = self.palette()
        painter.fillRect(self.rect(), palette.base())
        self._reference_y = None
        if not self.bars:
            painter.setPen(palette.text().color())
            painter.drawText(self.rect(), Qt.AlignCenter, "暂无可用历史数据")
            return
        low, high = self.price_range()
        reference = self.reference_price if self.view != "daily" else None
        max_volume = max(bar.volume for bar in self.bars)
        labels = [self._price(price) for price in (low, high)]
        if self.show_volume:
            labels.append(self._volume(max_volume))
        left = max(48, max(self.fontMetrics().horizontalAdvance(label) for label in labels) + 8)
        right = (max(self.fontMetrics().horizontalAdvance(self._percent(price))
                     for price in (low, high, reference)) + 8 if reference else 12)
        label_height = self.fontMetrics().height() + 4
        volume_gap = label_height + 4
        height = max(1, self.height() - 44)
        volume_height = min(height / 2, max(28, height * .25)) if self.show_volume else 0
        plot = self._plot = QRectF(left, 12, max(1, self.width() - left - right),
                                  max(1, height - volume_height - volume_gap) if self.show_volume else height)
        volume_rect = self._volume_plot = QRectF(plot.left(), plot.bottom() + volume_gap, plot.width(), volume_height)
        step = plot.width() / len(self._time_axis)
        x_at = lambda i: plot.left() + step * (self._positions[i] + .5)
        y_at = lambda price: plot.bottom() - (price - low) / (high - low) * plot.height()
        text_color = palette.text().color()
        grid_color = palette.mid().color()
        grid_color.setAlpha(80)
        reference_y = self._reference_y = y_at(reference) if reference else None

        def axis_label(price, y):
            painter.setPen(text_color)
            painter.drawText(QRectF(0, y - label_height / 2, left - 8, label_height),
                             Qt.AlignRight | Qt.AlignVCenter, self._price(price))
            if reference:
                painter.drawText(QRectF(plot.right() + 4, y - label_height / 2, right - 4, label_height),
                                 Qt.AlignLeft | Qt.AlignVCenter, self._percent(price))

        painter.setPen(QPen(grid_color, 1))
        ticks = min(5, int(plot.height() / (label_height + 4)) + 1)
        for i in range(ticks):
            fraction = i / (ticks - 1) if ticks > 1 else .5
            y = plot.top() + plot.height() * fraction
            if reference_y is not None and abs(y - reference_y) < label_height:
                continue
            painter.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
            axis_label(high - (high - low) * fraction, y)
            painter.setPen(QPen(grid_color, 1))
        dates = [trading_date(bar, self.instrument) for bar in self.bars]
        for i in range(1, len(dates)):
            if dates[i] != dates[i - 1] and self.view != "daily":
                painter.drawLine(QPointF(x_at(i) - step / 2, plot.top()),
                                 QPointF(x_at(i) - step / 2, volume_rect.bottom() if self.show_volume else plot.bottom()))
        if reference_y is not None:
            zero_color = QColor(text_color)
            zero_color.setAlpha(150)
            painter.setPen(QPen(zero_color, 1, Qt.DashLine))
            painter.drawLine(QPointF(plot.left(), reference_y), QPointF(plot.right(), reference_y))
            axis_label(reference, reference_y)
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
                    self._line(painter, self.averages[period], x_at, y_at, QColor(color),
                               width=self._indicator_width())
        else:
            self._line(painter, [bar.close for bar in self.bars], x_at, y_at, palette.highlight().color(), dates)
            if self.show_average:
                self._line(painter, self.minute_averages, x_at, y_at, QColor(MA_COLORS[0]), dates,
                           width=self._indicator_width())
        painter.setPen(text_color)
        if self.show_volume:
            separator = QColor(text_color)
            separator.setAlpha(110)
            painter.setPen(QPen(separator, 1))
            separator_y = (plot.bottom() + volume_rect.top()) / 2
            painter.drawLine(QPointF(0, separator_y), QPointF(self.width(), separator_y))
            painter.drawLine(volume_rect.bottomLeft(), volume_rect.bottomRight())
            scale_volume = max(1, max_volume)
            for i, bar in enumerate(self.bars):
                color = self.up_color if bar.close >= bar.open else self.down_color
                painter.fillRect(QRectF(x_at(i) - step * .3, volume_rect.bottom() - bar.volume / scale_volume * volume_rect.height(),
                                       max(1, step * .6), bar.volume / scale_volume * volume_rect.height()), color)
            painter.setPen(text_color)
            painter.drawText(QRectF(0, volume_rect.top(), left - 8, label_height), Qt.AlignRight, self._volume(max_volume))
            if volume_rect.height() >= label_height * 2 + 4:
                painter.drawText(QRectF(0, volume_rect.bottom() - label_height, left - 8, label_height),
                                 Qt.AlignRight | Qt.AlignBottom, "0")
        painter.setPen(text_color)
        for rect, label in self._time_labels():
            painter.drawText(rect, Qt.AlignCenter, label)
        if self._hover is not None:
            painter.setPen(QPen(palette.highlight().color(), 1, Qt.DashLine))
            painter.drawLine(QPointF(x_at(self._hover), plot.top()),
                             QPointF(x_at(self._hover), volume_rect.bottom() if self.show_volume else plot.bottom()))

    def _time_labels(self):
        """Keep both endpoints and omit intermediate ticks that would overlap."""
        if not self.bars:
            return []
        times = self._time_axis
        step = self._plot.width() / len(times)
        labels = []
        compact_dates = (self.view == "five_day" and
                         sum(self.fontMetrics().horizontalAdvance(stamp[5:16]) + 4
                             for stamp in (times[0], times[-1])) + 8 > self._plot.width())
        for i in sorted({0, len(times) // 4, len(times) // 2, 3 * len(times) // 4, len(times) - 1}):
            label = (times[i][5:10] if self.view == "daily" or compact_dates else
                     times[i][11:16] if self.view == "intraday" else times[i][5:16])
            width = self.fontMetrics().horizontalAdvance(label) + 4
            x = min(self.width() - width - 8, max(8,
                                                self._plot.left() + step * (i + .5) - width / 2))
            labels.append((QRectF(x, self.height() - 28, width, 22), label))
        visible = labels[:1]
        for rect, label in labels[1:-1]:
            if rect.left() >= visible[-1][0].right() + 8 and rect.right() + 8 <= labels[-1][0].left():
                visible.append((rect, label))
        return visible + labels[-1:] if len(labels) > 1 else visible

    @staticmethod
    def _line(painter, values, x_at, y_at, color, dates=None, *, width=1.5):
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
        painter.setPen(QPen(color, width))
        painter.drawPath(path)

    def mouseMoveEvent(self, event):
        slot = (int((event.position().x() - self._plot.left()) / self._plot.width() * len(self._time_axis))
                if self.bars and self._plot.contains(event.position()) else None)
        i = self._indices.get(slot)
        if i is None:
            self._hover = None
            QToolTip.hideText()
            self.update()
            return
        text = self.tooltip_text(i)
        self._hover = i
        QToolTip.showText(event.globalPosition().toPoint(), text, self)
        self.update()

    def tooltip_text(self, i):
        bar = self.bars[i]
        text = (f"{bar.time}\n开 {self._price(bar.open)}  高 {self._price(bar.high)}  "
                f"低 {self._price(bar.low)}  收 {self._price(bar.close)}")
        if self.show_volume:
            text += f"\n成交量 {self._volume(bar.volume)}"
        if self.view != "daily" and self.show_average:
            label = "分时均线" if self.instrument.get("type") == "指" else "均价"
            text += f"\n{label} {self._price(self.minute_averages[i])}"
        for period in sorted(self.periods):
            value = self.averages.get(period, [])[i] if self.view == "daily" else None
            if value is not None:
                text += f"\nMA{period}: {self._price(value)}"
        return text

    def leaveEvent(self, event):
        self._hover = None
        QToolTip.hideText()
        self.update()
        super().leaveEvent(event)
