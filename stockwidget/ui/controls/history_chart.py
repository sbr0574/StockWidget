"""History plotting with QtGui only; no extra charting/runtime dependency."""

import math
from datetime import datetime

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QWidget

from stockwidget.core.quote_presentation import format_value, price_precision, should_use_english_units, volume_lot_size
from stockwidget.data.bars import MA_PERIODS, intraday_average, intraday_timeline, market_now, moving_average, trading_date, volume_directions


MA_COLORS = ("#e3ac28", "#ba6ee0", "#2a9fd6", "#d97839", "#6fa84c")


class HistoryChart(QWidget):
    readout_changed = Signal(str)
    readout_layout_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 280)
        self.setMouseTracking(True)
        self.bars = ()
        self.averages = {}
        self.minute_averages = []
        self.opening_prices = {}
        self.volume_directions = []
        self.periods = set(MA_PERIODS)
        self.show_average = self.show_volume = True
        self.view = "intraday"
        self.instrument = {}
        self.unit_mode = "auto"
        self.reference_price = None
        self.up_color, self.down_color = QColor("#ff4444"), QColor("#00b060")
        self._hover = None
        self._cursor_pos = None
        self._hover_slot = None
        self._plot = QRectF()
        self._volume_plot = QRectF()
        self._reference_y = None
        self._time_axis = ()
        self._positions = ()
        self._indices = {}
        self._day_ranges = ()

    def set_data(self, bars, view, instrument, *, reference_price=None, now=None):
        if self.view != view or any(self.instrument.get(key) != instrument.get(key) for key in ("market", "code", "type")):
            self._cursor_pos = None
        self.view, self.instrument = view, instrument
        self.bars = tuple(bars[-30:] if view == "daily" else bars)
        self.averages = {period: moving_average(bars, period)[-30:] for period in MA_PERIODS} if view == "daily" else {}
        self.minute_averages = intraday_average(self.bars, instrument) if view != "daily" else []
        self.opening_prices = {}
        if view != "daily":
            for i, bar in enumerate(self.bars):
                first = not i or trading_date(bar, instrument) != trading_date(self.bars[i - 1], instrument)
                if first and bar.time[11:16] in {"09:30", "09:31", "09:00", "09:01", "21:00", "21:01"}:
                    self.opening_prices[i] = bar.open
        self.volume_directions = volume_directions(self.bars, view, instrument)
        self.reference_price = reference_price if reference_price and math.isfinite(reference_price) and reference_price > 0 else None
        timeline = intraday_timeline(self.bars, instrument, now) if view == "intraday" else ()
        self._day_ranges = []
        if view == "five_day":
            times = []
            for day in dict.fromkeys(trading_date(bar, instrument) for bar in self.bars):
                group = tuple(bar for bar in self.bars if trading_date(bar, instrument) == day)
                local = market_now(instrument, now)
                clock = datetime.fromisoformat(day).replace(hour=12, tzinfo=local.tzinfo)
                slots = intraday_timeline(group, instrument, clock) or tuple(bar.time for bar in group)
                self._day_ranges.append((len(times), len(times) + len(slots), day))
                times.extend(slots)
            timeline = tuple(times)
        self._time_axis = timeline or tuple(bar.time for bar in self.bars)
        slots = {datetime.fromisoformat(stamp).strftime("%Y-%m-%d %H:%M:00"): i for i, stamp in enumerate(timeline)}
        self._positions = tuple(slots[datetime.fromisoformat(bar.time).strftime("%Y-%m-%d %H:%M:00")]
                                for bar in self.bars) if timeline else tuple(range(len(self.bars)))
        self._indices = {slot: i for i, slot in enumerate(self._positions)}
        self._locate_cursor()
        self._emit_readout()
        self.readout_layout_changed.emit()
        self.update()

    def set_periods(self, periods):
        self.periods = set(periods)
        self._emit_readout()
        self.readout_layout_changed.emit()
        self.update()

    def set_options(self, periods, average, volume, unit_mode="auto"):
        self.show_average, self.show_volume = average, volume
        self.unit_mode = unit_mode
        self.set_periods(periods if average else ())

    def _price(self, value):
        precision = price_precision(self.instrument.get("type"), self.instrument.get("market", ""))
        return f"{value:.{precision}f}"

    def _indicator_width(self):
        return max(.65, min(1.5, 1.5 * min(self.width() / 740, self.height() / 420)))

    def _volume(self, value):
        market, security_type = self.instrument.get("market", ""), self.instrument.get("type")
        lot = volume_lot_size(security_type, market)
        unit = "" if lot == 100 else "合约" if security_type == "期" else "股"
        return format_value(value, lot, not should_use_english_units(self.unit_mode, market)) + unit

    def _percent(self, price):
        change = (price / self.reference_price - 1) * 100
        return "0%" if abs(change) < .005 else f"{change:+.2f}%"

    def price_range(self):
        values = [price for bar in self.bars for price in ((bar.low, bar.high) if self.view == "daily" else (bar.close,))]
        if self.view == "daily":
            values += [value for period in self.periods for value in self.averages[period] if value is not None]
        else:
            values += list(self.opening_prices.values())
            if self.show_average:
                values += self.minute_averages
            if self.reference_price:
                values.append(self.reference_price)
        low, high = min(values), max(values)
        padding = max((high - low) * .08, abs(high) * .002, .001)
        if self.view != "daily" and self.reference_price:
            radius = max(high - self.reference_price, self.reference_price - low) + padding
            return self.reference_price - radius, self.reference_price + radius
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
        height = max(1, self.height() - 44)
        volume_gap = min(label_height + 4, max(5, height * .08))
        volume_height = min(height / 2, max(12, height * .25)) if self.show_volume else 0
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
        for i in range(5):
            fraction = i / 4
            y = plot.top() + plot.height() * fraction
            painter.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
            if plot.height() >= 4 * label_height or i in {0, 2, 4}:
                axis_label(high - (high - low) * fraction, y)
            painter.setPen(QPen(grid_color, 1))
        dates = [trading_date(bar, self.instrument) for bar in self.bars]
        boundaries = ([plot.left() + plot.width() * i / 4 for i in range(5)] if self.view == "intraday" else
                      [plot.left() + step * start for start, stop, day in self._day_ranges] + [plot.right()]
                      if self.view == "five_day" else [])
        for x in boundaries:
            painter.drawLine(QPointF(x, plot.top()), QPointF(x, volume_rect.bottom() if self.show_volume else plot.bottom()))
        if reference_y is not None:
            zero_color = QColor(text_color)
            zero_color.setAlpha(150)
            painter.setPen(QPen(zero_color, 1))
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
            openings = {}
            for i, price in self.opening_prices.items():
                start = next((start for start, stop, day in self._day_ranges if start <= self._positions[i] < stop), 0)
                x = x_at(i) - step * (1.5 if self.bars[i].time[14:16] in {"01", "31"} else .5)
                openings[i] = QPointF(max(plot.left() + start * step, x), y_at(price))
            self._line(painter, [bar.close for bar in self.bars], x_at, y_at, palette.highlight().color(), dates,
                       openings=openings)
            if self.show_average:
                self._line(painter, self.minute_averages, x_at, y_at, QColor(MA_COLORS[0]), dates,
                           width=self._indicator_width())
        painter.setPen(text_color)
        if self.show_volume:
            painter.setRenderHint(QPainter.Antialiasing, False)
            separator = QColor(text_color)
            separator.setAlpha(110)
            painter.setPen(QPen(separator, 1))
            separator_y = (plot.bottom() + volume_rect.top()) / 2
            painter.drawLine(QPointF(0, separator_y), QPointF(self.width(), separator_y))
            painter.drawLine(volume_rect.bottomLeft(), volume_rect.bottomRight())
            scale_volume = max(1, max_volume)
            for i, bar in enumerate(self.bars):
                direction = self.volume_directions[i]
                color = self.up_color if direction > 0 else self.down_color if direction < 0 else text_color
                if bar.volume > 0:
                    width = max(1, round(step * .6))
                    height = max(1, round(bar.volume / scale_volume * volume_rect.height()))
                    painter.fillRect(QRectF(round(x_at(i)) - width // 2, round(volume_rect.bottom()) - height,
                                           width, height), color)
            painter.setPen(text_color)
            painter.drawText(QRectF(0, volume_rect.top(), left - 8, label_height), Qt.AlignRight, self._volume(max_volume))
            if volume_rect.height() >= label_height * 2 + 4:
                painter.drawText(QRectF(0, volume_rect.bottom() - label_height, left - 8, label_height),
                                 Qt.AlignRight | Qt.AlignBottom, "0")
            painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(text_color)
        for rect, label in self._time_labels():
            painter.drawText(rect, Qt.AlignCenter, label)
        self._draw_crosshair(painter, low, high, max_volume, left, right, label_height)

    def _time_labels(self):
        """Keep both endpoints and omit intermediate ticks that would overlap."""
        if not self.bars:
            return []
        times = self._time_axis
        step = self._plot.width() / len(times)
        labels = []
        if self.view == "five_day":
            for start, stop, day in self._day_ranges:
                label = day[5:10]
                width = self.fontMetrics().horizontalAdvance(label) + 4
                center = self._plot.left() + (start + stop) / 2 * step
                labels.append((QRectF(center - width / 2, self.height() - 28, width, 22), label))
            return labels
        for i in sorted({0, len(times) // 4, len(times) // 2, 3 * len(times) // 4, len(times) - 1}):
            label = (times[i][5:10] if self.view == "daily" else
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
    def _line(painter, values, x_at, y_at, color, dates=None, *, width=1.5, openings=None):
        path, started = QPainterPath(), False
        singletons, segment = [], []
        for i, value in enumerate(values):
            if value is None:
                if len(segment) == 1:
                    singletons.extend(segment)
                segment = []
                started = False
                continue
            if dates and i and dates[i] != dates[i - 1]:
                if len(segment) == 1:
                    singletons.extend(segment)
                segment = []
                started = False
            point = QPointF(x_at(i), y_at(value))
            if started:
                path.lineTo(point)
            else:
                opening = openings.get(i) if openings else None
                path.moveTo(opening if opening is not None else point)
                if opening is not None:
                    path.lineTo(point)
                    segment.append(opening)
                started = True
            segment.append(point)
        if len(segment) == 1:
            singletons.extend(segment)
        painter.setPen(QPen(color, width))
        painter.drawPath(path)
        # A lone observed minute still has a short line stroke, not a separate
        # opening-price marker. Opening and first close share the same path.
        if singletons:
            for point in singletons:
                painter.drawLine(point - QPointF(.75, 0), point + QPointF(.75, 0))

    def mouseMoveEvent(self, event):
        pos = event.position()
        self._cursor_pos = QPointF(pos) if self._plot.contains(pos) or self.show_volume and self._volume_plot.contains(pos) else None
        self._locate_cursor()
        self._emit_readout()
        self.update()

    def _locate_cursor(self):
        self._hover_slot = (min(len(self._time_axis) - 1, max(0, int(
            (self._cursor_pos.x() - self._plot.left()) / self._plot.width() * len(self._time_axis))))
            if self._cursor_pos is not None and self._time_axis and self._plot.width() else None)
        self._hover = self._indices.get(self._hover_slot)

    def _emit_readout(self):
        text = self.readout_text()
        self.setAccessibleDescription(text)
        self.readout_changed.emit(text)

    def readout_text(self, i=None):
        if not self.bars:
            return ("开 —  高 —  低 —  收 —  量 —" if self.view == "daily" else "均 —  价格 — (—)  量 —")
        if i is None:
            if self._cursor_pos is not None and self._hover is None and self._hover_slot is not None:
                return "均 —  价格 — (—)  量 —"
            i = self._hover if self._hover is not None else len(self.bars) - 1
        bar = self.bars[i]
        if self.view != "daily":
            average = self.minute_averages[i]
            change = f"{(bar.close / self.reference_price - 1) * 100:+.2f}%" if self.reference_price else "—"
            return f"均 {self._price(average)}  价格 {self._price(bar.close)} ({change})  量 {self._volume(bar.volume)}"
        text = (f"开 {self._price(bar.open)}  高 {self._price(bar.high)}  "
                f"低 {self._price(bar.low)}  收 {self._price(bar.close)}  量 {self._volume(bar.volume)}")
        averages = []
        for period in sorted(self.periods):
            value = self.averages[period][i]
            averages.append(f"MA{period}: {self._price(value) if value is not None else '—'}")
        if averages:
            text += "\n" + "  ".join(averages)
        return text

    def readout_size_sample(self):
        """Reserve footer space independently of which minute is hovered."""
        if not self.bars:
            return self.readout_text()
        prices = [bar.high for bar in self.bars]
        prices.extend(value for values in self.averages.values() for value in values if value is not None)
        price = self._price(max(prices))
        volume = max((self._volume(bar.volume) for bar in self.bars), key=len)
        if self.view != "daily":
            change = max((f"{(bar.close / self.reference_price - 1) * 100:+.2f}%" for bar in self.bars), key=len) if self.reference_price else "—"
            return f"均 {price}  价格 {price} ({change})  量 {volume}"
        text = f"开 {price}  高 {price}  低 {price}  收 {price}  量 {volume}"
        if self.periods:
            text += "\n" + "  ".join(f"MA{period}: {price}" for period in sorted(self.periods))
        return text

    def _draw_crosshair(self, painter, low, high, max_volume, left, right, label_height):
        if self._cursor_pos is None or self._hover_slot is None:
            return
        plot, volume = self._plot, self._volume_plot
        point = self._cursor_pos
        if not (plot.contains(point) or self.show_volume and volume.contains(point)):
            return
        color = self.palette().highlight().color()
        painter.setPen(QPen(color, 1, Qt.DashLine))
        painter.drawLine(QPointF(point.x(), plot.top()), QPointF(point.x(), volume.bottom() if self.show_volume else plot.bottom()))
        painter.drawLine(QPointF(plot.left(), point.y()), QPointF(plot.right(), point.y()))

        def badge(rect, text):
            painter.fillRect(rect, color)
            painter.setPen(self.palette().highlightedText().color())
            painter.drawText(rect, Qt.AlignCenter, text)

        if plot.contains(point):
            price = low + (plot.bottom() - point.y()) / plot.height() * (high - low)
            badge(QRectF(0, point.y() - label_height / 2, left - 3, label_height), self._price(price))
            if self.reference_price and self.view != "daily":
                badge(QRectF(plot.right() + 2, point.y() - label_height / 2, right - 2, label_height), self._percent(price))
        else:
            value = (volume.bottom() - point.y()) / volume.height() * max(1, max_volume)
            badge(QRectF(0, point.y() - label_height / 2, left - 3, label_height), self._volume(value))
        stamp = self._time_axis[self._hover_slot]
        label = stamp[11:16] if self.view == "intraday" else stamp[5:16] if self.view == "five_day" else stamp[5:10]
        width = self.fontMetrics().horizontalAdvance(label) + 8
        x = max(0, min(self.width() - width, point.x() - width / 2))
        badge(QRectF(x, self.height() - 28, width, 22), label)

    def leaveEvent(self, event):
        self._hover = None
        self._hover_slot = None
        self._cursor_pos = None
        self._emit_readout()
        self.update()
        super().leaveEvent(event)
