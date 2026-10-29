"""Small quote charts shared by the floating table, taskbar and settings previews."""

import math

from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QFontMetrics, QFontMetricsF, QLinearGradient, QPainter, QPen

from stockwidget.core.quote_presentation import MiniChartCell
from stockwidget.ui.controls.quote_view import KLineDelegate


AXIS_COLOR = QColor("#888888")
CHART_PADDING = 4


def mini_chart_width(font, kind="kline"):
    return 30 if kind == "kline" else QFontMetrics(font).horizontalAdvance("0" * 10)


def mini_chart_stroke(font, row_height):
    """Scale the stroke by the painted font; a tight row shrinks line and dot together."""
    return min(max(0.8, QFontMetricsF(font).height() / 9), max(0.5, (row_height - 4) / 3))


def _plot_rect(rect, font=None):
    padding = max(CHART_PADDING, QFontMetricsF(font).height() / 6 + 1) if font else CHART_PADDING
    return QRectF(rect).adjusted(padding, 2, -padding, -2)


def _valid_limits(limits):
    if limits is None or len(limits) != 2 or not all(math.isfinite(v) for v in limits):
        return None
    return tuple(sorted(limits))


def amplitude_x(rect, value, limits, font=None):
    """One horizontal mapping, including its inset, for ranges and the shared axis."""
    plot = _plot_rect(rect, font)
    low, high = limits
    return plot.center().x() if low == high else plot.left() + (value - low) / (high - low) * plot.width()


def paint_amplitude_axis(painter, rect, limits):
    """Draw one dashed zero axis over a complete visible chart column."""
    limits = _valid_limits(limits)
    if limits is None or not limits[0] <= 0 <= limits[1]:
        return
    rect = QRectF(rect)
    x = amplitude_x(rect, 0, limits, painter.font())
    painter.save()
    painter.setClipRect(rect, Qt.IntersectClip)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setPen(QPen(AXIS_COLOR, 1, Qt.DashLine))
    painter.drawLine(QPointF(x, rect.top()), QPointF(x, rect.bottom()))
    painter.restore()


class MiniChartDelegate(KLineDelegate):
    """Keep the candle renderer while supplying intraday and comparable range views."""

    def sizeHint(self, option, index):
        cell = index.data(Qt.UserRole)
        if isinstance(cell, MiniChartCell):
            return QSize(mini_chart_width(option.font, cell.kind), QFontMetrics(option.font).height())
        return super().sizeHint(option, index)

    def paint(self, painter, option, index):
        cell = index.data(Qt.UserRole)
        if isinstance(cell, MiniChartCell):
            painter.save()
            painter.setFont(option.font)
            self.paint_cell(painter, option.rect, cell,
                            draw_axis=not getattr(option.widget, "_shared_amplitude_axis", False))
            painter.restore()
        else:
            super().paint(painter, option, index)

    def paint_cell(self, painter, rect, cell, *, draw_axis=True):
        if not isinstance(cell, MiniChartCell):
            return
        if cell.kind == "kline":
            if len(cell.candle) == 5 and all(math.isfinite(v) for v in cell.candle):
                self.paint_candle(painter, rect, cell.candle)
        elif cell.kind == "trend":
            self._paint_trend(painter, rect, cell)
        elif cell.kind == "amplitude":
            self._paint_amplitude(painter, rect, cell, draw_axis=draw_axis)

    def amplitude_color(self, percent):
        if self.unicolor:
            return QColor(self.text_color)
        target = self.up_color if percent >= 0 else self.down_color
        strength = -math.expm1(-abs(percent) / 3)
        channels = [round(a + (b - a) * strength)
                    for a, b in zip(self.neutral_color.getRgb(), target.getRgb())]
        return QColor(*channels)

    def _paint_amplitude(self, painter, rect, cell, *, draw_axis):
        values = cell.amplitude
        limits = _valid_limits(cell.limits)
        if values is None:
            return
        low, current, high = values
        limits = limits or (low, high)
        plot = _plot_rect(rect, painter.font())
        if plot.width() <= 0 or plot.height() <= 0:
            return
        if draw_axis:
            paint_amplitude_axis(painter, rect, limits)
        x_low, x_current, x_high = [amplitude_x(rect, value, limits, painter.font()) for value in values]
        stroke = mini_chart_stroke(painter.font(), rect.height())
        y = plot.center().y()
        painter.save()
        painter.setClipRect(QRectF(rect))
        painter.setRenderHint(QPainter.Antialiasing)
        if self.unicolor or low == high:
            pen = QPen(self.amplitude_color(current), stroke)
        else:
            gradient = QLinearGradient(x_low, y, x_high, y)
            # Pixel-spaced samples approximate the nonlinear color curve smoothly.
            steps = max(2, math.ceil(abs(x_high - x_low)))
            stops = {i / steps for i in range(steps + 1)}
            if low < 0 < high:
                stops.add(-low / (high - low))
            for fraction in sorted(stops):
                gradient.setColorAt(fraction, self.amplitude_color(low + (high - low) * fraction))
            pen = QPen(gradient, stroke)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.drawLine(QPointF(x_low, y), QPointF(x_high, y))
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.amplitude_color(current))
        radius = stroke * 1.5
        painter.drawEllipse(QPointF(x_current, y), radius, radius)
        painter.restore()

    def _trend_color(self, price, reference):
        if self.unicolor:
            return QColor(self.text_color)
        if reference is None:
            return QColor(self.neutral_color)
        return QColor(self.up_color if price > reference else
                      self.down_color if price < reference else self.neutral_color)

    def _paint_trend(self, painter, rect, cell):
        points = [(x, price) for x, price in cell.points
                  if math.isfinite(x) and math.isfinite(price) and price > 0]
        if not points:
            return
        reference = cell.reference
        if not math.isfinite(reference) or reference <= 0:
            reference = None
        prices = [price for _, price in points]
        if reference is not None:
            prices.append(reference)
        low, high = min(prices), max(prices)
        stroke = mini_chart_stroke(painter.font(), rect.height()) * 0.75
        inset = max(2, stroke / 2)
        plot = QRectF(rect).adjusted(max(CHART_PADDING, inset), inset,
                                    -max(CHART_PADDING, inset), -inset)
        if plot.width() <= 0 or plot.height() <= 0:
            return

        def position(x, price):
            y = 0.5 if high == low else (high - price) / (high - low)
            return QPointF(plot.left() + max(0, min(1, x)) * plot.width(),
                           plot.top() + y * plot.height())

        painter.save()
        painter.setClipRect(QRectF(rect))
        painter.setRenderHint(QPainter.Antialiasing)
        if reference is not None:
            axis_y = position(0, reference).y()
            painter.setPen(QPen(AXIS_COLOR, 1, Qt.DashLine))
            painter.drawLine(QPointF(plot.left(), axis_y), QPointF(plot.right(), axis_y))
        pen = QPen()
        pen.setWidthF(stroke)
        pen.setCapStyle(Qt.RoundCap)
        current = cell.candle[1] if len(cell.candle) == 5 else points[-1][1]
        pen.setColor(self._trend_color(current, reference))
        painter.setPen(pen)
        if len(points) > 1:
            for (x0, p0), (x1, p1) in zip(points, points[1:]):
                painter.drawLine(position(x0, p0), position(x1, p1))
        painter.restore()
