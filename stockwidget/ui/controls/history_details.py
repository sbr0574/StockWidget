"""Live quote details shared by the chart popup and detached window."""

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPalette
from PySide6.QtWidgets import QLabel, QWidget

from stockwidget.core.quote_presentation import (
    COLOR_ROLE_DOWN, COLOR_ROLE_UP, QuoteDisplayOptions, direction_color_role,
    format_quote, format_value, price_precision, should_use_english_units, volume_lot_size,
)
from stockwidget.ui.controls.quote_view import DEFAULT_DOWN_COLOR, DEFAULT_UP_COLOR


def _set_color(label, role, up, down):
    palette = QPalette(label.parentWidget().palette())
    color = up if role == COLOR_ROLE_UP else down if role == COLOR_ROLE_DOWN else palette.windowText().color()
    palette.setColor(QPalette.WindowText, QColor(color))
    label.setPalette(palette)


class HistoryQuoteHeader(QWidget):
    """Consume the presenter's accepted quote; never issue a quote request."""

    def bind_fields(self):
        self.name = self.findChild(QLabel, "quote_name")
        self.values = {key: self.findChild(QLabel, f"quote_{field}") for key, field in (
            ("现价", "current"), ("涨跌", "change"), ("涨幅", "change_pct"),
            ("最高价", "high"), ("最低价", "low"), ("开盘价", "open"),
            ("成交额", "amount"), ("成交量", "volume"), ("均价", "average"), ("量比", "volume_ratio"),
        )}

    def set_quote(self, instrument, quote, unit_mode, up, down, volume_ratio=None):
        self.name.setText(f"{instrument.get('name') or (quote or {}).get('name') or instrument.get('code', '')}  {instrument.get('code', '')}")
        if not quote:
            for key, label in self.values.items():
                label.setText("—")
                _set_color(label, "", up, down)
            return
        market, kind = instrument.get("market", ""), instrument.get("type")
        precision = price_precision(kind, market)
        formatted, roles, values = format_quote(quote, kind, instrument.get("code", ""), market=market,
                                                options=QuoteDisplayOptions(unit_mode=unit_mode))
        for key in ("现价", "涨跌", "涨幅", "成交额", "成交量", "均价"):
            metric = "均值" if key == "均价" else key
            text = formatted[metric]
            if key == "现价":
                current = values["现价"]
                text = f"{current:.{precision}f}" if current else "—"
            elif key in {"涨跌", "涨幅"} and not quote.get("prev_close"):
                text = "—"
            elif key == "均价" and (not quote.get("deals_amt") or not quote.get("deals_vol") or kind == "期"):
                text = "—"
            self.values[key].setText(text)
            _set_color(self.values[key], roles[metric] if text not in {"—", "-"} else "", up, down)
        for key, field in (("最高价", "high_price"), ("最低价", "low_price"), ("开盘价", "opening_price")):
            price = quote.get(field, 0)
            self.values[key].setText(f"{price:.{precision}f}" if price else "—")
            role = direction_color_role(price - quote.get("prev_close", 0)) if price and quote.get("prev_close") else ""
            _set_color(self.values[key], role, up, down)
        self.values["量比"].setText(f"{volume_ratio:.2f}" if volume_ratio is not None else "—")


class OrderBookWidget(QWidget):
    """Five asks, aggregate buy/sell volumes, then five bids."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(164)
        self.setMinimumHeight(230)
        self.rows = ()
        self.buy_fraction = None
        self.up_color, self.down_color = QColor(DEFAULT_UP_COLOR), QColor(DEFAULT_DOWN_COLOR)
        self.setAccessibleName("五档买卖行情")

    def set_quote(self, instrument, quote, unit_mode, up, down):
        quote = quote or {}
        market, kind = instrument.get("market", ""), instrument.get("type")
        lot = volume_lot_size(kind, market)
        precision = price_precision(kind, market)
        previous = quote.get("prev_close", 0)
        self.up_color, self.down_color = QColor(up), QColor(down)
        self.rows = []
        totals = []
        for side, indices, title in (("seller", range(4, -1, -1), "卖"), ("purchaser", range(5), "买")):
            prices = (list(quote.get(f"{side}_price", ())) + [0] * 5)[:5]
            volumes = (list(quote.get(f"{side}_vol", ())) + [0] * 5)[:5]
            if kind == "指":
                prices, volumes = [0] * 5, [0] * 5
            totals.append(sum(max(0, volume) for price, volume in zip(prices, volumes) if price > 0))
            for index in indices:
                price, volume = prices[index], volumes[index]
                role = direction_color_role(price - previous) if price > 0 and previous else ""
                quantity = (format_value(volume, lot, not should_use_english_units(unit_mode, market))
                            if price > 0 and volume > 0 else "—")
                self.rows.append((f"{title}{index + 1}", f"{price:.{precision}f}" if price > 0 else "—", quantity, role))
        sell, buy = totals
        self.buy_fraction = buy / (buy + sell) if buy + sell else None
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), self.palette().base())
        neutral = self.palette().text().color()
        painter.setPen(self.palette().mid().color())
        painter.drawLine(0, 0, 0, self.height())
        margin, width = 10, self.width() - 20
        line = (self.height() - 12) / 11
        for index, (title, price, quantity, role) in enumerate(self.rows):
            y = 6 + (index + (index >= 5)) * line
            painter.setPen(neutral)
            painter.drawText(QRectF(margin, y, 28, line), Qt.AlignVCenter, title)
            color = self.up_color if role == COLOR_ROLE_UP else self.down_color if role == COLOR_ROLE_DOWN else neutral
            painter.setPen(color)
            painter.drawText(QRectF(margin + 26, y, 57, line), Qt.AlignVCenter | Qt.AlignRight, price)
            painter.setPen(neutral)
            painter.drawText(QRectF(margin + 86, y, width - 86, line), Qt.AlignVCenter | Qt.AlignRight, quantity)
        track = QRectF(margin, 6 + line * 5 + (line - 4) / 2, width, 4)
        painter.fillRect(track, self.palette().mid())
        if self.buy_fraction is not None:
            split = track.width() * self.buy_fraction
            painter.fillRect(QRectF(track.left(), track.top(), split, track.height()), DEFAULT_UP_COLOR)
            painter.fillRect(QRectF(track.left() + split, track.top(), track.width() - split, track.height()), DEFAULT_DOWN_COLOR)
