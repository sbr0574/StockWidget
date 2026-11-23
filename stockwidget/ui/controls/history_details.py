"""Live quote details shared by the chart popup and detached window."""

from PySide6.QtCore import QEvent, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFontMetrics, QPainter, QPalette
from PySide6.QtWidgets import QLabel, QWidget

from stockwidget.core.quote_presentation import (
    COLOR_ROLE_DOWN, COLOR_ROLE_UP, MARKET_BREADTH_BUCKETS, QuoteDisplayOptions, direction_color_role,
    format_quote, format_value, is_cn_index, price_precision, quote_detail_values,
    should_use_english_units, volume_lot_size,
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
        self.code = self.findChild(QLabel, "quote_code")
        self.values = {key: self.findChild(QLabel, f"quote_{field}") for key, field in (
            ("现价", "current"), ("涨跌", "change"), ("涨幅", "change_pct"),
            ("最高价", "high"), ("最低价", "low"), ("开盘价", "open"),
            ("成交额", "amount"), ("成交量", "volume"), ("均价", "average"), ("量比", "volume_ratio"),
            ("昨收", "prev_close"), ("内盘", "inside"), ("外盘", "outside"),
            ("换手率", "turnover"), ("振幅", "amplitude"), ("市盈（TTM）", "pe_ttm"),
            ("市盈（静）", "pe_static"), ("市盈（动）", "pe_dynamic"), ("市净率", "pb"),
            ("总市值", "market_cap"), ("总股本", "total_shares"), ("流通市值", "float_market_cap"), ("流通股", "float_shares"),
        )}

    def set_quote(self, instrument, quote, unit_mode, up, down, volume_ratio=None, details=None):
        self._name_text = instrument.get('name') or (quote or {}).get('name') or instrument.get('code', '')
        self.name.setText(self.name.fontMetrics().elidedText(self._name_text, Qt.ElideRight, self.name.width()))
        self.code.setText(f"{instrument.get('market', '').upper()}{instrument.get('code', '').upper()}")
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
        for key, field in (("最高价", "high_price"), ("最低价", "low_price"), ("开盘价", "opening_price"), ("昨收", "prev_close")):
            price = quote.get(field, 0)
            self.values[key].setText(f"{price:.{precision}f}" if price else "—")
            role = direction_color_role(price - quote.get("prev_close", 0)) if price and quote.get("prev_close") else ""
            _set_color(self.values[key], role, up, down)
        self.values["量比"].setText(f"{volume_ratio:.2f}" if volume_ratio is not None else "—")
        _set_color(self.values["量比"], direction_color_role(volume_ratio - 1) if volume_ratio is not None else "", up, down)
        numbers = quote_detail_values(quote, instrument, details)
        english = should_use_english_units(unit_mode, market)
        for caption, field in (("内盘", "inside_volume"), ("外盘", "outside_volume"),
                               ("换手率", "turnover_rate"), ("振幅", "amplitude"),
                               ("市盈（TTM）", "pe_ttm"), ("市盈（静）", "pe_static"), ("市盈（动）", "pe_dynamic"), ("市净率", "pb"),
                               ("总市值", "market_cap"), ("总股本", "total_shares"),
                               ("流通市值", "float_market_cap"), ("流通股", "float_shares")):
            value = numbers[field]
            if value is None:
                text = "—"
            elif field in {"inside_volume", "outside_volume", "market_cap", "float_market_cap", "total_shares", "float_shares"}:
                lot = volume_lot_size(kind, market) if field in {"inside_volume", "outside_volume"} else 1
                text = format_value(value, lot, not english)
            else:
                text = f"{value:.2f}" + ("%" if field in {"turnover_rate", "amplitude"} else "")
            self.values[caption].setText(text)
            role = (COLOR_ROLE_DOWN if field == "inside_volume" else COLOR_ROLE_UP if field == "outside_volume" else "") if value is not None else ""
            _set_color(self.values[caption], role, up, down)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "_name_text"):
            self.name.setText(self.name.fontMetrics().elidedText(self._name_text, Qt.ElideRight, self.name.width()))


class OrderBookWidget(QWidget):
    """Five asks, aggregate buy/sell volumes, then five bids — or 11-tier CN market breadth."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(164)
        self.setMinimumHeight(200)
        self.flat = False
        self.neutral_color = None
        self.breadth_mode = False
        self.rows = ()
        self.buy_fraction = None
        self.breadth_rows = ()
        self.breadth_summary = None
        self.up_color, self.down_color = QColor(DEFAULT_UP_COLOR), QColor(DEFAULT_DOWN_COLOR)
        self.setAccessibleName("五档买卖行情")

    def set_breadth_mode(self, enabled):
        breadth = bool(enabled)
        changed = self.breadth_mode != breadth
        self.breadth_mode = breadth
        self.setAccessibleName("全市场涨跌分布" if breadth else "五档买卖行情")
        if changed:
            self._sync_size()
            self.update()

    def _sync_size(self):
        if not self.breadth_mode:
            self.setFixedWidth(164)
            self.setMinimumHeight(200)
            return
        metrics = self.fontMetrics()
        bold_font = self.font()
        bold_font.setBold(True)
        bold_metrics = QFontMetrics(bold_font)
        label_width = max(metrics.horizontalAdvance(label) for label, _ in MARKET_BREADTH_BUCKETS)
        count_width = metrics.horizontalAdvance("9999")
        summary_width = bold_metrics.horizontalAdvance("9999 / 9999 / 9999")
        width = max(164, 20 + label_width + count_width + 12 + 48, 20 + summary_width)
        row_height = max(16, metrics.height() + 1)
        height = max(200, 12 + 4 + row_height * 12)
        self.setFixedWidth(width)
        self.setMinimumHeight(height)

    def changeEvent(self, event):
        super().changeEvent(event)
        if self.breadth_mode and event.type() in {QEvent.FontChange, QEvent.DevicePixelRatioChange}:
            self._sync_size()

    def set_quote(self, instrument, quote, unit_mode, up, down, *, market_breadth=None):
        quote = quote or {}
        self.up_color, self.down_color = QColor(up), QColor(down)
        self.set_breadth_mode(is_cn_index(instrument))
        if self.breadth_mode:
            self.rows = ()
            self.buy_fraction = None
            buckets = market_breadth.get("buckets") if isinstance(market_breadth, dict) else None
            if isinstance(buckets, (list, tuple)) and len(buckets) == len(MARKET_BREADTH_BUCKETS):
                counts = tuple(max(0, int(value)) for value in buckets)
                peak = max(counts, default=0)
                self.breadth_rows = tuple(
                    (label, str(count), (count / peak) if peak > 0 else 0.0, role)
                    for (label, role), count in zip(MARKET_BREADTH_BUCKETS, counts)
                )
                self.breadth_summary = (
                    int(market_breadth.get("up", sum(counts[:5]))),
                    int(market_breadth.get("flat", counts[5])),
                    int(market_breadth.get("down", sum(counts[6:]))),
                )
            else:
                self.breadth_rows = tuple((label, "—", 0.0, role) for label, role in MARKET_BREADTH_BUCKETS)
                self.breadth_summary = None
            self.update()
            return
        self.breadth_rows = ()
        self.breadth_summary = None
        market, kind = instrument.get("market", ""), instrument.get("type")
        lot = volume_lot_size(kind, market)
        precision = price_precision(kind, market)
        previous = quote.get("prev_close", 0)
        purchase, sale = quote.get("purchaser_price", ()), quote.get("seller_price", ())
        auction = bool(purchase and sale and purchase[0] == sale[0] > 0)
        self.rows = []
        totals = []
        for side, indices, title in (("seller", range(4, -1, -1), "卖"), ("purchaser", range(5), "买")):
            prices = (list(quote.get(f"{side}_price", ())) + [0] * 5)[:5]
            volumes = (list(quote.get(f"{side}_vol", ())) + [0] * 5)[:5]
            if kind == "指":
                prices, volumes = [0] * 5, [0] * 5
            totals.append(sum(max(0, volume) for index, (price, volume) in enumerate(zip(prices, volumes))
                              if price > 0 or (auction and index == 1)))
            for index in indices:
                price, volume = prices[index], volumes[index]
                role = direction_color_role(price - previous) if price > 0 and previous else ""
                quantity = (format_value(volume, lot, not should_use_english_units(unit_mode, market))
                            if (price > 0 or (auction and index == 1)) and volume > 0 else "—")
                self.rows.append((f"{title}{index + 1}", f"{price:.{precision}f}" if price > 0 else "—", quantity, role))
        sell, buy = totals
        self.buy_fraction = buy / (buy + sell) if buy + sell else None
        self.update()

    def _flat_color(self, neutral):
        if self.neutral_color is not None:
            return QColor(self.neutral_color)
        return self.palette().color(QPalette.PlaceholderText)

    def _breadth_geometry(self, metrics):
        margin = 10
        content_width = self.width() - margin * 2
        header_gap = 4
        line = (self.height() - 12 - header_gap) / 12
        label_w = max(metrics.horizontalAdvance(label) for label, _ in MARKET_BREADTH_BUCKETS)
        count_w = max(
            metrics.horizontalAdvance("9999"),
            *(metrics.horizontalAdvance(row[1]) for row in self.breadth_rows),
        )
        gap = 6
        bar_x = margin + label_w + gap
        bar_max_w = max(8.0, content_width - label_w - count_w - gap * 2)
        count_x = margin + content_width - count_w
        rows_top = 6 + line + header_gap
        bar_h = max(6.0, min(10.0, line * 0.44))
        return {
            "margin": margin,
            "content_width": content_width,
            "header_y": 6.0,
            "header_gap": header_gap,
            "line": line,
            "label_w": label_w,
            "bar_x": bar_x,
            "bar_max_w": bar_max_w,
            "count_x": count_x,
            "count_w": count_w,
            "rows_top": rows_top,
            "bar_h": bar_h,
        }

    def _paint_breadth(self, painter, neutral):
        metrics = painter.fontMetrics()
        geom = self._breadth_geometry(metrics)
        margin, content_width = geom["margin"], geom["content_width"]
        header_y, line, header_gap = geom["header_y"], geom["line"], geom["header_gap"]
        flat_color = self._flat_color(neutral)
        sep_color = QColor(neutral)
        sep_color.setAlpha(150)

        if self.breadth_summary is not None:
            up_text, flat_text, down_text = (str(value) for value in self.breadth_summary)
            segments = (
                (up_text, self.up_color),
                (" / ", sep_color),
                (flat_text, flat_color),
                (" / ", sep_color),
                (down_text, self.down_color),
            )
        else:
            segments = (
                ("—", neutral),
                (" / ", sep_color),
                ("—", neutral),
                (" / ", sep_color),
                ("—", neutral),
            )

        base_font = painter.font()
        bold_font = painter.font()
        bold_font.setBold(True)
        bold_metrics = QFontMetrics(bold_font)
        total_w = sum((bold_metrics if idx % 2 == 0 else metrics).horizontalAdvance(text)
                      for idx, (text, _) in enumerate(segments))
        x = margin + max(0.0, (content_width - total_w) / 2)
        for idx, (text, color) in enumerate(segments):
            is_value = idx % 2 == 0
            painter.setFont(bold_font if is_value else base_font)
            seg_w = (bold_metrics if is_value else metrics).horizontalAdvance(text)
            painter.setPen(color)
            painter.drawText(QRectF(x, header_y, seg_w + 1, line), Qt.AlignVCenter | Qt.AlignLeft, text)
            x += seg_w
        painter.setFont(base_font)

        divider_y = header_y + line + header_gap / 2
        divider = QColor(self.palette().mid().color())
        divider.setAlpha(120)
        painter.setPen(divider)
        painter.drawLine(QPointF(margin, divider_y), QPointF(margin + content_width, divider_y))

        painter.setRenderHint(QPainter.Antialiasing, True)
        for index, (label, count_text, fraction, role) in enumerate(self.breadth_rows):
            y = geom["rows_top"] + index * line
            role_color = (
                self.up_color if role == COLOR_ROLE_UP
                else self.down_color if role == COLOR_ROLE_DOWN
                else flat_color
            )
            painter.setPen(neutral)
            painter.drawText(QRectF(margin, y, geom["label_w"] + 1, line), Qt.AlignVCenter | Qt.AlignLeft, label)
            if fraction > 0:
                bar_w = max(2.0, geom["bar_max_w"] * min(1.0, fraction))
                bar_rect = QRectF(geom["bar_x"], y + (line - geom["bar_h"]) / 2, bar_w, geom["bar_h"])
                painter.setPen(Qt.NoPen)
                painter.setBrush(role_color)
                painter.drawRoundedRect(bar_rect, 2.0, 2.0)
            painter.setPen(neutral if count_text == "—" else role_color)
            painter.drawText(QRectF(geom["count_x"], y, geom["count_w"], line),
                             Qt.AlignVCenter | Qt.AlignRight, count_text)

    def paintEvent(self, event):
        painter = QPainter(self)
        if not self.flat:
            painter.fillRect(self.rect(), self.palette().base())
        neutral = self.palette().text().color()
        painter.setPen(self.palette().mid().color())
        painter.drawLine(0, 0, 0, self.height())
        if self.breadth_mode:
            self._paint_breadth(painter, neutral)
            return
        margin, width = 10, self.width() - 20
        line = (self.height() - 12) / 11
        for index, (title, price, quantity, role) in enumerate(self.rows):
            y = 6 + (index + (index >= 5)) * line
            painter.setPen(neutral)
            painter.drawText(QRectF(margin, y, 28, line), Qt.AlignVCenter, title)
            color = (self.up_color if role == COLOR_ROLE_UP else self.down_color if role == COLOR_ROLE_DOWN else
                     self.neutral_color if self.neutral_color is not None else neutral)
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
