"""Live quote details shared by the chart popup and detached window."""

import math

from PySide6.QtCore import QEvent, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontInfo, QFontMetrics, QPainter, QPalette, QPen
from PySide6.QtWidgets import QLabel, QWidget

from stockwidget.core.quote_presentation import (
    COLOR_ROLE_DOWN, COLOR_ROLE_NEUTRAL, COLOR_ROLE_UP, MARKET_BREADTH_BUCKETS,
    MARKET_OVERVIEW_INDICES, QuoteDisplayOptions, compute_market_turnover, direction_color_role,
    format_overview_index, format_quote, format_signed_value, format_value, is_cn_index,
    price_precision, quote_detail_values, should_use_english_units, volume_lot_size,
)
from stockwidget.ui.controls.quote_view import DEFAULT_DOWN_COLOR, DEFAULT_UP_COLOR


def _set_color(label, role, up, down):
    palette = QPalette(label.parentWidget().palette())
    color = up if role == COLOR_ROLE_UP else down if role == COLOR_ROLE_DOWN else palette.windowText().color()
    palette.setColor(QPalette.WindowText, QColor(color))
    label.setPalette(palette)


def _paint_card_surface(painter, rect, palette):
    """Draw a clean rounded card surface with a soft hairline border."""
    card = QRectF(rect).adjusted(0.5, 0.5, -0.5, -0.5)
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setBrush(palette.base())
    border = QColor(palette.mid().color())
    painter.setPen(QPen(border, 1))
    painter.drawRoundedRect(card, 10.0, 10.0)
    painter.restore()


class HistoryQuoteHeader(QWidget):
    """Consume the presenter's accepted quote; never issue a quote request."""

    def paintEvent(self, event):
        painter = QPainter(self)
        _paint_card_surface(painter, self.rect(), self.palette())

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
        self._sync_size()

    def set_breadth_mode(self, enabled):
        breadth = bool(enabled)
        changed = self.breadth_mode != breadth
        self.breadth_mode = breadth
        self.setAccessibleName("全市场涨跌分布" if breadth else "五档买卖行情")
        if changed:
            self._sync_size()
            self.update()

    def unified_minimum_size(self) -> tuple[int, int]:
        metrics = self.fontMetrics()
        bold_font = self.font()
        bold_font.setBold(True)
        bold_metrics = QFontMetrics(bold_font)
        label_width = max(metrics.horizontalAdvance(label) for label, _ in MARKET_BREADTH_BUCKETS)
        count_width = metrics.horizontalAdvance("9999")
        summary_width = bold_metrics.horizontalAdvance("涨9999 平999 跌9999")
        width = max(164, 20 + label_width + count_width + 12 + 48, 20 + summary_width)
        row_height = max(16, metrics.height() + 1)
        height = max(200, 12 + 4 + row_height * 12)
        return width, height

    def _sync_size(self):
        breadth_width, unified_height = self.unified_minimum_size()
        self.setFixedWidth(breadth_width if self.breadth_mode else 164)
        self.setMinimumHeight(unified_height)

    @property
    def breadth_summary_text(self) -> str:
        if self.breadth_summary is None:
            return "涨— 平— 跌—"
        up, flat, down = self.breadth_summary
        return f"涨{up} 平{flat} 跌{down}"

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() in {QEvent.FontChange, QEvent.DevicePixelRatioChange}:
            self._sync_size()

    def set_quote(self, instrument, quote, unit_mode, up, down, *, market_breadth=None, force_breadth=False):
        quote = quote or {}
        self.up_color, self.down_color = QColor(up), QColor(down)
        self.set_breadth_mode(force_breadth or is_cn_index(instrument))
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
        count_w = max((
            metrics.horizontalAdvance("9999"),
            *(metrics.horizontalAdvance(row[1]) for row in self.breadth_rows),
        ))
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

        if self.breadth_summary is not None:
            up_value, flat_value, down_value = self.breadth_summary
            segments = (
                (f"涨{up_value}", self.up_color),
                (" ", neutral),
                (f"平{flat_value}", flat_color),
                (" ", neutral),
                (f"跌{down_value}", self.down_color),
            )
        else:
            segments = (
                ("涨—", neutral),
                (" ", neutral),
                ("平—", neutral),
                (" ", neutral),
                ("跌—", neutral),
            )

        base_font = painter.font()
        bold_font = painter.font()
        bold_font.setBold(True)
        bold_metrics = QFontMetrics(bold_font)
        total_w = sum((bold_metrics if idx % 2 == 0 else metrics).horizontalAdvance(text)
                      for idx, (text, _) in enumerate(segments))
        use_bold = total_w <= content_width
        if not use_bold:
            total_w = sum(metrics.horizontalAdvance(text) for text, _ in segments)
        x = margin + max(0.0, (content_width - total_w) / 2)
        for idx, (text, color) in enumerate(segments):
            is_value = use_bold and idx % 2 == 0
            painter.setFont(bold_font if is_value else base_font)
            seg_w = (bold_metrics if is_value else metrics).horizontalAdvance(text)
            painter.setPen(color)
            painter.drawText(QRectF(x, header_y, seg_w + 1, line), Qt.AlignVCenter | Qt.AlignLeft, text)
            x += seg_w
        painter.setFont(base_font)

        divider_y = header_y + line + header_gap / 2
        divider = QColor(neutral)
        divider.setAlpha(46 if not self.flat else 64)
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
        neutral = self.palette().text().color()
        if not self.flat:
            _paint_card_surface(painter, self.rect(), self.palette())
        else:
            divider = QColor(neutral)
            divider.setAlpha(60)
            painter.setPen(divider)
            painter.drawLine(QPointF(0.5, 6), QPointF(0.5, self.height() - 6))
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


def _scaled_font(base: QFont, factor: float, *, bold: bool = False) -> QFont:
    font = QFont(base)
    font.setBold(bold)
    if base.pointSizeF() > 0:
        font.setPointSizeF(max(6.0, base.pointSizeF() * factor))
    elif base.pixelSize() > 0:
        font.setPixelSize(max(8, round(base.pixelSize() * factor)))
    else:
        px = QFontInfo(base).pixelSize()
        if px > 0:
            font.setPixelSize(max(8, round(px * factor)))
    return font


def _fit_font(font: QFont, text: str, max_width: float) -> QFont:
    metrics = QFontMetrics(font)
    width = metrics.horizontalAdvance(text)
    if width <= max_width or max_width <= 8:
        return font
    return _scaled_font(font, max(0.6, max_width / width), bold=font.bold())


class MarketOverviewWidget(QWidget):
    """Nine CN benchmark index cards (3x3) plus optional full-market turnover summary."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.flat = True
        self.show_turnover = True
        self.neutral_color = None
        self.up_color, self.down_color = QColor(DEFAULT_UP_COLOR), QColor(DEFAULT_DOWN_COLOR)
        self.index_cards = tuple(
            (key, format_overview_index(instrument, None))
            for key, instrument in MARKET_OVERVIEW_INDICES
        )
        self.turnover_values = {"today": None, "yesterday": None, "delta": None, "date": None, "prev_date": None}
        self.turnover_items = (
            ("今日实时", "—", ""),
            ("昨日总额", "—", ""),
            ("较上日此时", "—", ""),
        )
        self.setAccessibleName("大盘总览")
        self._sync_size()

    def set_show_turnover(self, enabled: bool) -> bool:
        show = bool(enabled)
        if self.show_turnover == show:
            return False
        self.show_turnover = show
        self._sync_size()
        self.update()
        return True

    def _fonts(self):
        base = self.font()
        return {
            "name": _scaled_font(base, 0.94, bold=False),
            "price": _scaled_font(base, 1.08, bold=True),
            "change": _scaled_font(base, 0.86, bold=False),
            "section": _scaled_font(base, 0.86, bold=False),
            "value": _scaled_font(base, 0.96, bold=True),
        }

    def _sync_size(self):
        # Responsive painted surface: stays within the unified popup minimum so
        # entering overview or toggling turnover never forces window resizing.
        self.setMinimumSize(220, 180)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() in {QEvent.FontChange, QEvent.DevicePixelRatioChange}:
            self._sync_size()

    def set_overview(self, overview_quotes, turnover_bars, unit_mode, up, down):
        self.up_color, self.down_color = QColor(up), QColor(down)
        quotes = overview_quotes or {}
        self.index_cards = tuple(
            (key, format_overview_index(instrument, quotes.get(key)))
            for key, instrument in MARKET_OVERVIEW_INDICES
        )
        turnover = compute_market_turnover(quotes, turnover_bars)
        self.turnover_values = turnover
        unit_cn = not should_use_english_units(unit_mode, "sh")
        today_text = (format_value(turnover["today"], 1, unit_cn=unit_cn)
                      if turnover["today"] is not None else "—")
        yesterday_text = (format_value(turnover["yesterday"], 1, unit_cn=unit_cn)
                          if turnover["yesterday"] is not None else "—")
        delta = turnover["delta"]
        delta_text = format_signed_value(delta, 1, unit_cn=unit_cn) if delta is not None else "—"
        delta_role = direction_color_role(delta) if delta is not None else ""
        self.turnover_items = (
            ("今日实时", today_text, ""),
            ("昨日总额", yesterday_text, ""),
            ("较上日此时", delta_text, delta_role),
        )
        self.update()

    def _role_color(self, role, neutral):
        if role == COLOR_ROLE_UP:
            return self.up_color
        if role == COLOR_ROLE_DOWN:
            return self.down_color
        if role == COLOR_ROLE_NEUTRAL:
            if self.neutral_color is not None:
                return QColor(self.neutral_color)
            return self.palette().color(QPalette.PlaceholderText)
        return neutral

    def _layout_rects(self):
        fonts = self._fonts()
        section_m = QFontMetrics(fonts["section"])
        value_m = QFontMetrics(fonts["value"])
        card_gap = 4.0
        section_gap = 5.0
        width = float(self.width())
        height = float(self.height())
        if self.show_turnover:
            ideal_turnover_h = float(8 + section_m.height() + 2 + section_m.height() + 1 + value_m.height())
            turnover_h = min(ideal_turnover_h, max(38.0, height * 0.28))
            grid_h = max(0.0, height - turnover_h - section_gap)
            turnover_rect = QRectF(0.0, grid_h + section_gap, width, max(0.0, height - grid_h - section_gap))
        else:
            grid_h = height
            turnover_rect = QRectF()
        card_w = max(0.0, (width - card_gap * 2.0) / 3.0)
        card_h = max(0.0, (grid_h - card_gap * 2.0) / 3.0)
        cards = tuple(
            QRectF((idx % 3) * (card_w + card_gap), (idx // 3) * (card_h + card_gap), card_w, card_h)
            for idx in range(len(MARKET_OVERVIEW_INDICES))
        )
        return cards, turnover_rect

    def _paint_tile(self, painter, rect, neutral):
        if not self.flat:
            _paint_card_surface(painter, rect, self.palette())
            return
        card = QRectF(rect).adjusted(0.5, 0.5, -0.5, -0.5)
        fill = QColor(neutral)
        fill.setAlpha(15)
        border = QColor(neutral)
        border.setAlpha(38)
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setBrush(fill)
        painter.setPen(QPen(border, 1.0))
        painter.drawRoundedRect(card, 8.0, 8.0)
        painter.restore()

    def paintEvent(self, event):
        painter = QPainter(self)
        neutral = self.palette().text().color()
        muted = QColor(neutral)
        muted.setAlpha(round(neutral.alpha() * 0.68))
        fonts = self._fonts()
        name_m = QFontMetrics(fonts["name"])
        price_m = QFontMetrics(fonts["price"])
        change_m = QFontMetrics(fonts["change"])
        section_m = QFontMetrics(fonts["section"])
        value_m = QFontMetrics(fonts["value"])

        cards, turnover_rect = self._layout_rects()
        for rect, (_key, item) in zip(cards, self.index_cards):
            self._paint_tile(painter, rect, neutral)
            inner = rect.adjusted(5.0, 3.0, -5.0, -3.0)
            if inner.width() <= 4 or inner.height() <= 4:
                continue
            card_name_font, card_price_font, card_change_font = fonts["name"], fonts["price"], fonts["change"]
            c_name_m, c_price_m, c_change_m = name_m, price_m, change_m
            line_sum = c_name_m.height() + c_price_m.height() + c_change_m.height()
            if line_sum > inner.height() and line_sum > 0:
                v_scale = max(0.68, inner.height() / line_sum)
                card_name_font = _scaled_font(card_name_font, v_scale, bold=card_name_font.bold())
                card_price_font = _scaled_font(card_price_font, v_scale, bold=card_price_font.bold())
                card_change_font = _scaled_font(card_change_font, v_scale, bold=card_change_font.bold())
                c_name_m = QFontMetrics(card_name_font)
                c_price_m = QFontMetrics(card_price_font)
                c_change_m = QFontMetrics(card_change_font)
            name_h = c_name_m.height()
            price_h = c_price_m.height()
            change_h = c_change_m.height()
            slack = max(0.0, inner.height() - (name_h + price_h + change_h))
            y_name = inner.top()
            y_price = inner.top() + name_h + slack * 0.45
            y_change = inner.bottom() - change_h

            painter.setFont(card_name_font)
            painter.setPen(neutral)
            elided_name = c_name_m.elidedText(item["name"], Qt.ElideRight, int(inner.width()))
            painter.drawText(QRectF(inner.left(), y_name, inner.width(), name_h),
                             Qt.AlignVCenter | Qt.AlignLeft, elided_name)

            role_color = self._role_color(item["role"], neutral)
            price_font = _fit_font(card_price_font, item["price"], inner.width())
            painter.setFont(price_font)
            painter.setPen(role_color)
            painter.drawText(QRectF(inner.left(), y_price, inner.width(), price_h),
                             Qt.AlignVCenter | Qt.AlignLeft, item["price"])

            change_font = card_change_font
            combined_w = (c_change_m.horizontalAdvance(item["change"])
                          + c_change_m.horizontalAdvance(item["change_pct"]) + 4)
            if combined_w > inner.width() and inner.width() > 8:
                change_font = _scaled_font(change_font, max(0.60, inner.width() / combined_w))
            painter.setFont(change_font)
            painter.setPen(role_color)
            change_rect = QRectF(inner.left(), y_change, inner.width(), change_h)
            painter.drawText(change_rect, Qt.AlignVCenter | Qt.AlignLeft, item["change"])
            painter.drawText(change_rect, Qt.AlignVCenter | Qt.AlignRight, item["change_pct"])

        if not self.show_turnover or turnover_rect.isEmpty():
            return
        self._paint_tile(painter, turnover_rect, neutral)
        t_inner = turnover_rect.adjusted(7.0, 4.0, -7.0, -4.0)
        if t_inner.width() <= 4 or t_inner.height() <= 4:
            return
        t_sec_font, t_val_font = fonts["section"], fonts["value"]
        t_sec_m, t_val_m = section_m, value_m
        t_sum = t_sec_m.height() * 2 + t_val_m.height() + 3
        if t_sum > t_inner.height() and t_sum > 0:
            t_scale = max(0.68, t_inner.height() / t_sum)
            t_sec_font = _scaled_font(t_sec_font, t_scale, bold=t_sec_font.bold())
            t_val_font = _scaled_font(t_val_font, t_scale, bold=t_val_font.bold())
            t_sec_m = QFontMetrics(t_sec_font)
            t_val_m = QFontMetrics(t_val_font)
        title_h = t_sec_m.height()
        caption_h = t_sec_m.height()
        value_h = t_val_m.height()
        t_slack = max(0.0, t_inner.height() - (title_h + caption_h + value_h))

        painter.setFont(t_sec_font)
        painter.setPen(muted)
        painter.drawText(QRectF(t_inner.left(), t_inner.top(), t_inner.width(), title_h),
                         Qt.AlignVCenter | Qt.AlignLeft, "全市场成交额")

        col_gap = 6.0
        col_w = max(0.0, (t_inner.width() - col_gap * 2.0) / 3.0)
        y_caption = t_inner.top() + title_h + t_slack * 0.6
        y_value = t_inner.bottom() - value_h
        for index, (caption, value_text, role) in enumerate(self.turnover_items):
            cx = t_inner.left() + index * (col_w + col_gap)
            painter.setFont(t_sec_font)
            painter.setPen(muted)
            painter.drawText(QRectF(cx, y_caption, col_w, caption_h),
                             Qt.AlignVCenter | Qt.AlignLeft, caption)
            val_font = _fit_font(t_val_font, value_text, col_w)
            painter.setFont(val_font)
            painter.setPen(self._role_color(role, neutral))
            painter.drawText(QRectF(cx, y_value, col_w, value_h),
                             Qt.AlignVCenter | Qt.AlignLeft, value_text)
