"""Resizable chart geometry, detachment, and accepted live quote presentation."""

from dataclasses import replace
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

from PySide6.QtCore import QEvent, QPoint, QPointF, Qt
from PySide6.QtGui import QColor, QMouseEvent, QPainter, QPixmap
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QStyle, QStyleOptionButton

from stockwidget.data.bars import Bar, BarResult
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import quote_moving_average, quote_volume_ratio
from stockwidget.ui.controls.history_chart import HistoryChart, MA_COLORS
from stockwidget.ui.controls.quote_view import DEFAULT_DOWN_COLOR, DEFAULT_UP_COLOR
from tests.data.test_bars import sample_bars
from tests.ui.test_history import HistoryTestCase


class HistoryDetailsTests(HistoryTestCase):
    def test_navigation_cycles_checked_canonical_order_and_keeps_window_view_and_size(self):
        _, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=1)
        for entry in window.watchlist.values():
            entry["name"] = "同名标的"
        window.quotes.set_sort("现价", Qt.DescendingOrder)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        history.dialog.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        detached.set_view("daily")
        self.wait_ready(window)
        geometry = detached.geometry()
        for button, code in ((detached.ui.btn_next, "600000"), (detached.ui.btn_previous, "600006"),
                             (detached.ui.btn_previous, "600005"), (detached.ui.btn_next, "600006")):
            button.click()
            self.wait_ready(window)
            self.assertEqual(history.instrument["code"], code)
            self.assertIs(history.dialog, detached)
            self.assertEqual(detached.geometry(), geometry)
            self.assertEqual(detached.view, "daily")
            self.assertEqual(detached.quote_header.code.text(), "SH" + code)
            self.assertEqual(button.toolTip(), "")

    def test_twenty_header_values_share_live_quotes_and_optional_detail_result(self):
        _, window = self.make_window(chart_enabled=True)
        details = {"annual_profit": 2_000_000, "ttm_profit": 4_000_000, "annualized_profit": 5_000_000, "book_value_per_share": 8,
                   "total_shares": 1000000, "float_shares": 500000}
        with patch("stockwidget.ui.floating.presenter.request_quote_details", return_value=(details, {"inside_volume": 50, "outside_volume": 50})) as request:
            history = window.history
            history.open_row("float", 0)
            self.wait_ready(window)
            history.dialog.pop_out_button.click()
            self.wait_ready(window)
            self.wait_until(lambda: not window.quotes._detail_busy)
            header = history.dialog.quote_header
            self.assertEqual(header.values["市盈（TTM）"].text(), "2.50")
            self.assertEqual(header.values["市盈（静）"].text(), "5.00")
            self.assertEqual(header.values["市净率"].text(), "1.25")
            self.assertEqual(header.values["总股本"].text(), "100.00万")
            self.assertEqual(header.values["流通股"].text(), "50.00万")
            self.assertEqual(header.values["市盈（动）"].text(), "2.00")
            self.assertNotIn("每股收益", header.values)
            for row, caption in enumerate(("市盈（静）", "市盈（动）", "市盈（TTM）", "市净率")):
                self.assertEqual(history.dialog.ui.quote_column_3_layout.itemAtPosition(row, 0).widget().text(), caption)
            for theme in ("light", "dark"):
                window.set_view_options(color_mode=theme)
                history.dialog.resize(780, 400)
                self.app.processEvents()
                for col in range(5):
                    block = getattr(history.dialog.ui, f"quote_column_{col}")
                    for row in range(4):
                        caption = block.layout().itemAtPosition(row, 0).widget()
                        value = block.layout().itemAtPosition(row, 1).widget()
                        self.assertTrue(block.rect().contains(caption.geometry()))
                        self.assertTrue(block.rect().contains(value.geometry()))
                        self.assertLess(caption.geometry().right(), value.geometry().left())
            quote = window.quotes._latest_quotes["sh600000"]
            data = {**window.quotes._latest_quotes, "sh600000": {**quote, "current_price": 12, "time": "10:00:01"}}
            window.quotes.accept_result((True, data, None))
            self.assertEqual(header.values["市盈（TTM）"].text(), "3.00")
            self.assertEqual(header.values["总市值"].text(), "1200.00万")
            self.wait_until(lambda: not window.quotes._detail_busy)
            self.assertEqual(request.call_count, 2)
            self.assertEqual(request.call_args.args[1], details)

    def test_daily_detail_cache_and_selection_close_reject_late_activity(self):
        _, window = self.make_window(chart_enabled=True)
        presenter = window.quotes
        first, second = list(window.checked_codes.values())[:2]
        cache = window.history.cache
        clock = [cache.clock()]
        cache.clock = lambda: clock[0]
        saved_at, cache_generation = clock[0].timestamp(), cache.generation
        values = {"annual_profit": 2_000_000}
        with patch.object(presenter, "_detail_worker") as worker:
            presenter.set_detail_instrument(first)
            self.wait_until(lambda: worker.call_count == 1)
            generation = presenter._detail_generation
            presenter._accept_details((generation, first, "2026-09-30", values, {"inside_volume": 50}, saved_at, cache_generation, True))
            key = presenter._detail_key(first)
            clock[0] += timedelta(hours=2)
            presenter._refresh_details()
            self.wait_until(lambda: worker.call_count == 2)
            self.assertEqual(worker.call_args.args[3], values)
            presenter._accept_details((generation, first, "2026-09-30", values, {}, saved_at, cache_generation, False))
            self.assertEqual(presenter._fundamentals[key][1], saved_at)
            clock[0] += timedelta(days=1)
            presenter._refresh_details()
            self.wait_until(lambda: worker.call_count == 3)
            self.assertIsNone(worker.call_args.args[3])
            presenter.set_detail_instrument(second)
            presenter._accept_details((generation, first, "2026-09-30", {"annual_profit": 999}, {"inside_volume": 999}, saved_at, cache_generation, True))
            self.wait_until(lambda: worker.call_count == 4)
            self.assertEqual(presenter._fundamentals[key][0], values)
            self.assertNotIn("inside_volume", presenter.details_for(second))
            generation = presenter._detail_generation
            presenter.set_detail_instrument(None)
            presenter._accept_details((generation, second, "2026-09-30", {"annual_profit": 3}, {"inside_volume": 3}, saved_at, cache_generation, True))
            self.assertNotIn(presenter._detail_key(second), presenter._fundamentals)

    def test_reopening_reuses_daily_disk_financials_and_clearing_rejects_inflight_writes(self):
        _, window = self.make_window(chart_enabled=True)
        presenter, cache = window.quotes, window.history.cache
        instrument = next(iter(window.checked_codes.values()))
        details = {"annual_profit": 2_000_000, "total_shares": 1_000_000}
        with patch("stockwidget.ui.floating.presenter.request_quote_details", return_value=(details, {})) as request:
            presenter.set_detail_instrument(instrument)
            self.wait_until(lambda: not presenter._detail_busy)
            self.assertIsNone(request.call_args.args[1])
            presenter.clear_details()
            restarted = BarCache(cache.directory, clock=cache.clock)
            window.history.cache = restarted
            presenter.set_detail_instrument(instrument)
            self.wait_until(lambda: not presenter._detail_busy)
            self.assertEqual(request.call_args.args[1], details)
            self.assertEqual(request.call_count, 2)
        with patch.object(presenter, "_detail_worker") as worker:
            presenter._refresh_details()
            self.wait_until(lambda: worker.called)
            generation, cache_generation = presenter._detail_generation, restarted.generation
            self.assertTrue(window.history.clear_cache())
            presenter._accept_details((generation, instrument, "2026-09-30", details, {},
                                       restarted.clock().timestamp(), cache_generation, True))
            self.assertFalse(restarted.database.exists())
            self.assertFalse(presenter.details_for(instrument))

    def test_hover_readout_and_crosshair_leave_restore_latest_without_tooltip(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog
        popup.resize(560, 380)
        self.app.processEvents()
        chart = popup.chart
        latest = chart.bars[-1]
        selected = next(i for i, bar in enumerate(chart.bars) if bar.time.endswith("09:40:00"))
        bars = list(chart.bars)
        bars[selected] = replace(bars[selected], close=10.25, volume=300, amount=3000)
        chart.set_data(bars, "intraday", history.instrument, reference_price=9, now=history.cache.clock())
        self.app.processEvents()
        geometry = (popup.geometry(), chart.geometry(), popup.ui.chart_readout.geometry())
        baseline = chart.grab().toImage()
        point = QPointF(chart._plot.left() + chart._plot.width() * (chart._positions[selected] + .5) / len(chart._time_axis),
                        chart._plot.top() + chart._plot.height() * .31)
        event = QMouseEvent(QEvent.MouseMove, point, QPointF(chart.mapToGlobal(point.toPoint())), Qt.NoButton, Qt.NoButton, Qt.NoModifier)
        with patch("PySide6.QtWidgets.QToolTip.showText") as tooltip:
            self.app.sendEvent(chart, event)
            self.app.processEvents()
            tooltip.assert_not_called()
        self.assertEqual(chart._hover, selected)
        self.assertEqual((popup.geometry(), chart.geometry(), popup.ui.chart_readout.geometry()), geometry)
        self.assertEqual(popup.ui.chart_readout.text(), "均 10.00  价格 10.25 (+13.89%)  量 3")
        hovered = chart.grab().toImage()
        scale = hovered.devicePixelRatio()
        x, y = round(point.x() * scale), round(point.y() * scale)
        self.assertTrue(any(hovered.pixelColor(px, y) != baseline.pixelColor(px, y)
                            for px in range(round(chart._plot.left() * scale), round(chart._plot.right() * scale))))
        self.assertTrue(any(hovered.pixelColor(x, py) != baseline.pixelColor(x, py)
                            for py in range(round(chart._plot.top() * scale), round(chart._plot.bottom() * scale))))
        self.app.sendEvent(chart, QEvent(QEvent.Leave))
        self.assertIn(f"价格 {chart._price(latest.close)}", popup.ui.chart_readout.text())
        self.assertIsNone(chart._cursor_pos)
        popup.set_view("daily")
        self.wait_ready(window)
        readout = popup.ui.chart_readout.text().splitlines()
        self.assertEqual(len(readout), 2)
        self.assertIn("量 ", readout[0])
        self.assertNotIn("均", "".join(readout))
        self.assertTrue(readout[1].startswith("MA5:"))
        chart.grab()
        point = chart._plot.center()
        event = QMouseEvent(QEvent.MouseMove, point, QPointF(chart.mapToGlobal(point.toPoint())), Qt.NoButton, Qt.NoButton, Qt.NoModifier)
        self.app.sendEvent(chart, event)
        image = chart.grab().toImage()
        scale = image.devicePixelRatio()
        # Daily candles have only a left price scale, so no cropped percentage
        # badge may spill into the narrow right margin.
        self.assertEqual(image.pixelColor(round((chart.width() - 4) * scale), round(point.y() * scale)), chart.palette().base().color())

    def test_ratio_and_trade_direction_colors_follow_global_up_down_colors(self):
        _, window = self.make_window(chart_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        window.history.dialog.pop_out_button.click()
        self.wait_ready(window)
        header = window.history.dialog.quote_header
        instrument = window.history.instrument
        quote = window.quotes.quote_for(instrument)
        for ratio, expected in ((2, QColor("blue")), (.5, QColor("orange")), (1, None)):
            header.set_quote(instrument, quote, "cn", "blue", "orange", ratio, {"inside_volume": 50, "outside_volume": 50})
            self.assertEqual(header.values["外盘"].palette().windowText().color(), QColor("blue"))
            self.assertEqual(header.values["内盘"].palette().windowText().color(), QColor("orange"))
            if expected:
                self.assertEqual(header.values["量比"].palette().windowText().color(), expected)
            else:
                self.assertNotIn(header.values["量比"].palette().windowText().color(), (QColor("blue"), QColor("orange")))

    def test_zero_center_and_five_session_dates_are_centered_under_their_columns(self):
        from shiboken6 import delete
        chart = HistoryChart()
        try:
            chart.resize(740, 420)
            days = ("2026-10-22", "2026-10-23", "2026-10-26", "2026-10-27", "2026-10-28")
            bars = tuple(Bar(day + " " + time, 10, 12, 9, 11, 100, 1100)
                         for day in days for time in ("09:31:00", "09:32:00", "15:00:00"))
            chart.set_data(bars[-3:], "intraday", {"market": "sh", "type": "沪"}, reference_price=10,
                           now=datetime(2026, 10, 28, 2, tzinfo=timezone.utc))
            chart.grab()
            self.assertAlmostEqual(chart._reference_y, chart._plot.center().y())
            chart.set_data(bars, "five_day", {"market": "sh", "type": "沪"}, reference_price=10)
            chart.grab()
            labels = chart._time_labels()
            self.assertEqual([label for rect, label in labels], [day[5:] for day in days])
            for i, (rect, label) in enumerate(labels):
                self.assertAlmostEqual(rect.center().x(), chart._plot.left() + chart._plot.width() * (i + .5) / 5)
            point = QPointF(chart._plot.left() + chart._plot.width() * (chart._positions[0] + .5) / len(chart._time_axis),
                            chart._plot.center().y())
            event = QMouseEvent(QEvent.MouseMove, point, QPointF(chart.mapToGlobal(point.toPoint())),
                                Qt.NoButton, Qt.NoButton, Qt.NoModifier)
            self.app.sendEvent(chart, event)
            texts = []

            class Recorder(QPainter):
                def drawText(self, *args):
                    texts.append(args[-1])
                    return super().drawText(*args)

            with patch("stockwidget.ui.controls.history_chart.QPainter", Recorder):
                chart.grab()
            self.assertIn("10-22 09:31", texts)
            self.assertEqual(chart.readout_text(), "均 11.00  价格 11.00 (+10.00%)  量 1")
        finally:
            delete(chart)

    def test_panel_buttons_use_hover_for_selected_view_and_quote_columns_align(self):
        _, window = self.make_window(chart_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        for theme in ("light", "dark"):
            window.set_view_options(color_mode=theme)
            self.app.processEvents()
            active = popup.view_buttons["intraday"]
            other = popup.view_buttons["five_day"]
            def background(button, hover=False):
                if hover:
                    ratio = button.devicePixelRatioF()
                    pixmap = QPixmap(round(button.width() * ratio), round(button.height() * ratio))
                    pixmap.setDevicePixelRatio(ratio)
                    pixmap.fill(button.palette().window().color())
                    option = QStyleOptionButton()
                    button.initStyleOption(option)
                    option.state |= QStyle.State_MouseOver
                    painter = QPainter(pixmap)
                    button.style().drawControl(QStyle.CE_PushButton, option, painter, button)
                    painter.end()
                    image = pixmap.toImage()
                else:
                    image = button.grab().toImage()
                scale = image.devicePixelRatio()
                return image.pixelColor(round(button.width() / 2 * scale), round(3 * scale))
            selected = background(active)
            self.assertEqual(selected, background(other, hover=True))
            close = background(popup.close_button)
            self.assertGreater(close.red(), close.green())
            for button in (popup.close_button, popup.pop_out_button):
                self.assertEqual(button.width(), button.height())
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = window.history.dialog
        for name, caption in (("amount", "成交额"), ("volume", "成交量"), ("average", "均价"), ("volume_ratio", "量比"),
                              ("high", "最高价"), ("low", "最低价"), ("open", "开盘价")):
            label = getattr(detached.ui, f"quote_{name}_caption")
            value = detached.quote_header.values[caption]
            self.assertLess(label.geometry().right(), value.geometry().left())
            self.assertLessEqual(abs(label.geometry().center().y() - value.geometry().center().y()), 1)

    def test_compact_toolbar_and_daily_ma_controls_save_preferences_without_refetch(self):
        settings, window = self.make_window(chart_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        popup.resize(300, 210)
        self.app.processEvents()
        self.assertLessEqual(popup.width(), 320)
        self.assertLess(popup.height(), 260)
        for button in popup.view_buttons.values():
            self.assertLessEqual(button.width(), button.fontMetrics().horizontalAdvance(button.text()) + 24)
        for control in (*popup.view_buttons.values(), popup.ui.cb_volume, popup.ui.cb_average, popup.ui.cb_order_book):
            point = control.mapTo(popup, control.rect().bottomRight())
            self.assertTrue(popup.rect().contains(point))
            self.assertEqual(control.toolTip(), "")
        self.assertTrue(popup.ui.ma_row.isHidden())
        popup.set_view("daily")
        self.wait_ready(window)
        self.assertTrue(popup.ui.ma_row.isVisible())
        self.assertFalse(popup.ui.cb_order_book.isEnabled())
        calls = window.history.cache.get.call_count
        popup.ui.cb_ma60.setChecked(False)
        popup.ui.cb_average.setChecked(False)
        self.assertTrue(popup.ui.ma_row.isHidden())
        self.assertEqual(popup.chart.periods, set())
        popup.ui.cb_average.setChecked(True)
        self.assertEqual(popup.chart.periods, {5, 10, 20, 30})
        self.assertTrue(popup.ui.ma_row.isVisible())
        for period, color in zip((5, 10, 20, 30, 60), MA_COLORS):
            self.assertEqual(getattr(popup.ui, f"cb_ma{period}").palette().windowText().color(), QColor(color))
        popup.ui.cb_volume.setChecked(False)
        self.assertFalse(window.current_config()["chart_volume_enabled"])
        self.assertEqual(window.history.cache.get.call_count, calls)
        self.assertEqual(popup.order_book.toolTip(), "")
        self.assertEqual(popup.quote_header.values["量比"].toolTip(), "")
        self.assertFalse(hasattr(settings.ui, "cb_chart_ma5"))

    def test_escape_cancels_a_border_resize_then_closes_popup(self):
        _, window = self.make_window(chart_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        original = popup.geometry()
        start = QPoint(2, 2)
        QTest.mousePress(popup, Qt.LeftButton, pos=start)
        QTest.mouseMove(popup, start - QPoint(20, 20))
        self.assertNotEqual(popup.geometry(), original)
        QTest.keyClick(popup, Qt.Key_Escape)
        self.assertTrue(popup.isVisible())
        self.assertEqual(popup.geometry(), original)
        QTest.mouseRelease(popup, Qt.LeftButton, pos=start)
        QTest.keyClick(popup, Qt.Key_Escape)
        self.assertFalse(popup.isVisible())

    def test_drag_resize_remembers_both_sizes_and_quotes_do_not_reset_geometry(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog
        before = popup.size()
        corner = QPoint(popup.width() - 3, popup.height() - 3)
        QTest.mousePress(popup, Qt.LeftButton, pos=corner)
        QTest.mouseMove(popup, corner + QPoint(55, 45))
        QTest.mouseRelease(popup, Qt.LeftButton, pos=corner + QPoint(55, 45))
        self.assertEqual(popup.width(), before.width() + 55)
        self.assertEqual(popup.height(), before.height() + 45)
        saved = [popup.width(), popup.height()]
        self.assertEqual(window.current_config()["chart_popup_size"], saved)
        geometry = popup.geometry()
        history._quotes_updated()
        self.assertEqual(popup.geometry(), geometry)
        popup.close()
        history.open_row("float", 0)
        self.assertEqual([popup.width(), popup.height()], saved)
        QTest.mouseClick(popup.pop_out_button, Qt.LeftButton)
        self.wait_ready(window)
        detached = history.dialog
        detached.resize(780, 470)
        self.wait_until(lambda: window.current_config()["chart_window_size"] == [780, 470])
        self.assertEqual(window.current_config()["chart_popup_size"], saved)
        detached.close()
        history.open_row("float", 1)
        self.assertEqual(history.dialog.display_mode, "popup")
        self.assertEqual([history.dialog.width(), history.dialog.height()], saved)
        history.dialog.pop_out_button.click()
        self.wait_ready(window)
        self.assertEqual([history.dialog.width(), history.dialog.height()], [780, 470])

    def test_sizes_are_bounded_and_order_book_fits_at_minimum_size(self):
        dialog, window = self.make_window(chart_enabled=True, chart_popup_size=[16000, 16000])
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog
        bounds = self.app.primaryScreen().availableGeometry()
        self.assertTrue(bounds.contains(popup.geometry()))
        popup.resize(1, 1)
        self.assertEqual(popup.size(), popup.minimumSize())
        popup.ui.cb_order_book.setChecked(True)
        self.app.processEvents()
        self.assertFalse(popup.order_book.isHidden())
        self.assertTrue(popup.rect().contains(popup.order_book.geometry()))
        self.assertTrue(popup.rect().contains(popup.chart.geometry()))
        self.assertLess(popup.order_book.geometry().bottom(), popup.ui.chart_readout.geometry().top())
        self.assertLess(popup.chart.geometry().bottom(), popup.ui.chart_readout.geometry().top())
        self.assertGreaterEqual(popup.chart.width(), 200)
        popup.set_view("five_day")
        self.wait_ready(window)
        self.assertTrue(popup.order_book.isHidden())
        popup.set_view("daily")
        self.wait_ready(window)
        self.assertTrue(popup.order_book.isHidden())
        popup.set_view("intraday")
        self.wait_ready(window)
        self.assertFalse(popup.order_book.isHidden())
        self.assertTrue(window.current_config()["chart_order_book_enabled"])
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        detached.resize(16000, 16000)
        self.app.processEvents()
        self.assertTrue(bounds.contains(detached.frameGeometry()))
        detached.resize(1, 1)
        self.app.processEvents()
        self.assertTrue(detached.rect().contains(detached.order_book.geometry()))
        self.assertTrue(detached.rect().contains(detached.quote_header.geometry()))
        self.assertLess(detached.order_book.geometry().bottom(), detached.ui.chart_readout.geometry().top())

    def test_pop_out_keeps_inflight_result_and_uses_one_worker(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        with patch.object(history, "_start") as start:
            history.open_row("float", 0)
            request = start.call_args.args[0]
            history._busy = True
            old = history.dialog
            old.pop_out_button.click()
            self.assertFalse(old.isVisible())
            detached = history.dialog
            self.assertTrue(detached.isVisible())
            self.assertIsNone(QApplication.activePopupWidget())
            start.assert_called_once()
            self.assertEqual(history._generation, request[0])
            history._accept((*request, BarResult(sample_bars(), "sina")))
            self.assertTrue(detached.chart.bars)
            self.assertEqual(start.call_count, 2)  # Existing worker finishes before daily details bootstrap.
            detached.close()
            self.assertFalse(detached.isVisible())

    def test_quote_header_and_depth_refresh_without_new_quote_requests(self):
        _, window = self.make_window(chart_enabled=True, chart_order_book_enabled=True)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        history.dialog.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        self.assertFalse(detached.quote_header.isHidden())
        values = dict(window.quotes._latest_quotes)
        values["sh600000"] = {**values["sh600000"], "time": "10:00:01", "current_price": 12,
                              "high_price": 13, "deals_vol": 6000, "deals_amt": 66000,
                              "purchaser_price": [11.9, 11.8, 11.7, 11.6, 11.5],
                              "purchaser_vol": [100, 200, 300, 400, 500],
                              "seller_price": [12.1, 12.2, 12.3, 12.4, 12.5],
                              "seller_vol": [200, 400, 600, 800, 1000]}
        with patch("stockwidget.ui.floating.presenter.fetch_quote_result") as request:
            window.quotes.accept_result((True, values, None))
            self.app.processEvents()
            request.assert_not_called()
        header, book = detached.quote_header, detached.order_book
        self.assertIn("标的0", header.name.text())
        self.assertEqual(header.values["现价"].text(), "12.00")
        self.assertEqual(header.values["涨跌"].text(), "+3.00")
        self.assertEqual(header.values["涨幅"].text(), "+33.33%")
        self.assertEqual(header.values["最高价"].text(), "13.00")
        self.assertEqual(header.values["最低价"].text(), "8.00")
        self.assertEqual(header.values["开盘价"].text(), "10.00")
        self.assertEqual(header.values["成交量"].text(), "60")
        self.assertEqual(header.values["均价"].text(), "11.00")
        self.assertNotEqual(header.values["量比"].text(), "—")
        self.assertEqual([row[0] for row in book.rows], ["卖5", "卖4", "卖3", "卖2", "卖1", "买1", "买2", "买3", "买4", "买5"])
        self.assertEqual(book.rows[0][1:3], ("12.50", "10"))
        self.assertEqual(book.rows[5][1:3], ("11.90", "1"))
        self.assertAlmostEqual(book.buy_fraction, 1 / 3)
        daily = history._series_for(history.instrument, "daily")
        daily.result = replace(daily.result, stale=True)
        history._update_details()
        self.assertEqual(header.values["量比"].text(), "—")
        values["sh600000"] = {**values["sh600000"], "time": "10:00:02", "current_price": 0,
                              "purchaser_price": [11] * 5, "seller_price": [11] * 5}
        window.quotes.accept_result((True, values, None))
        self.assertEqual(header.values["现价"].text(), "11.00")
        self.assertEqual(header.values["涨跌"].text(), "+2.00")

    def test_indicator_uses_fixed_buy_red_sell_green_and_missing_quotes_clear_depth(self):
        _, window = self.make_window(chart_enabled=True, chart_order_book_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        book = window.history.dialog.order_book
        instrument = window.history.instrument
        quote = {"purchaser_price": [10] * 5, "seller_price": [11] * 5,
                 "purchaser_vol": [100] * 5, "seller_vol": [300] * 5, "prev_close": 10}
        book.set_quote(instrument, quote, "cn", QColor("blue"), QColor("yellow"))
        image = book.grab().toImage()
        y = round(6 + (book.height() - 12) / 11 * 5.5)
        scale = image.devicePixelRatio()
        self.assertEqual(image.pixelColor(round(15 * scale), round(y * scale)), DEFAULT_UP_COLOR)
        self.assertEqual(image.pixelColor(round((book.width() - 15) * scale), round(y * scale)), DEFAULT_DOWN_COLOR)
        self.assertAlmostEqual(book.buy_fraction, .25)
        self.assertNotIn("%", book.toolTip())
        book.set_quote({**instrument, "type": "指"}, quote, "cn", "red", "green")
        self.assertIsNone(book.buy_fraction)
        self.assertTrue(all(row[1:3] == ("—", "—") for row in book.rows))
        book.set_quote(instrument, None, "cn", "red", "green")
        self.assertIsNone(book.buy_fraction)
        self.assertTrue(all(row[1:3] == ("—", "—") for row in book.rows))


class QuoteHistoryMetricTests(HistoryTestCase):
    def test_ratio_and_ma_share_daily_download_with_both_surfaces_and_chart(self):
        settings, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=1,
                                             taskbar_enabled=True, taskbar_sync_metrics=False,
                                             taskbar_metrics=["average", "volume_ratio"], taskbar_rows=1,
                                             taskbar_sync_paging=False, taskbar_page_mode="manual")
        window.set_visible_metrics(["name", "average", "volume_ratio"])
        settings.ui.cmb_average_period.setCurrentIndex(settings.ui.cmb_average_period.findData(5))
        history = window.history
        self.wait_until(lambda: history.cache.get.call_count == 1 and not history._busy)
        instrument = window.checked_codes[window.quotes._ordered_keys[0]]
        quote = window.quotes.quote_for(instrument)
        bars = history._series_for(instrument, "daily").result.bars
        mean = quote_moving_average(bars, quote, instrument, 5)
        ratio = quote_volume_ratio(bars, quote, instrument)
        self.assertEqual([window.model.headerData(i, Qt.Horizontal) for i in range(window.model.columnCount())],
                         ["名称", "均值", "量比"])
        self.assertEqual(window.quotes._last_full_rows[0]["均值"], f"{mean:.2f}")
        self.assertEqual(window.quotes._last_full_rows[0]["量比"], f"{ratio:.2f}")
        self.assertEqual([window.taskbar_model.index(0, i).data() for i in range(2)], [f"{mean:.2f}", f"{ratio:.2f}"])
        history.open_row("float", 0)
        self.wait_ready(window)
        history.dialog.set_view("daily")
        self.wait_ready(window)
        self.assertEqual(history.cache.get.call_count, 2)  # One daily and one shared minute request.
        history.dialog.pop_out_button.click()
        self.wait_ready(window)
        self.assertEqual(history.cache.get.call_count, 2)
        self.assertEqual(history.dialog.quote_header.values["量比"].text(), f"{ratio:.2f}")
        settings.ui.cmb_average_period.setCurrentIndex(settings.ui.cmb_average_period.findData(0))
        self.assertEqual(window.current_config()["average_period"], 0)
        self.assertEqual(window.quotes._last_full_rows[0]["均值"], "10.00")
        self.assertEqual(history.cache.get.call_count, 2)

    def test_sorting_derived_metric_loads_all_checked_rows_once_and_updates_live(self):
        _, window = self.make_window(float_paging_enabled=True, float_max_rows=1)
        history = window.history
        window.set_visible_metrics(["name", "average", "volume_ratio"])
        window.set_view_options(average_period=5)
        self.wait_until(lambda: history.cache.get.call_count == 1 and not history._busy)
        window.quotes.set_sort("均值", Qt.DescendingOrder)
        self.wait_until(lambda: history.cache.get.call_count == 7 and not history._busy)
        self.assertEqual(window.quotes.instrument_at("float", 0)["code"], "600006")
        data = dict(window.quotes._latest_quotes)
        data["sh600000"] = {**data["sh600000"], "current_price": 100, "high_price": 100, "time": "10:00:01"}
        window.quotes.accept_result((True, data, None))
        self.app.processEvents()
        self.assertEqual(window.quotes.instrument_at("float", 0)["code"], "600000")
        window.quotes.set_sort("量比", Qt.DescendingOrder)
        self.app.processEvents()
        self.assertEqual(history.cache.get.call_count, 7)
        self.assertTrue(all("量比" in row for row in window.quotes._last_sort_values))

    def test_daily_failure_and_late_result_after_clear_do_not_fabricate_values(self):
        _, window = self.make_window(float_paging_enabled=True, float_max_rows=1)
        history = window.history
        history.cache.get.side_effect = None
        history.cache.get.return_value = BarResult(message="下载失败")
        window.set_visible_metrics(["average", "volume_ratio"])
        window.set_view_options(average_period=60)
        self.wait_until(lambda: history.cache.get.call_count == 1 and not history._busy)
        self.assertEqual(window.quotes._last_full_rows[0]["均值"], "-")
        self.assertEqual(window.quotes._last_full_rows[0]["量比"], "-")
        data = dict(window.quotes._latest_quotes)
        window.quotes.accept_result((True, data, None))
        self.app.processEvents()
        self.assertEqual(history.cache.get.call_count, 1)
        with patch.object(history, "_start") as start:
            history.clear_cache()
            history._refresh_background()
            request = start.call_args.args[0]
            history._busy = True
            history.clear_cache()
            history._accept((*request, BarResult(sample_bars(), "sina")))
            self.assertFalse(history._series)
            self.assertEqual(window.quotes._last_full_rows[0]["均值"], "-")
