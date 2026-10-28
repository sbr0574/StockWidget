"""Resizable chart geometry, detachment, and accepted live quote presentation."""

from dataclasses import replace
from unittest.mock import patch

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QColor, QPainter, QPixmap
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QStyle, QStyleOptionButton

from stockwidget.data.bars import BarResult
from stockwidget.data.bars import quote_moving_average, quote_volume_ratio
from stockwidget.ui.controls.history_chart import MA_COLORS
from stockwidget.ui.controls.quote_view import DEFAULT_DOWN_COLOR, DEFAULT_UP_COLOR
from tests.data.test_bars import sample_bars
from tests.ui.test_history import HistoryTestCase


class HistoryDetailsTests(HistoryTestCase):
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
