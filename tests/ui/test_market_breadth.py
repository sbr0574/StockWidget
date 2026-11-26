"""CN index 11-tier market breadth refresh lifecycle, UI toggle, and rendering."""

from dataclasses import replace
from datetime import datetime, timedelta, timezone
import tempfile
from unittest.mock import Mock, patch

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtTest import QTest

from stockwidget.core.quote_presentation import MARKET_BREADTH_BUCKETS, MARKET_OVERVIEW_INDICES, MARKET_OVERVIEW_QUOTES
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import Bar, BarResult
from stockwidget.data.quotes import MARKET_BREADTH_KEY, _new_entry, parse_sina_market_breadth
from stockwidget.ui.controls.history_chart import VOLUME_HEIGHT_RATIO
from stockwidget.ui.floating.presenter import QuotePresenter
from tests.data.test_bars import sample_bars
from tests.ui.test_history import HistoryTestCase


_REAL_REFRESH = QuotePresenter.refresh
SAMPLE_BREADTH = parse_sina_market_breadth(
    "4632536,2026-11-23,179,306,508,1931,1573,165,756,187,27,10,5,14:45:04".split(",")
)


class MarketBreadthTests(HistoryTestCase):
    def make_mixed_window(self, **settings):
        codes = {
            "sh000001": {"code": "000001", "market": "sh", "type": "指", "name": "上证指数"},
            "sz399001": {"code": "399001", "market": "sz", "type": "指", "name": "深证成指"},
            "bj899050": {"code": "899050", "market": "bj", "type": "指", "name": "北证50"},
            "hkhsi": {"code": "HSI", "market": "hk", "type": "指", "name": "恒生指数"},
            "sh600000": {"code": "600000", "market": "sh", "type": "沪", "name": "浦发银行"},
        }
        watchlist = {key: {**value, "checked": True} for key, value in codes.items()}
        dialog, window = self._make_dialog(watchlist=watchlist, codes=codes, **settings)
        data = {
            key: _new_entry(
                value["name"], 100, 99, 101, 102, 98, 1000, 100000,
                [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-11-23", "14:45:00",
            )
            for key, value in codes.items()
        }
        window.quotes.accept_result((True, data, None))
        daily = sample_bars()
        shift = datetime(2026, 11, 23) - datetime.fromisoformat(daily[-1].time)
        daily = tuple(
            replace(bar, time=(datetime.fromisoformat(bar.time) + shift).strftime("%Y-%m-%d"))
            for bar in daily
        )
        start = datetime(2026, 11, 23, 9, 31)
        minutes = tuple(
            Bar((start + timedelta(minutes=i)).isoformat(" "), 100, 102, 98, 101, 10, 1000, 100)
            for i in range(30)
        )
        window.history.cache = BarCache(self.enterContext(tempfile.TemporaryDirectory()))
        window.history.cache.get = Mock(
            side_effect=lambda instrument, view, **kwargs: BarResult(
                daily if view == "daily" else minutes, "sina"
            )
        )
        window.history.cache.clock = lambda: datetime(2026, 11, 23, 6, 45, tzinfo=timezone.utc)
        window.show()
        self.app.processEvents()
        return dialog, window, data

    def test_cn_index_toggle_label_and_conditional_refresh_lifecycle(self):
        _, window, data = self.make_mixed_window(chart_enabled=True, chart_order_book_enabled=False)
        history = window.history

        # No chart open -> refresh must not include zdp_updown.
        self.assertFalse(history.needs_market_breadth())
        with patch.object(window.quotes, "_fetch_worker") as worker:
            _REAL_REFRESH(window.quotes)
            window.quotes._refresh_thread.join(timeout=2)
            self.assertFalse(worker.call_args.args[2])

        # Open CN index with breadth disabled -> top-right button says "涨跌分布", checkbox hidden in popup.
        window.quotes.refresh.reset_mock()
        history.open_row("float", 0)  # sh000001
        self.wait_ready(window)
        popup = history.dialog
        self.assertEqual(popup.windowTitle(), "上证指数")
        self.assertEqual(popup.title.text(), "上证指数")
        self.assertTrue(popup.ui.cb_order_book.isHidden())
        self.assertTrue(popup.ui.cb_volume.isHidden())
        self.assertEqual(popup.order_book_button.toolTip(), "涨跌分布")
        self.assertEqual(popup.order_book_button.accessibleName(), "涨跌分布")
        self.assertFalse(history.needs_market_breadth())
        self.assertTrue(popup.order_book.isHidden())

        # Enable "涨跌分布" via top-right button -> no immediate out-of-cadence quote fetch; shows placeholder "—".
        popup.order_book_button.click()
        self.app.processEvents()
        window.quotes.refresh.assert_not_called()

        self.assertTrue(history.needs_market_breadth())
        self.assertFalse(popup.order_book.isHidden())
        self.assertTrue(popup.order_book.breadth_mode)
        self.assertIsNone(popup.order_book.breadth_summary)
        self.assertEqual(popup.order_book.breadth_summary_text, "涨— 平— 跌—")
        self.assertEqual(
            [row[0] for row in popup.order_book.breadth_rows],
            [label for label, _ in MARKET_BREADTH_BUCKETS],
        )
        self.assertTrue(all(row[1] == "—" and row[2] == 0.0 for row in popup.order_book.breadth_rows))

        # Next global refresh includes breadth.
        with patch.object(window.quotes, "_fetch_worker") as worker:
            _REAL_REFRESH(window.quotes)
            window.quotes._refresh_thread.join(timeout=2)
            self.assertTrue(worker.call_args.args[2])

        # Delivering merged quote + breadth populates header and 11 tiers without leaking key into quotes.
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.app.processEvents()
        self.assertNotIn(MARKET_BREADTH_KEY, window.quotes._latest_quotes)
        self.assertEqual(popup.order_book.breadth_summary, (4497, 165, 985))
        self.assertEqual(popup.order_book.breadth_summary_text, "涨4497 平165 跌985")
        self.assertEqual(
            [int(row[1]) for row in popup.order_book.breadth_rows],
            list(SAMPLE_BREADTH["buckets"]),
        )
        self.assertAlmostEqual(popup.order_book.breadth_rows[3][2], 1.0)  # 2~5% peak (1931)

        # Popping out to detached window keeps label, breadth mode, and accepted data.
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        self.assertEqual(detached.display_mode, "window")
        self.assertEqual(detached.ui.cb_order_book.text(), "涨跌分布")
        self.assertTrue(history.needs_market_breadth())
        self.assertEqual(detached.order_book.breadth_summary, (4497, 165, 985))

        # Navigating to another CN index (sz399001 -> bj899050) keeps breadth mode and data.
        detached.ui.btn_next.click()
        self.wait_ready(window)
        self.assertEqual(history.instrument["code"], "399001")
        self.assertEqual(detached.ui.cb_order_book.text(), "涨跌分布")
        self.assertEqual(detached.order_book.breadth_summary, (4497, 165, 985))
        detached.ui.btn_next.click()
        self.wait_ready(window)
        self.assertEqual(history.instrument["code"], "899050")
        self.assertEqual(detached.ui.cb_order_book.text(), "涨跌分布")
        self.assertEqual(detached.order_book.breadth_summary, (4497, 165, 985))

        # Navigating to HK index reverts label to "五档", disables breadth refresh, and clears cached breadth.
        detached.ui.btn_next.click()
        self.wait_ready(window)
        self.assertEqual(history.instrument["market"], "hk")
        self.assertEqual(detached.ui.cb_order_book.text(), "五档")
        self.assertFalse(detached.order_book.breadth_mode)
        self.assertFalse(history.needs_market_breadth())
        self.assertIsNone(window.quotes.market_breadth())

        # Navigating to A stock keeps "五档" mode.
        detached.ui.btn_next.click()
        self.wait_ready(window)
        self.assertEqual(history.instrument["code"], "600000")
        self.assertEqual(detached.ui.cb_order_book.text(), "五档")
        self.assertFalse(history.needs_market_breadth())

        # Navigating back to CN index starts with placeholder until next refresh.
        detached.ui.btn_next.click()
        self.wait_ready(window)
        self.assertEqual(history.instrument["code"], "000001")
        self.assertEqual(detached.ui.cb_order_book.text(), "涨跌分布")
        self.assertIsNone(detached.order_book.breadth_summary)

    def test_view_switch_uncheck_and_close_discard_cached_and_late_breadth(self):
        _, window, data = self.make_mixed_window(chart_enabled=True, chart_order_book_enabled=True)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog

        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.assertEqual(popup.order_book.breadth_summary, (4497, 165, 985))

        # Switching to daily hides panel and clears breadth; returning to intraday shows "—" until refreshed.
        popup.set_view("daily")
        self.wait_ready(window)
        self.assertFalse(history.needs_market_breadth())
        self.assertIsNone(window.quotes.market_breadth())
        popup.set_view("intraday")
        self.wait_ready(window)
        self.assertIsNone(popup.order_book.breadth_summary)

        # Unchecking clears breadth and stops requesting it.
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.assertIsNotNone(window.quotes.market_breadth())
        popup.order_book_button.click()
        self.app.processEvents()
        self.assertFalse(history.needs_market_breadth())
        self.assertIsNone(window.quotes.market_breadth())

        # Closing discards breadth, and a late in-flight result with breadth must not repopulate it.
        popup.order_book_button.click()
        self.app.processEvents()
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        popup.close()
        self.app.processEvents()
        self.assertFalse(history.needs_market_breadth())
        self.assertIsNone(window.quotes.market_breadth())
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.assertIsNone(window.quotes.market_breadth())

    def test_breadth_bars_follow_custom_up_down_colors_and_fit_minimum_size(self):
        _, window, data = self.make_mixed_window(chart_enabled=True, chart_order_book_enabled=False)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog

        # Even when unchecked at minimum size, toolbar and all 6 top-right buttons fit inside popup.
        popup.resize(1, 1)
        self.app.processEvents()
        self.assertTrue(popup.ui.cb_average.isHidden())
        for control in (*popup.view_buttons.values(),
                        popup.volume_button, popup.order_book_button, popup.overview_button,
                        popup.pin_button, popup.pop_out_button, popup.close_button):
            self.assertTrue(popup.rect().contains(control.mapTo(popup, control.rect().bottomRight())))

        # Enable breadth and verify minimum layout aligns left column with right breadth and compresses height.
        popup.order_book_button.click()
        self.app.processEvents()
        popup.resize(1, 1)
        self.app.processEvents()
        book = popup.order_book
        self.assertTrue(popup.rect().contains(book.geometry()))
        self.assertEqual(popup.ui.toolbar.geometry().top(), book.geometry().top())
        self.assertEqual(popup.ui.chart_readout.geometry().bottom(), book.geometry().bottom())
        self.assertLess(popup.chart.geometry().bottom(), popup.ui.chart_readout.geometry().top())
        self.assertLess(popup.ui.chart_readout.geometry().right(), book.geometry().left())
        self.assertLess(popup.minimumHeight(), 300)
        geom = book._breadth_geometry(book.fontMetrics())
        self.assertGreaterEqual(geom["line"], book.fontMetrics().height())

        # Custom up/down/neutral colors must render directly onto the proportion bars.
        up_color, down_color, flat_color = QColor("#00c853"), QColor("#ff5252"), QColor("#b0bec5")
        window.set_up_color(up_color)
        window.set_down_color(down_color)
        window.set_neutral_color(flat_color)
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.app.processEvents()

        image = book.grab().toImage()
        scale = image.devicePixelRatio()
        geom = book._breadth_geometry(book.fontMetrics())
        for row_idx, expected in ((3, up_color), (5, flat_color), (6, down_color)):
            px = round((geom["bar_x"] + 2.5) * scale)
            py = round((geom["rows_top"] + (row_idx + 0.5) * geom["line"]) * scale)
            self.assertEqual(image.pixelColor(px, py), expected)

        # Swapping up/down colors updates the bars immediately.
        window.set_up_color(QColor("#dd2100"))
        window.set_down_color(QColor("#019933"))
        self.app.processEvents()
        swapped = book.grab().toImage()
        px = round((geom["bar_x"] + 2.5) * scale)
        py_up = round((geom["rows_top"] + 3.5 * geom["line"]) * scale)
        py_down = round((geom["rows_top"] + 6.5 * geom["line"]) * scale)
        self.assertEqual(swapped.pixelColor(px, py_up), QColor("#dd2100"))
        self.assertEqual(swapped.pixelColor(px, py_down), QColor("#019933"))

        # All-zero buckets render safely without division by zero.
        zeros = parse_sina_market_breadth(["4632600", "2026-11-23", *(["0"] * 11), "14:46:00"])
        window.quotes.accept_result((True, {**data, MARKET_BREADTH_KEY: zeros}, None))
        self.app.processEvents()
        self.assertEqual(book.breadth_summary, (0, 0, 0))
        self.assertTrue(all(row[1] == "0" and row[2] == 0.0 for row in book.breadth_rows))
        self.assertFalse(book.grab().isNull())

        # Detached window minimum layout also fits cleanly.
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        detached.resize(1, 1)
        self.app.processEvents()
        self.assertTrue(detached.rect().contains(detached.order_book.geometry()))
        self.assertLess(detached.order_book.geometry().bottom(), detached.ui.chart_readout.geometry().top())

    def test_popup_max_size_chart_stability_dark_gridlines_and_fixed_volume_ratio(self):
        _, window, _ = self.make_mixed_window(chart_enabled=True, chart_order_book_enabled=False, color_mode="dark")
        history = window.history
        history.open_row("float", 4)  # sh600000 stock
        self.wait_ready(window)
        popup = history.dialog
        base_min = popup.minimumSize()
        bounds = popup.screen().availableGeometry().size()
        expected_max = base_min * 2
        self.assertEqual(popup.maximumSize(), expected_max.boundedTo(bounds))

        # Toggling order book keeps popup maximum size capped at ~2x no-book minimum.
        popup.order_book_button.click()
        self.app.processEvents()
        self.assertEqual(popup.maximumSize(), expected_max.boundedTo(bounds).expandedTo(popup.minimumSize()).boundedTo(bounds))

        # Fixed top/bottom heights and zero chart height twitch during background reload.
        chart_h = popup.chart.height()
        title_h = popup.title_row.height()
        toolbar_h = popup.ui.toolbar.height()
        readout_h = popup.ui.chart_readout.height()
        series = history._series_for(history.instrument, "intraday")
        series.retry_after = 0
        series.repair_needed = True
        with patch.object(history, "_start"):
            history.reload()
            self.app.processEvents()
            self.assertTrue(popup.status.isHidden())
            self.assertEqual(popup.chart.height(), chart_h)
            self.assertEqual(popup.title_row.height(), title_h)
            self.assertEqual(popup.ui.toolbar.height(), toolbar_h)
            self.assertEqual(popup.ui.chart_readout.height(), readout_h)
        series.repair_needed = False

        # Fixed upper/lower allocation ratio across chart sizes.
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        ratios = []
        for size in ((800, 440), (960, 560)):
            detached.resize(*size)
            self.app.processEvents()
            detached.chart.grab()
            upper = detached.chart._plot.height()
            lower = detached.chart._volume_plot.height()
            ratios.append(lower / (upper + lower))
        for ratio in ratios:
            self.assertAlmostEqual(ratio, VOLUME_HEIGHT_RATIO, delta=0.01)

        # Dark-mode detached chart auxiliary gridlines are clearly visible against base surface.
        chart_img = detached.chart.grab().toImage()
        scale = chart_img.devicePixelRatio()
        base = detached.chart.palette().color(QPalette.Base)
        grid_y = round((detached.chart._plot.top() + detached.chart._plot.height() * 0.25) * scale)
        sample_x = round((detached.chart._plot.left() + detached.chart._plot.width() * 0.6) * scale)
        grid_colors = [chart_img.pixelColor(sample_x, row) for row in range(grid_y - 1, grid_y + 2)]
        max_contrast = max(abs(c.red() - base.red()) + abs(c.green() - base.green()) + abs(c.blue() - base.blue())
                           for c in grid_colors)
        self.assertGreater(max_contrast, 30)

        # Light-mode detached order book card has soft border (not harsh black) on its left edge.
        window.set_view_options(color_mode="light")
        self.app.processEvents()
        book_img = detached.order_book.grab().toImage()
        mid_y = round(detached.order_book.height() * 0.5 * book_img.devicePixelRatio())
        edge_color = book_img.pixelColor(0, mid_y)
        self.assertGreater(edge_color.lightness(), 180)


class MarketOverviewTests(HistoryTestCase):
    @staticmethod
    def _overview_payload(time_str="11:30:00", date="2026-11-25"):
        specs = {
            "sh000001": ("上证指数", 3522.59, 3548.37, 3512.66, 352_070_000_000.0),
            "sz399001": ("深证成指", 10272.19, 10423.60, 10300.01, 410_980_000_000.0),
            "sz399106": ("深证综指", 1940.08, 1965.43, 1949.22, 410_980_000_000.0),
            "sz399006": ("创业板指", 2233.18, 2278.99, 2278.99, 204_620_000_000.0),
            "sh000680": ("科创综指", 1369.21, 1400.63, 1420.13, 119_600_000_000.0),
            "bj899050": ("北证50", 782.74, 793.60, 782.97, 5_910_000_000.0),
            "sh000016": ("上证50", 2680.10, 2700.00, 2715.20, 95_000_000_000.0),
            "sh000300": ("沪深300", 3929.84, 3963.82, 3925.57, 212_610_000_000.0),
            "sh000905": ("中证500", 5800.00, 5820.00, 5860.50, 145_000_000_000.0),
            "sh000852": ("中证1000", 6100.00, 6150.00, 6120.25, 130_000_000_000.0),
        }
        return {
            key: _new_entry(
                name, opening, prev, current, max(opening, current), min(opening, current),
                100000, amt, [0] * 5, [0] * 5, [0] * 5, [0] * 5, date, time_str,
            )
            for key, (name, opening, prev, current, amt) in specs.items()
        }

    @staticmethod
    def _prev_day_minutes(per_min_by_key):
        def build(per_min):
            bars = []
            for hour, start_m, end_m in ((9, 31, 60), (10, 0, 60), (11, 0, 31), (13, 1, 60), (14, 0, 59)):
                for minute in range(start_m, end_m):
                    bars.append(Bar(f"2026-11-24 {hour:02d}:{minute:02d}:00", 100, 100, 100, 100, 10, per_min))
            bars.append(Bar("2026-11-24 15:00:00", 100, 100, 100, 100, 10, per_min * 2))
            return tuple(bars)
        return {key: build(val) for key, val in per_min_by_key.items()}

    def make_stock_only_window(self, **settings):
        codes = {
            "sh600000": {"code": "600000", "market": "sh", "type": "沪", "name": "浦发银行"},
            "sz000001": {"code": "000001", "market": "sz", "type": "深", "name": "平安银行"},
        }
        watchlist = {key: {**value, "checked": True} for key, value in codes.items()}
        dialog, window = self._make_dialog(watchlist=watchlist, codes=codes, **settings)
        data = {
            key: _new_entry(
                value["name"], 10.0, 9.9, 10.1, 10.2, 9.8, 1000, 100000,
                [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-11-25", "11:30:00",
            )
            for key, value in codes.items()
        }
        window.quotes.accept_result((True, data, None))

        prev_bars = self._prev_day_minutes({
            "sh000001": 3_500_000_000.0,   # full 8400亿, <=11:30 4200亿
            "sz399106": 4_500_000_000.0,   # full 10800亿, <=11:30 5400亿
            "bj899050": 100_000_000.0,     # full 240亿, <=11:30 120亿 -> full 1.94万亿, <=11:30 9720亿
        })
        stock_minutes = tuple(
            Bar((datetime(2026, 11, 25, 9, 31) + timedelta(minutes=i)).isoformat(" "), 10, 10.2, 9.8, 10.1, 10, 1000, 10)
            for i in range(30)
        )
        daily = sample_bars()

        def fake_get(instrument, view, **kwargs):
            key = f"{instrument.get('market', '')}{instrument.get('code', '')}"
            if view == "daily":
                return BarResult(daily, "sina")
            if key in prev_bars:
                return BarResult(prev_bars[key], "sina")
            return BarResult(stock_minutes, "sina")

        window.history.cache = BarCache(self.enterContext(tempfile.TemporaryDirectory()))
        window.history.cache.get = Mock(side_effect=fake_get)
        window.history.cache.clock = lambda: datetime(2026, 11, 25, 3, 30, tzinfo=timezone.utc)
        window.show()
        self.app.processEvents()
        return dialog, window, data

    def test_top_buttons_overview_nine_indices_volume_breadth_toggles_and_restore(self):
        _, window, stock_data = self.make_stock_only_window(
            chart_enabled=True, chart_order_book_enabled=False, chart_volume_enabled=True,
        )
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog

        # Title keeps security name only; 6 buttons sit in top-right with tooltips; popup hides volume/order_book checkboxes.
        self.assertEqual(popup.windowTitle(), "浦发银行")
        self.assertEqual(popup.title.text(), "浦发银行")
        self.assertTrue(popup.ui.cb_volume.isHidden())
        self.assertTrue(popup.ui.cb_average.isHidden())
        self.assertTrue(popup.ui.cb_order_book.isHidden())
        buttons = [
            popup.ui.title_layout.itemAt(i).widget()
            for i in range(1, popup.ui.title_layout.count())
        ]
        self.assertEqual(buttons, [
            popup.volume_button, popup.order_book_button, popup.overview_button,
            popup.pin_button, popup.pop_out_button, popup.close_button,
        ])
        self.assertEqual(
            [btn.toolTip() for btn in buttons],
            ["成交量", "五档", "大盘总览", "固定图表", "打开独立窗口", "关闭图表"],
        )
        self.assertFalse(popup.overview_mode)
        self.assertTrue(popup.market_overview.isHidden())

        # Switch to daily + shrink to small saved size so we verify full state & preferred-size restoration.
        popup.set_view("daily")
        self.wait_ready(window)
        self.assertFalse(popup.order_book_button.isEnabled())
        popup.resize(340, 250)
        self.app.processEvents()
        popup.save_size()
        saved_size = list(window.current_config()["chart_popup_size"])
        self.assertFalse(popup.ui.ma_row.isHidden())
        self.assertTrue(popup.order_book.isHidden())

        # Enter 大盘总览 while chart_order_book_enabled=False -> right breadth is hidden, order_book_button enabled & says "涨跌分布".
        window.quotes.refresh.reset_mock()
        popup.overview_button.click()
        self.wait_ready(window)
        self.assertTrue(popup.overview_mode)
        self.assertTrue(history.needs_market_overview())
        self.assertFalse(history.needs_market_breadth())
        self.assertEqual(popup.windowTitle(), "大盘总览")
        self.assertEqual(popup.title.text(), "大盘总览")
        self.assertFalse(popup.market_overview.isHidden())
        self.assertTrue(popup.market_overview.show_turnover)
        self.assertTrue(popup.order_book.isHidden())
        self.assertTrue(popup.order_book_button.isEnabled())
        self.assertEqual(popup.order_book_button.toolTip(), "涨跌分布")
        self.assertTrue(popup.chart.isHidden())
        self.assertTrue(popup.ui.toolbar.isHidden())
        self.assertTrue(popup.ui.ma_row.isHidden())
        self.assertTrue(popup.ui.chart_readout.isHidden())
        self.assertEqual(window.current_config()["chart_popup_size"], saved_size)
        window.quotes.refresh.assert_called_once()

        # When breadth toggle is off in overview, refresh merges the 10 overview symbols without zdp_updown.
        with patch.object(window.quotes, "_fetch_worker") as worker:
            _REAL_REFRESH(window.quotes)
            window.quotes._refresh_thread.join(timeout=2)
            req_codes, _gen, include_breadth = worker.call_args.args
            self.assertFalse(include_breadth)
            self.assertTrue(set(MARKET_OVERVIEW_QUOTES).issubset(req_codes))
            self.assertIn("sh600000", req_codes)

        # Turn ON 涨跌分布 switch in overview -> right breadth appears and next refresh includes zdp_updown.
        popup.order_book_button.click()
        self.app.processEvents()
        self.assertFalse(popup.order_book.isHidden())
        self.assertTrue(popup.order_book.breadth_mode)
        self.assertTrue(history.needs_market_breadth())
        with patch.object(window.quotes, "_fetch_worker") as worker:
            _REAL_REFRESH(window.quotes)
            window.quotes._refresh_thread.join(timeout=2)
            self.assertTrue(worker.call_args.args[2])

        # Deliver merged live quotes + breadth.
        overview_quotes = self._overview_payload()
        window.quotes.accept_result((True, {**stock_data, **overview_quotes, MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.app.processEvents()

        # Non-watchlist indices must not leak into _latest_quotes or main table.
        self.assertEqual(set(window.quotes._latest_quotes), {"sh600000", "sz000001"})
        self.assertEqual(window.model.rowCount(), 2)

        # Verify all 9 index cards in exact 3x3 order.
        cards = popup.market_overview.index_cards
        self.assertEqual([key for key, _ in cards], [key for key, _ in MARKET_OVERVIEW_INDICES])
        self.assertEqual(
            [item["name"] for _, item in cards],
            ["上证指数", "深证成指", "创业板指", "科创综指", "北证50", "上证50", "沪深300", "中证500", "中证1000"],
        )
        by_key = dict(cards)
        self.assertEqual(by_key["sh000001"]["price"], "3512.66")
        self.assertEqual(by_key["sh000001"]["change"], "-35.71")
        self.assertEqual(by_key["sh000001"]["change_pct"], "-1.01%")
        self.assertEqual(by_key["sh000001"]["role"], "down")
        self.assertEqual(by_key["sh000016"]["role"], "up")
        self.assertEqual(by_key["sh000905"]["role"], "up")
        self.assertEqual(by_key["sh000852"]["role"], "down")
        self.assertEqual(by_key["sz399006"]["role"], "neutral")

        # Verify full-market turnover values (today=7689.60亿, yesterday=1.94万亿, delta=-2030.40亿).
        items = dict((cap, (val, role)) for cap, val, role in popup.market_overview.turnover_items)
        self.assertEqual(items["今日实时"], ("7689.60亿", ""))
        self.assertEqual(items["昨日总额"], ("1.94万亿", ""))
        self.assertEqual(items["较上日此时"], ("-2030.40亿", "down"))

        # Verify breadth header uses space separator (no "/").
        self.assertEqual(popup.order_book.breadth_summary_text, "涨4497 平165 跌985")
        self.assertNotIn("/", popup.order_book.breadth_summary_text)

        # Turn OFF 成交量 switch in overview -> bottom turnover section disappears and 3x3 cards expand.
        cards_with_vol, turnover_with_vol = popup.market_overview._layout_rects()
        self.assertFalse(turnover_with_vol.isEmpty())
        popup.volume_button.click()
        self.app.processEvents()
        self.assertFalse(popup.market_overview.show_turnover)
        cards_no_vol, turnover_no_vol = popup.market_overview._layout_rects()
        self.assertTrue(turnover_no_vol.isEmpty())
        self.assertGreater(cards_no_vol[0].height(), cards_with_vol[0].height())
        popup.volume_button.click()
        self.app.processEvents()
        self.assertTrue(popup.market_overview.show_turnover)

        # Subsequent quote tick does not re-download already-loaded turnover baseline bars.
        calls_after_load = history.cache.get.call_count
        window.quotes.accept_result((True, {**stock_data, **self._overview_payload("11:30:05"), MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.wait_ready(window)
        self.assertEqual(history.cache.get.call_count, calls_after_load)

        # Unit mode switch updates turnover formatting without refetch.
        window.set_unit_mode("en")
        self.app.processEvents()
        items_en = dict((cap, val) for cap, val, _ in popup.market_overview.turnover_items)
        self.assertEqual(items_en["今日实时"], "768.96B")
        self.assertEqual(items_en["昨日总额"], "1.94T")
        self.assertEqual(items_en["较上日此时"], "-203.04B")
        window.set_unit_mode("auto")
        self.app.processEvents()

        # Turn breadth back off then exit overview -> restores daily chart, MA row, original title & saved size.
        popup.order_book_button.click()
        self.app.processEvents()
        popup.overview_button.click()
        self.wait_ready(window)
        self.assertFalse(popup.overview_mode)
        self.assertFalse(history.needs_market_overview())
        self.assertFalse(history.needs_market_breadth())
        self.assertEqual(popup.windowTitle(), "浦发银行")
        self.assertTrue(popup.market_overview.isHidden())
        self.assertFalse(popup.chart.isHidden())
        self.assertFalse(popup.ui.toolbar.isHidden())
        self.assertFalse(popup.ui.ma_row.isHidden())
        self.assertFalse(popup.ui.chart_readout.isHidden())
        self.assertTrue(popup.order_book.isHidden())
        self.assertEqual(popup.order_book_button.toolTip(), "五档")
        self.assertFalse(popup.order_book_button.isEnabled())
        self.assertEqual([popup.width(), popup.height()], saved_size)
        self.assertEqual(window.current_config()["chart_popup_size"], saved_size)
        self.assertEqual(window.quotes.market_overview_quotes(), {})
        self.assertIsNone(window.quotes.market_breadth())

    def test_overview_layout_screen_clamping_navigation_popout_and_cleanup(self):
        _, window, stock_data = self.make_stock_only_window(
            chart_enabled=True, chart_order_book_enabled=True, chart_volume_enabled=True,
        )
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog

        popup.overview_button.click()
        self.wait_ready(window)
        window.quotes.accept_result((True, {**stock_data, **self._overview_payload(), MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.app.processEvents()

        # Minimum size layout fits all 9 non-overlapping index cards (3x3), turnover box, and right breadth cleanly.
        for mode in ("dark", "light"):
            with self.subTest(color_mode=mode):
                window.set_view_options(color_mode=mode)
                popup.resize(1, 1)
                self.app.processEvents()
                overview = popup.market_overview
                book = popup.order_book
                self.assertTrue(popup.rect().contains(overview.geometry()))
                self.assertTrue(popup.rect().contains(book.geometry()))
                self.assertLess(overview.geometry().right(), book.geometry().left())
                cards, turnover_rect = overview._layout_rects()
                self.assertEqual(len(cards), 9)
                for rect in (*cards, turnover_rect):
                    self.assertTrue(overview.rect().contains(rect.toAlignedRect()))
                    self.assertGreater(rect.width(), 70)
                    self.assertGreater(rect.height(), 36)
                for i, first in enumerate(cards):
                    for second in cards[i + 1:]:
                        self.assertFalse(first.intersects(second))
                    self.assertLess(first.bottom(), turnover_rect.top())
                self.assertFalse(popup.grab().isNull())

        # Content-driven size expansion near screen bottom-right edge must keep entire popup on screen.
        bounds = popup.screen().availableGeometry()
        popup.set_overview_mode(False)
        window.set_view_options(chart_order_book_enabled=False, chart_volume_enabled=False)
        self.app.processEvents()
        popup.resize(popup.minimumSize())
        self.app.processEvents()
        popup.move(bounds.right() - popup.width() + 1, bounds.bottom() - popup.height() + 1)
        self.app.processEvents()
        # Expanding by enabling depth/breadth keeps popup inside screen.
        popup.order_book_button.click()
        self.app.processEvents()
        self.assertTrue(bounds.contains(popup.geometry()))
        # Expanding further by entering 3x3 overview + enabling volume keeps popup inside screen.
        popup.move(bounds.right() - popup.width() + 1, bounds.bottom() - popup.height() + 1)
        self.app.processEvents()
        popup.overview_button.click()
        self.wait_ready(window)
        self.assertTrue(bounds.contains(popup.geometry()))
        popup.move(bounds.right() - popup.width() + 1, bounds.bottom() - popup.height() + 1)
        self.app.processEvents()
        popup.volume_button.click()
        self.wait_ready(window)
        self.assertTrue(bounds.contains(popup.geometry()))

        # Dragging on market_overview moves popup; dragging/double-clicking overview_button does not move/hide.
        popup.move(bounds.topLeft())
        start = popup.market_overview.mapTo(popup, popup.market_overview.rect().center())
        original = popup.geometry()
        offset = QPoint(30, 20)
        handle = popup.windowHandle()
        QTest.mousePress(handle, Qt.LeftButton, pos=start)
        QTest.mouseMove(handle, start + offset)
        QTest.mouseRelease(handle, Qt.LeftButton, pos=start)
        self.assertEqual(popup.geometry(), original.translated(offset))

        for pinned in (False, True):
            with self.subTest(overview_button_pinned=pinned):
                popup.pin_button.setChecked(pinned)
                popup.set_overview_mode(False)
                self.app.processEvents()
                original = popup.geometry()
                start = popup.overview_button.mapTo(popup, popup.overview_button.rect().center())
                finish = popup.chart.mapTo(popup, popup.chart.rect().center())
                QTest.mousePress(handle, Qt.LeftButton, pos=start)
                QTest.mouseMove(handle, finish)
                QTest.mouseRelease(handle, Qt.LeftButton, pos=finish)
                self.assertEqual(popup.geometry(), original)
                QTest.mouseDClick(handle, Qt.LeftButton, pos=start)
                self.app.processEvents()
                self.assertTrue(popup.isVisible())
                self.assertTrue(window.widget_visible)
                self.assertEqual(popup.pos(), original.topLeft())
        popup.set_overview_mode(True)
        self.app.processEvents()

        # Clicking another watchlist row exits overview and shows that row's chart with name-only title.
        history.open_row("float", 1)
        self.wait_ready(window)
        self.assertFalse(popup.overview_mode)
        self.assertFalse(popup.overview_button.isChecked())
        self.assertEqual(popup.windowTitle(), "平安银行")

        # Popping out while in overview opens detached window in normal chart mode with checkboxes intact & name-only title.
        popup.overview_button.click()
        self.wait_ready(window)
        self.assertTrue(popup.overview_mode)
        popup.pop_out_button.click()
        self.wait_ready(window)
        detached = history.dialog
        self.assertEqual(detached.display_mode, "window")
        self.assertFalse(detached.overview_mode)
        self.assertEqual(detached.windowTitle(), "平安银行")
        self.assertTrue(detached.ui.cb_volume.isVisible())
        self.assertTrue(detached.ui.cb_order_book.isVisible())
        self.assertFalse(history.needs_market_overview())
        detached.close()

        # Reopen, enter overview, then clear_cache discards overview state and closes popup.
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog
        popup.overview_button.click()
        self.wait_ready(window)
        window.quotes.accept_result((True, {**stock_data, **self._overview_payload(), MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.assertTrue(history.clear_cache())
        self.assertFalse(popup.isVisible())
        self.assertEqual(window.quotes.market_overview_quotes(), {})
        self.assertIsNone(window.quotes.market_breadth())

        # Late quote result after closing must not repopulate overview quotes or breadth.
        window.quotes.accept_result((True, {**stock_data, **self._overview_payload(), MARKET_BREADTH_KEY: SAMPLE_BREADTH}, None))
        self.assertEqual(window.quotes.market_overview_quotes(), {})
        self.assertIsNone(window.quotes.market_breadth())
