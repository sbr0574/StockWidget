"""CN index 11-tier market breadth refresh lifecycle, UI toggle, and rendering."""

from dataclasses import replace
from datetime import datetime, timedelta, timezone
import tempfile
from unittest.mock import Mock, patch

from PySide6.QtGui import QColor

from stockwidget.core.quote_presentation import MARKET_BREADTH_BUCKETS
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import Bar, BarResult
from stockwidget.data.quotes import MARKET_BREADTH_KEY, _new_entry, parse_sina_market_breadth
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

        # Open CN index with breadth disabled -> label says "涨跌分布", no extra quote fetch.
        window.quotes.refresh.reset_mock()
        history.open_row("float", 0)  # sh000001
        self.wait_ready(window)
        popup = history.dialog
        self.assertEqual(popup.ui.cb_order_book.text(), "涨跌分布")
        self.assertFalse(history.needs_market_breadth())
        self.assertTrue(popup.order_book.isHidden())

        # Enable "涨跌分布" -> still no immediate out-of-cadence quote fetch; shows placeholder "—".
        popup.ui.cb_order_book.setChecked(True)
        self.app.processEvents()
        window.quotes.refresh.assert_not_called()

        self.assertTrue(history.needs_market_breadth())
        self.assertFalse(popup.order_book.isHidden())
        self.assertTrue(popup.order_book.breadth_mode)
        self.assertIsNone(popup.order_book.breadth_summary)
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
        popup.ui.cb_order_book.setChecked(False)
        self.app.processEvents()
        self.assertFalse(history.needs_market_breadth())
        self.assertIsNone(window.quotes.market_breadth())

        # Closing discards breadth, and a late in-flight result with breadth must not repopulate it.
        popup.ui.cb_order_book.setChecked(True)
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

        # Even when unchecked at minimum size, "涨跌分布" checkbox fits inside popup.
        popup.resize(1, 1)
        self.app.processEvents()
        for control in (*popup.view_buttons.values(), popup.ui.cb_volume, popup.ui.cb_average, popup.ui.cb_order_book):
            self.assertTrue(popup.rect().contains(control.mapTo(popup, control.rect().bottomRight())))

        # Enable breadth and verify minimum layout fits all 12 rows without clipping.
        popup.ui.cb_order_book.setChecked(True)
        self.app.processEvents()
        popup.resize(1, 1)
        self.app.processEvents()
        book = popup.order_book
        self.assertTrue(popup.rect().contains(book.geometry()))
        self.assertLess(book.geometry().bottom(), popup.ui.chart_readout.geometry().top())
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
