"""Shared history window/popup and bounded asynchronous request lifecycle."""

import threading

from PySide6.QtCore import QObject, QPoint, QRect, QSize, Qt, QTimer, QSignalBlocker, Signal
from PySide6.QtGui import QColor, QCursor, QPalette
from PySide6.QtWidgets import QApplication, QButtonGroup, QDialog, QToolTip

from stockwidget.core.window_rules import adjacent_popup_position, best_screen
from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import MA_PERIODS, BarResult, HistorySeries, history_day, merge_quote, mini_trend_points, quote_moving_average, quote_volume_ratio, select_bars
from stockwidget.ui.controls.history_chart import MA_COLORS
from stockwidget.ui.generated.ui_history import Ui_HistoryDialog
from stockwidget.ui.controls.style import apply_settings_theme, is_dark_theme, theme_palette


class HistoryDialog(QDialog):
    view_changed = Signal()
    dismissed = Signal()
    pop_out = Signal()
    size_changed = Signal(str, object)
    options_changed = Signal(object)

    def __init__(self, parent, display_mode="popup"):
        floating = display_mode == "popup"
        flags = (Qt.Popup | Qt.FramelessWindowHint if floating else
                 Qt.Window | Qt.WindowTitleHint | Qt.WindowSystemMenuHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        super().__init__(parent if floating else None, flags)
        self.display_mode = display_mode
        self._restoring = False
        self._saved_size = None
        self._gesture = None
        self._book_enabled = False
        self._theme_palette = None
        self.setMouseTracking(True)
        self.setAttribute(Qt.WA_NoMouseReplay, floating)
        self.setModal(False)
        self.ui = Ui_HistoryDialog()
        self.ui.setupUi(self)
        for name in ("title_row", "title", "pop_out_button", "close_button", "quote_header", "chart", "order_book", "status"):
            setattr(self, name, getattr(self.ui, name))
        self.quote_header.bind_fields()
        self._title_text = self._status_text = ""
        self.title_row.setVisible(floating)
        self.quote_header.setVisible(not floating)
        self.order_book.hide()
        self.ui.ma_row.hide()
        self.status.hide()
        self.view_group = QButtonGroup(self)
        self._view = "intraday"
        self.view_buttons = {view: getattr(self.ui, f"btn_{view}") for view in ("intraday", "five_day", "daily")}
        for button in self.view_buttons.values():
            self.view_group.addButton(button)
            button.setFixedWidth(button.fontMetrics().horizontalAdvance(button.text()) + 12)
        self.view_group.buttonToggled.connect(self._view_changed)
        self.pop_out_button.clicked.connect(self.pop_out)
        self.close_button.clicked.connect(self.close)
        for control, key in ((self.ui.cb_order_book, "chart_order_book_enabled"),
                             (self.ui.cb_volume, "chart_volume_enabled"),
                             (self.ui.cb_average, "chart_average_enabled")):
            control.toggled.connect(lambda checked, key=key: self.options_changed.emit({key: checked}))
        for period in MA_PERIODS:
            getattr(self.ui, f"cb_ma{period}").toggled.connect(self._ma_changed)
        self.setWindowOpacity(.92 if floating else 1.)
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(200)
        self._save_timer.timeout.connect(self.save_size)

    def set_chart_options(self, options, unit_mode):
        self.chart.set_options(options.chart_ma_periods, options.chart_average_enabled,
                               options.chart_volume_enabled, unit_mode)
        for control, checked in ((self.ui.cb_order_book, options.chart_order_book_enabled),
                                 (self.ui.cb_volume, options.chart_volume_enabled),
                                 (self.ui.cb_average, options.chart_average_enabled)):
            with QSignalBlocker(control):
                control.setChecked(checked)
        for period, color in zip(MA_PERIODS, MA_COLORS):
            control = getattr(self.ui, f"cb_ma{period}")
            with QSignalBlocker(control):
                control.setChecked(period in options.chart_ma_periods)
            palette = QPalette(self.palette())
            palette.setColor(QPalette.WindowText, QColor(color))
            control.setPalette(palette)
        self.ui.ma_row.setVisible(self.view == "daily" and options.chart_average_enabled)
        self.ui.cb_order_book.setEnabled(self.view == "intraday")
        self.set_order_book(options.chart_order_book_enabled)

    def _ma_changed(self):
        self.options_changed.emit({"chart_ma_periods": [period for period in MA_PERIODS
                                                       if getattr(self.ui, f"cb_ma{period}").isChecked()]})

    def set_order_book(self, enabled):
        visible = bool(enabled and self.view == "intraday")
        if self._book_enabled != visible:
            self._book_enabled = visible
            self._restoring = True
            self.order_book.setVisible(visible)
            self.update_size_limits()
            self.layout().activate()
            self._restoring = False

    def update_size_limits(self, bounds=None):
        screen = QApplication.screenAt(self.frameGeometry().center()) or QApplication.primaryScreen()
        bounds = bounds or screen.availableGeometry()
        # Reserve native title-bar/frame space for the detached window.
        frame = self.frameGeometry().size() - self.size() if self.isVisible() else QSize(16, 40)
        maximum = bounds.size() if self.display_mode == "popup" else bounds.size() - frame
        minimum = QSize(680, 400) if self.display_mode == "window" else QSize(300, 210)
        if self._book_enabled:
            minimum = minimum.expandedTo(QSize(460, 300))
        maximum = maximum.expandedTo(QSize(200, 200))
        self.setMinimumSize(minimum.boundedTo(maximum))
        self.setMaximumSize(maximum)

    def restore_size(self, size, bounds):
        self._restoring = True
        self.update_size_limits(bounds)
        self.resize(QSize(*size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
        self.layout().activate()
        self._restoring = False

    def save_size(self):
        self._save_timer.stop()
        if self._saved_size is not None:
            size, self._saved_size = self._saved_size, None
            self.size_changed.emit(self.display_mode, size)

    def hideEvent(self, event):
        self._gesture = None
        self.save_size()
        super().hideEvent(event)
        QToolTip.hideText()
        if not (self.display_mode == "window" and event.spontaneous() and self.isMinimized()):
            self.dismissed.emit()

    def showEvent(self, event):
        super().showEvent(event)
        self.layout().activate()
        self._elide_labels()

    @property
    def view(self):
        return self._view

    def set_view(self, view):
        self.view_buttons[view].setChecked(True)

    def _view_changed(self, button, checked):
        if not checked:
            return
        self._view = next(view for view, candidate in self.view_buttons.items() if candidate is button)
        self.view_changed.emit()

    def set_title(self, text):
        self._title_text = text
        self.setWindowTitle(text)
        self._elide_labels()

    def set_status(self, text):
        self._status_text = text
        self.status.setVisible(bool(text))
        self._elide_labels()

    def _elide_labels(self):
        for label, text in ((self.title, self._title_text), (self.status, self._status_text)):
            label.setText(label.fontMetrics().elidedText(text, Qt.ElideRight, max(1, label.width())))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.layout().activate()
        self._elide_labels()
        if self.isVisible() and not self._restoring and not self.isMaximized() and not self.isMinimized():
            screen = QApplication.screenAt(self.frameGeometry().center()) or QApplication.primaryScreen()
            bounds, frame = screen.availableGeometry(), self.frameGeometry()
            x = max(bounds.left(), min(bounds.right() - frame.width() + 1, frame.left()))
            y = max(bounds.top(), min(bounds.bottom() - frame.height() + 1, frame.top()))
            if (x, y) != (frame.x(), frame.y()):
                self.move(self.pos() + QPoint(x - frame.x(), y - frame.y()))
            self._saved_size = [self.width(), self.height()]
            self._save_timer.start()

    def moveEvent(self, event):
        super().moveEvent(event)
        if self.isVisible() and not self._restoring:
            self._restoring = True
            self.update_size_limits()
            self._restoring = False

    def _edges(self, position):
        if self.display_mode != "popup" or not self.rect().contains(position):
            return Qt.Edges()
        edges = Qt.Edges()
        for hit, edge in ((position.x() < 7, Qt.LeftEdge), (position.x() >= self.width() - 7, Qt.RightEdge),
                          (position.y() < 7, Qt.TopEdge), (position.y() >= self.height() - 7, Qt.BottomEdge)):
            if hit:
                edges |= edge
        return edges

    def mousePressEvent(self, event):
        edges = self._edges(event.position().toPoint())
        if event.button() == Qt.LeftButton and edges:
            self._gesture = edges, event.globalPosition().toPoint(), self.geometry()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._gesture:
            edges, start, original = self._gesture
            delta = event.globalPosition().toPoint() - start
            width = original.width() + (delta.x() if edges & Qt.RightEdge else -delta.x() if edges & Qt.LeftEdge else 0)
            height = original.height() + (delta.y() if edges & Qt.BottomEdge else -delta.y() if edges & Qt.TopEdge else 0)
            width = max(self.minimumWidth(), min(self.maximumWidth(), width))
            height = max(self.minimumHeight(), min(self.maximumHeight(), height))
            x = original.right() - width + 1 if edges & Qt.LeftEdge else original.x()
            y = original.bottom() - height + 1 if edges & Qt.TopEdge else original.y()
            screen = QApplication.screenAt(original.center()) or QApplication.primaryScreen()
            bounds = screen.availableGeometry()
            self.setGeometry(max(bounds.left(), min(bounds.right() - width + 1, x)),
                             max(bounds.top(), min(bounds.bottom() - height + 1, y)), width, height)
            event.accept()
            return
        edges = self._edges(event.position().toPoint())
        horizontal, vertical = bool(edges & (Qt.LeftEdge | Qt.RightEdge)), bool(edges & (Qt.TopEdge | Qt.BottomEdge))
        if horizontal and vertical:
            cursor = Qt.SizeFDiagCursor if bool(edges & Qt.LeftEdge) == bool(edges & Qt.TopEdge) else Qt.SizeBDiagCursor
        else:
            cursor = Qt.SizeHorCursor if horizontal else Qt.SizeVerCursor if vertical else Qt.ArrowCursor
        self.setCursor(cursor)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._gesture and event.button() == Qt.LeftButton:
            self._gesture = None
            self.save_size()
            event.accept()
        else:
            super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if self._gesture and event.key() == Qt.Key_Escape:
            original = self._gesture[2]
            self._gesture = None
            self.setGeometry(original)
            self.save_size()
            event.accept()
        else:
            super().keyPressEvent(event)


class HistoryController(QObject):
    result_ready = Signal(object)

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        for block, table in enumerate(window.float_tables):
            window.register_drag_region(table.viewport(),
                                        lambda pos, table=table, block=block: self.request_row(
                                            "float", table.indexAt(pos).row(), block))
        self.dialog = None
        self.instrument = None
        self.cache = BarCache(history_cache_dir())
        self._generation = 0
        self._discard_before = 0
        self._busy = False
        self._pending = None
        self._displayed_key = None
        self._series = {}
        self._history_pending = {}
        self._history_scheduled = False
        self.result_ready.connect(self._accept)
        self.window.view_options_changed.connect(self._options_changed)
        self.window.display_flags_changed.connect(self._apply_chart_options)
        self.window.quotes.quotes_updated.connect(self._quotes_updated)
        self.window.presentation_changed.connect(self.schedule_history_refresh)
        self.click_timer = QTimer(self)
        self.click_timer.setSingleShot(True)
        self.click_timer.setTimerType(Qt.PreciseTimer)
        self.click_timer.timeout.connect(self._open_pending)
        self._clicked_instrument = None
        self._surface = "float"
        self._taskbar_position = QPoint()
        self.window.widget_visibility_changed.connect(self._visibility_changed)

    def request_row(self, surface, row, block=0, *, global_pos=None):
        if not self.window.view_options.chart_enabled:
            return
        self._clicked_instrument = self.window.quotes.instrument_at(surface, row, block)
        if self._clicked_instrument:
            self._set_anchor(surface, global_pos)
            # Preserve the existing double-click-to-hide gesture without a popup.
            self.click_timer.start(QApplication.doubleClickInterval())

    def _open_pending(self):
        if self._clicked_instrument and self.window.view_options.chart_enabled and self.window.widget_visible:
            self._open(self._clicked_instrument)

    def _visibility_changed(self):
        if not self.window.widget_visible:
            self.click_timer.stop()
            self._closed()
            self._discard_before = self._generation
            self._history_pending.clear()
            if self.dialog:
                self.dialog.close()
        else:
            self.schedule_history_refresh()

    def open_row(self, surface, row, block=0, *, global_pos=None):
        if not self.window.view_options.chart_enabled:
            return
        instrument = self.window.quotes.instrument_at(surface, row, block)
        if instrument is None:
            return
        self._set_anchor(surface, global_pos)
        self._open(instrument)

    def _set_anchor(self, surface, global_pos):
        self._surface = surface
        if surface == "taskbar":
            self._taskbar_position = QPoint(global_pos if global_pos is not None else QCursor.pos())

    def _open(self, instrument):
        if self.dialog is None or (self.dialog.display_mode == "window" and not self.dialog.isVisible()):
            previous = self.dialog
            self.dialog = self._create_dialog("popup")
            if previous:
                previous.deleteLater()
        self.instrument = dict(instrument)
        self.dialog.set_title(f"{instrument.get('name') or instrument.get('code')} · 行情图表")
        self._apply_chart_options()
        self._apply_theme()
        self._show()
        self.reload()

    def _create_dialog(self, mode):
        dialog = HistoryDialog(self.window, mode)
        self.window.destroyed.connect(dialog.deleteLater)
        dialog.view_changed.connect(self._view_changed)
        dialog.dismissed.connect(self._closed)
        dialog.pop_out.connect(self._pop_out)
        dialog.size_changed.connect(self._save_size)
        dialog.options_changed.connect(lambda changes: self.window.set_view_options(**changes))
        return dialog

    def _save_size(self, mode, size):
        self.window.set_view_options(**{f"chart_{mode}_size": size})

    def _view_changed(self):
        self._apply_chart_options()
        self.reload()

    def _pop_out(self):
        if not self.dialog or self.dialog.display_mode != "popup":
            return
        previous = self.dialog
        previous.save_size()
        # Separate native handles prevent queued Popup close/resize events from
        # affecting the detached window. Keep request generation and series.
        previous.dismissed.disconnect(self._closed)
        self.dialog = self._create_dialog("window")
        self.dialog.set_view(previous.view)
        self.dialog.set_title(previous.windowTitle())
        self.dialog.set_status(previous._status_text)
        previous.close()
        previous.deleteLater()
        self._apply_chart_options()
        self._apply_theme()
        self._show()
        series = self._series_for(self.instrument, self.dialog.view)
        if series.loaded_day is not None:
            self._render(series, self.dialog.view)
        self._start_next_background()

    def _show(self):
        self._place_floating(restore=not self.dialog.isVisible())
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()

    def _place_floating(self, *, restore=False):
        self.dialog.layout().activate()
        anchor = (self.window.position_controller.full_geometry() if self._surface == "float"
                  else QRect(self._taskbar_position, QSize(1, 1)))
        bounds = best_screen(*anchor.getRect(), [screen.availableGeometry().getRect()
                                                for screen in QApplication.screens()])
        if bounds:
            if restore:
                size = getattr(self.window.view_options, f"chart_{self.dialog.display_mode}_size")
                self.dialog.restore_size(size, QRect(*bounds))
            if self.dialog.display_mode == "popup":
                self.dialog.move(*adjacent_popup_position(anchor.getRect(), bounds,
                                                         self.dialog.width(), self.dialog.height(),
                                                         prefer_above=self._surface == "taskbar"))
            elif restore:
                rect = QRect(QPoint(), self.dialog.size())
                rect.moveCenter(QRect(*bounds).center())
                self.dialog.move(rect.topLeft())

    def _apply_theme(self):
        if self.dialog is None:
            return
        mode = self.window.view_options.color_mode
        dark = mode == "dark" or (mode == "system" and is_dark_theme())
        palette = theme_palette(dark, self.window.palette())
        if self.dialog._theme_palette != palette:
            apply_settings_theme(self.dialog, dark, palette)
            self.dialog._theme_palette = palette
        self.dialog.chart.up_color = self.window.up_color
        self.dialog.chart.down_color = self.window.down_color
        self._update_details()

    def _options_changed(self):
        self.schedule_history_refresh()
        if not self.window.view_options.chart_enabled:
            self.click_timer.stop()
            if self.dialog:
                self.dialog.close()
        else:
            self._apply_theme()
            self._apply_chart_options()

    def _apply_chart_options(self):
        if self.dialog:
            options = self.window.view_options
            self.dialog.set_chart_options(options, self.window.unit_mode)
            self._update_details()

    def _update_details(self):
        if not self.dialog or not self.instrument:
            return
        quote = self.window.quotes.quote_for(self.instrument)
        self.dialog.order_book.set_quote(self.instrument, quote, self.window.unit_mode,
                                         self.window.up_color, self.window.down_color)
        if self.dialog.display_mode == "window":
            daily = self._series_for(self.instrument, "daily")
            ratio = None if daily.result.stale else quote_volume_ratio(daily.result.bars, quote, self.instrument)
            self.dialog.quote_header.set_quote(self.instrument, quote, self.window.unit_mode,
                                               self.window.up_color, self.window.down_color, ratio)

    def _closed(self, *_args):
        self._generation += 1
        self._pending = None

    def invalidate_quotes(self):
        """Watchlist changes invalidate both consumers of history."""
        self._closed()
        self._discard_before = self._generation
        self._history_pending.clear()
        self._series.clear()
        self._displayed_key = None
        if self.window.view_options.mini_chart_type == "trend":
            self.window.quotes.reproject()

    def clear_cache(self):
        cleared = self.cache.clear()
        self._discard_cached_history()
        return cleared

    def _discard_cached_history(self):
        self.click_timer.stop()
        self._clicked_instrument = None
        self._closed()
        self._discard_before = self._generation
        self._series.clear()
        self._history_pending.clear()
        self._displayed_key = None
        if self.dialog:
            self.dialog.close()
            self.dialog.chart.set_data((), self.dialog.view, self.instrument or {})

    @staticmethod
    def _key(instrument, view):
        return (*(instrument.get(field) for field in ("market", "code", "type")),
                "daily" if view == "daily" else "five_day")

    def _series_for(self, instrument, view):
        key = self._key(instrument, view)
        if key not in self._series:
            self._series[key] = HistorySeries(instrument, view)
        series = self._series[key]
        quote = self.window.quotes.quote_for(instrument)
        if quote:
            series.update_quote(quote)
        return series

    def _quotes_updated(self):
        for series in self._series.values():
            quote = self.window.quotes.quote_for(series.instrument)
            if quote:
                series.update_quote(quote)
        if self.dialog and self.dialog.isVisible() and self.instrument:
            self._update_details()
            series = self._series_for(self.instrument, self.dialog.view)
            if series.loaded_day is not None:
                self._render(series, self.dialog.view)
            if not self._busy and series.needs_download(self.cache.clock()):
                self.reload()
            self._start_next_background()
        self.window.quotes.reproject()
        self.schedule_history_refresh()

    def quote_metrics(self, instrument, quote):
        series = self._series.get(self._key(instrument, "daily"))
        if series is None or series.result.stale:
            return {}
        return {"volume_ratio": quote_volume_ratio(series.result.bars, quote, instrument),
                "average_price": quote_moving_average(series.result.bars, quote, instrument,
                                                      self.window.view_options.average_period)}

    def mini_chart_data(self, instrument):
        """Use the popup's minute series and live quote at the same cadence."""
        series = self._series_for(instrument, "intraday")
        bars = series.result.bars
        quote = self.window.quotes.quote_for(instrument)
        if quote:
            bars = merge_quote(bars, instrument, "five_day", quote)
        reference = series.reference_price_for(bars)
        bars = select_bars(bars, "intraday", instrument)
        return mini_trend_points(bars, instrument, self.cache.clock()), reference

    def schedule_history_refresh(self):
        if not self._history_scheduled:
            self._history_scheduled = True
            QTimer.singleShot(0, self, self._refresh_background)

    def _refresh_background(self):
        self._history_scheduled = False
        self._history_pending = {self._key(instrument, view): (instrument, view)
                              for view in ("intraday", "daily")
                              for instrument in self.window.quotes.history_instruments(view).values()}
        self._start_next_background()

    def _start_next_background(self):
        if self._busy:
            return
        if (self.dialog and self.dialog.isVisible() and self.dialog.display_mode == "window"
                and self.instrument and self.instrument.get("market") in {"sh", "sz", "bj"}
                and self.instrument.get("type") not in {"期", "指"}):
            now = self.cache.clock()
            series = self._series_for(self.instrument, "daily")
            if series.needs_download(now):
                self._start((self._generation, dict(self.instrument), "daily", history_day(self.instrument, now), False))
                return
        wanted = {self._key(entry, view) for view in ("intraday", "daily")
                  for entry in self.window.quotes.history_instruments(view).values()}
        while self._history_pending:
            key = next(iter(self._history_pending))
            instrument, view = self._history_pending.pop(key)
            if key not in wanted:
                continue
            now = self.cache.clock()
            series = self._series_for(instrument, view)
            if series.needs_download(now):
                self._start((self._generation, dict(instrument), view,
                             history_day(instrument, now), series.loaded_day == history_day(instrument, now)
                             and series.repair_needed))
                return

    def reload(self):
        if not self.dialog or not self.dialog.isVisible() or not self.instrument or not self.window.view_options.chart_enabled:
            return
        self._generation += 1
        view = self.dialog.view
        series = self._series_for(self.instrument, view)
        now = self.cache.clock()
        day = history_day(self.instrument, now)
        request = (self._generation, dict(self.instrument), view, day,
                   series.loaded_day == day and series.repair_needed)
        key = (self._key(self.instrument, view), view)
        if key != self._displayed_key:
            self.dialog.chart.set_data((), self.dialog.view, self.instrument)
        if not series.needs_download(now):
            self._pending = None
            self._render(series, view)
            self._start_next_background()
            return
        self.dialog.set_status("更新中…" if self.dialog.chart.bars else "加载中…")
        if self._busy:
            self._pending = request  # At most one worker; latest selection wins.
        else:
            self._start(request)

    def _start(self, request):
        self._busy = True
        self._pending = None
        threading.Thread(target=self._worker, args=(request, self.cache.generation), daemon=True).start()

    def _worker(self, request, cache_generation):
        generation, instrument, view, day, repair = request
        try:
            result = self.cache.get(instrument, "daily" if view == "daily" else "five_day",
                                    repair=repair, generation=cache_generation)
        except Exception:
            result = BarResult(message="暂无可用历史数据 · 历史数据处理失败")
        try:
            self.result_ready.emit((*request, result))
        except RuntimeError:
            pass  # Parent was destroyed while the network request was running.

    def _accept(self, payload):
        self._busy = False
        generation, instrument, view, day, repair, result = payload
        if generation >= self._discard_before and self.window.widget_visible:
            series = self._series_for(instrument, view)
            series.set_history(result, day, self.cache.clock())
            if self.dialog and self.dialog.isVisible():
                self._update_details()
            if (generation == self._generation and self.dialog and self.dialog.isVisible() and self.instrument
                    and self._key(instrument, view) == self._key(self.instrument, self.dialog.view)):
                self._render(series, self.dialog.view)
            self.window.quotes.reproject()
        if self._pending is not None:
            pending, self._pending = self._pending, None
            if pending[0] == self._generation and self.dialog and self.dialog.isVisible():
                # The completed worker may have loaded the same minute family
                # while the user switched from intraday to five-day view.
                self.reload()
        self._start_next_background()

    def _render(self, series, view):
        result = series.result
        instrument = series.instrument
        bars = select_bars(result.bars, view, instrument)
        self.dialog.chart.set_data(bars, view, instrument, reference_price=series.reference_price, now=self.cache.clock())
        self._displayed_key = (self._key(instrument, view), view)
        pieces = [result.message]
        if result.stale:
            pieces.append("缓存（更新失败）")
        if bars and view == "daily" and len(bars) < 89:
            pieces.append("历史不足时，部分均线从满足周期处开始显示")
        self.dialog.set_status(" · ".join(piece for piece in pieces if piece))
