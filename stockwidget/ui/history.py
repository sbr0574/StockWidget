"""Shared history window/popup and bounded asynchronous request lifecycle."""

import threading

from PySide6.QtCore import QObject, QPoint, QRect, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QDialog, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QToolTip, QVBoxLayout, QWidget,
)

from stockwidget.core.window_rules import adjacent_popup_position, adjacent_popup_size, best_screen
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import BarResult, HistorySeries, history_day, select_bars
from stockwidget.ui.controls.history_chart import HistoryChart
from stockwidget.ui.controls.style import apply_settings_theme, is_dark_theme, theme_palette


class HistoryDialog(QDialog):
    view_changed = Signal()
    dismissed = Signal()
    SIZES = {"window": (760, 490), "large": (560, 400), "medium": (420, 300), "small": (320, 230)}

    def __init__(self, parent, display_mode="window"):
        super().__init__(parent, Qt.Dialog)
        self.setModal(False)
        layout = QVBoxLayout(self)
        self.title_row = QWidget()
        title_layout = QHBoxLayout(self.title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        self.title = QLabel()
        self.title.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self._title_text = self._status_text = ""
        title_layout.addWidget(self.title, 1)
        self.close_button = QPushButton("×")
        self.close_button.setFixedSize(26, 26)
        self.close_button.setAutoDefault(False)
        self.close_button.setToolTip("关闭图表（Esc）")
        self.close_button.setAccessibleName("关闭图表")
        self.close_button.clicked.connect(self.close)
        title_layout.addWidget(self.close_button)
        layout.addWidget(self.title_row)
        controls = QHBoxLayout()
        self.view_buttons = {}
        self.view_group = QButtonGroup(self)
        self._view = "intraday"
        for label, value in (("分时", "intraday"), ("5日", "five_day"), ("日K", "daily")):
            button = QPushButton(label)
            button.setCheckable(True)
            button.setAutoDefault(False)
            button.setMinimumSize(60, 28)
            button.setAccessibleName(label)
            button.setToolTip("最近30个交易日的日K线" if value == "daily" else "当日分时" if value == "intraday" else "最近5个交易日分时")
            self.view_buttons[value] = button
            self.view_group.addButton(button)
            controls.addWidget(button, 1)
        self.view_buttons[self._view].setChecked(True)
        self.view_group.buttonToggled.connect(self._view_changed)
        layout.addLayout(controls)
        self.chart = HistoryChart()
        layout.addWidget(self.chart, 1)
        self.status = QLabel()
        self.status.setProperty("settingDescription", True)
        self.status.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status.hide()
        layout.addWidget(self.status)
        self.display_mode = None
        self.set_display_mode(display_mode)

    def set_display_mode(self, mode):
        if mode == self.display_mode:
            return
        floating = mode != "window"
        if self.display_mode is not None and floating != (self.display_mode != "window"):
            # Reusing a Windows handle across Dialog/Popup types lets queued
            # native close/resize events dismiss or resize the new popup.
            self.hide()
            self.destroy()
        self.display_mode = mode
        # Popup capture also handles outside clicks in other applications.
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint if floating else Qt.Dialog)
        # A dismissal click on a quote row must not replay and reopen the chart.
        self.setAttribute(Qt.WA_NoMouseReplay, floating)
        self.title_row.setVisible(floating)
        chart_minimum, minimum = {"window": ((400, 280), (530, 390)), "large": ((360, 160), (400, 300)),
                                  "medium": ((280, 150), (330, 250)), "small": ((220, 110), (260, 210))}[mode]
        self.chart.setMinimumSize(*chart_minimum)
        self.setMinimumSize(*minimum)
        margin = 6 if floating else 10
        self.layout().setContentsMargins(margin, margin, margin, margin)
        self.layout().setSpacing(4 if floating else 6)
        self.resize(*self.SIZES[mode])
        self.setWindowOpacity(.92 if floating else 1.)

    def hideEvent(self, event):
        super().hideEvent(event)
        QToolTip.hideText()
        self.dismissed.emit()

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
        self.title.setToolTip(text)
        self._elide_labels()

    def set_status(self, text):
        self._status_text = text
        self.status.setVisible(bool(text))
        self.status.setToolTip(text)
        self._elide_labels()

    def _elide_labels(self):
        for label, text in ((self.title, self._title_text), (self.status, self._status_text)):
            label.setText(label.fontMetrics().elidedText(text, Qt.ElideRight, max(1, label.width())))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._elide_labels()


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
        self.cache = BarCache()
        self._generation = 0
        self._discard_before = 0
        self._busy = False
        self._pending = None
        self._displayed_key = None
        self._series = {}
        self.result_ready.connect(self._accept)
        self.window.view_options_changed.connect(self._options_changed)
        self.window.display_flags_changed.connect(self._apply_chart_options)
        self.window.quotes.quotes_updated.connect(self._quotes_updated)
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
            if self.dialog:
                self.dialog.close()

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
        if self.dialog is None:
            self.dialog = HistoryDialog(self.window, self.window.view_options.chart_display_mode)
            self.dialog.view_changed.connect(self.reload)
            self.dialog.dismissed.connect(self._closed)
        self.instrument = dict(instrument)
        self.dialog.set_title(f"{instrument.get('name') or instrument.get('code')} · 行情图表")
        self._apply_chart_options()
        self._apply_theme()
        self._show()
        self.reload()

    def _show(self):
        self._place_floating()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()

    def _place_floating(self):
        self.dialog.layout().activate()
        if self.dialog.display_mode != "window":
            anchor = (self.window.position_controller.full_geometry() if self._surface == "float"
                      else QRect(self._taskbar_position, QSize(1, 1)))
            bounds = best_screen(*anchor.getRect(), [screen.availableGeometry().getRect()
                                                    for screen in QApplication.screens()])
            if bounds:
                minimum = self.dialog.minimumSize().expandedTo(self.dialog.minimumSizeHint())
                self.dialog.resize(*adjacent_popup_size(anchor.getRect(), bounds, self.dialog.SIZES[self.dialog.display_mode],
                                                       (minimum.width(), minimum.height())))
                self.dialog.move(*adjacent_popup_position(anchor.getRect(), bounds,
                                                         self.dialog.width(), self.dialog.height(),
                                                         prefer_above=self._surface == "taskbar"))

    def _apply_theme(self):
        if self.dialog is None:
            return
        mode = self.window.view_options.color_mode
        dark = mode == "dark" or (mode == "system" and is_dark_theme())
        apply_settings_theme(self.dialog, dark, theme_palette(dark, self.window.palette()))
        self.dialog.chart.up_color = self.window.up_color
        self.dialog.chart.down_color = self.window.down_color

    def _options_changed(self):
        if not self.window.view_options.chart_enabled:
            self.click_timer.stop()
            if self.dialog:
                self.dialog.close()
        else:
            if self.dialog and self.dialog.display_mode != self.window.view_options.chart_display_mode:
                visible = self.dialog.isVisible()
                self.dialog.set_display_mode(self.window.view_options.chart_display_mode)
                if visible:
                    self._show()
                    self.reload()
            self._apply_theme()
            self._apply_chart_options()

    def _apply_chart_options(self):
        if self.dialog:
            options = self.window.view_options
            self.dialog.chart.set_options(options.chart_ma_periods, options.chart_average_enabled,
                                          options.chart_volume_enabled, self.window.unit_mode)

    def _closed(self, *_args):
        self._generation += 1
        self._pending = None

    def clear_cache(self):
        cleared = self.cache.clear()
        self.click_timer.stop()
        self._clicked_instrument = None
        self._closed()
        self._discard_before = self._generation
        self._series.clear()
        self._displayed_key = None
        if self.dialog:
            self.dialog.close()
            self.dialog.chart.set_data((), self.dialog.view, self.instrument or {})
        return cleared

    @staticmethod
    def _key(instrument, view, source):
        return (*(instrument.get(field) for field in ("market", "code", "type")), source,
                "daily" if view == "daily" else "five_day")

    def _series_for(self, instrument, view, source):
        key = self._key(instrument, view, source)
        if key not in self._series:
            self._series[key] = HistorySeries(instrument, view)
        series = self._series[key]
        quote = self.window.quotes.quote_for(instrument) if source == self.window.data_source else None
        if quote:
            series.update_quote(quote)
        return series

    def _quotes_updated(self):
        for key, series in self._series.items():
            if key[-2] == self.window.data_source:
                quote = self.window.quotes.quote_for(series.instrument)
                if quote:
                    series.update_quote(quote)
        if self.dialog and self.dialog.isVisible() and self.instrument:
            series = self._series_for(self.instrument, self.dialog.view, self.window.data_source)
            if series.loaded_day is not None:
                self._render(series, self.dialog.view, self.window.data_source)
            if not self._busy and series.needs_download(self.cache.clock()):
                self.reload()

    def reload(self):
        if not self.dialog or not self.dialog.isVisible() or not self.instrument or not self.window.view_options.chart_enabled:
            return
        self._generation += 1
        view, source = self.dialog.view, self.window.data_source
        series = self._series_for(self.instrument, view, source)
        now = self.cache.clock()
        day = history_day(self.instrument, now)
        request = (self._generation, dict(self.instrument), view, source, day,
                   series.loaded_day == day and series.repair_needed)
        key = (self._key(self.instrument, view, source), view)
        if key != self._displayed_key:
            self.dialog.chart.set_data((), self.dialog.view, self.instrument)
        if not series.needs_download(now):
            self._pending = None
            self._render(series, view, source)
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
        generation, instrument, view, source, day, repair = request
        try:
            result = self.cache.get(instrument, "daily" if view == "daily" else "five_day", source,
                                    repair=repair, generation=cache_generation)
        except Exception:
            result = BarResult(message="暂无可用历史数据 · 历史数据处理失败")
        try:
            self.result_ready.emit((*request, result))
        except RuntimeError:
            pass  # Parent was destroyed while the network request was running.

    def _accept(self, payload):
        self._busy = False
        generation, instrument, view, preferred, day, repair, result = payload
        if generation >= self._discard_before:
            series = self._series_for(instrument, view, preferred)
            series.set_history(result, day, self.cache.clock())
            if generation == self._generation and self.dialog and self.dialog.isVisible():
                self._render(series, view, preferred)
        if self._pending is not None:
            pending, self._pending = self._pending, None
            if pending[0] == self._generation and self.dialog.isVisible():
                # The completed worker may have loaded the same minute family
                # while the user switched from intraday to five-day view.
                self.reload()

    def _render(self, series, view, preferred):
        result = series.result
        instrument = series.instrument
        bars = select_bars(result.bars, view, instrument)
        self.dialog.chart.set_data(bars, view, instrument, reference_price=series.reference_price, now=self.cache.clock())
        self._displayed_key = (self._key(instrument, view, preferred), view)
        pieces = [result.message]
        if result.stale:
            pieces.append("缓存（更新失败）")
        if bars and view == "daily" and len(bars) < 89:
            pieces.append("历史不足时，部分均线从满足周期处开始显示")
        self.dialog.set_status(" · ".join(piece for piece in pieces if piece))
        self._place_floating()
