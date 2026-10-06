"""Nonmodal history dialog and bounded asynchronous request lifecycle."""

import threading

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtWidgets import QApplication, QCheckBox, QComboBox, QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import BarResult, MA_PERIODS
from stockwidget.ui.controls.history_chart import HistoryChart, MA_COLORS
from stockwidget.ui.controls.style import build_settings_stylesheet, is_dark_theme, theme_palette


class HistoryDialog(QDialog):
    view_changed = Signal()
    refresh_requested = Signal()

    def __init__(self, parent):
        super().__init__(parent, Qt.Dialog)
        self.resize(760, 490)
        self.setMinimumSize(530, 390)
        self.setModal(False)
        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        self.view_combo = QComboBox()
        for label, value in (("当日分时", "intraday"), ("5日分时", "five_day"), ("日K线 · 最近30个交易日", "daily")):
            self.view_combo.addItem(label, value)
        controls.addWidget(self.view_combo)
        controls.addStretch()
        self.refresh_button = QPushButton("刷新")
        self.refresh_button.setToolTip("按缓存更新周期检查历史数据")
        controls.addWidget(self.refresh_button)
        layout.addLayout(controls)
        ma_layout = QHBoxLayout()
        self.ma_boxes = []
        for period, color in zip(MA_PERIODS, MA_COLORS):
            box = QCheckBox(f"MA{period}")
            box.setChecked(True)
            box.setStyleSheet(f"QCheckBox {{ color: {color}; }} QCheckBox:disabled {{ color: palette(mid); }}")
            box.toggled.connect(self._sync_periods)
            self.ma_boxes.append(box)
            ma_layout.addWidget(box)
        ma_layout.addStretch()
        layout.addLayout(ma_layout)
        self.chart = HistoryChart()
        layout.addWidget(self.chart, 1)
        self.status = QLabel()
        self.status.setWordWrap(True)
        self.status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.status)
        self.view_combo.currentIndexChanged.connect(self._view_changed)
        self.refresh_button.clicked.connect(self.refresh_requested)
        self._view_changed()

    @property
    def view(self):
        return self.view_combo.currentData()

    def _view_changed(self):
        for box in self.ma_boxes:
            box.setEnabled(self.view == "daily")
        self.view_changed.emit()

    def _sync_periods(self):
        self.chart.set_periods([period for period, box in zip(MA_PERIODS, self.ma_boxes) if box.isChecked()])


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
        self._busy = False
        self._pending = None
        self._displayed_key = None
        self.result_ready.connect(self._accept)
        self.window.view_options_changed.connect(self._options_changed)
        self.timer = QTimer(self)
        self.timer.setInterval(60000)
        self.timer.timeout.connect(self.reload)
        self.click_timer = QTimer(self)
        self.click_timer.setSingleShot(True)
        self.click_timer.setTimerType(Qt.PreciseTimer)
        self.click_timer.timeout.connect(self._open_pending)
        self._clicked_instrument = None
        self.window.widget_visibility_changed.connect(self._visibility_changed)

    def request_row(self, surface, row, block=0):
        if not self.window.view_options.chart_enabled:
            return
        self._clicked_instrument = self.window.quotes.instrument_at(surface, row, block)
        if self._clicked_instrument:
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

    def open_row(self, surface, row, block=0):
        if not self.window.view_options.chart_enabled:
            return
        instrument = self.window.quotes.instrument_at(surface, row, block)
        if instrument is None:
            return
        self._open(instrument)

    def _open(self, instrument):
        if self.dialog is None:
            self.dialog = HistoryDialog(self.window)
            self.dialog.view_changed.connect(self.reload)
            self.dialog.refresh_requested.connect(self.reload)
            self.dialog.finished.connect(self._closed)
        self.instrument = dict(instrument)
        self.dialog.setWindowTitle(f"{instrument.get('name') or instrument.get('code')} · 行情图表")
        self._apply_theme()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()
        self.timer.start()
        self.reload()

    def _apply_theme(self):
        if self.dialog is None:
            return
        mode = self.window.view_options.color_mode
        dark = mode == "dark" or (mode == "system" and is_dark_theme())
        self.dialog.setStyleSheet(build_settings_stylesheet(dark))
        self.dialog.setPalette(theme_palette(dark, self.window.palette()))
        self.dialog.chart.up_color = self.window.up_color
        self.dialog.chart.down_color = self.window.down_color

    def _options_changed(self):
        if not self.window.view_options.chart_enabled:
            self.click_timer.stop()
            if self.dialog:
                self.dialog.close()
        else:
            self._apply_theme()

    def _closed(self, *_args):
        self._generation += 1
        self._pending = None
        self.timer.stop()

    def reload(self):
        if not self.dialog or not self.dialog.isVisible() or not self.instrument or not self.window.view_options.chart_enabled:
            return
        self._generation += 1
        request = (self._generation, dict(self.instrument), self.dialog.view, self.window.data_source)
        key = (self.instrument.get("market"), self.instrument.get("code"), self.dialog.view, self.window.data_source)
        if key != self._displayed_key:
            self.dialog.chart.set_data((), self.dialog.view, self.instrument)
        self.dialog.status.setText("更新中…" if self.dialog.chart.bars else "加载中…")
        if self._busy:
            self._pending = request  # At most one worker; latest selection wins.
        else:
            self._start(request)

    def _start(self, request):
        self._busy = True
        self._pending = None
        threading.Thread(target=self._worker, args=(request,), daemon=True).start()

    def _worker(self, request):
        generation, instrument, view, source = request
        try:
            result = self.cache.get(instrument, view, source)
        except Exception:
            result = BarResult(message="暂无可用历史数据 · 历史数据处理失败")
        try:
            self.result_ready.emit((generation, instrument, view, source, result))
        except RuntimeError:
            pass  # Parent was destroyed while the network request was running.

    def _accept(self, payload):
        self._busy = False
        generation, instrument, view, preferred, result = payload
        if generation == self._generation and self.dialog and self.dialog.isVisible():
            self.dialog.chart.set_data(result.bars, view, instrument)
            self._displayed_key = (instrument.get("market"), instrument.get("code"), view, preferred)
            provider = {"sina": "新浪", "eastmoney": "东方财富"}.get(result.source, "")
            pieces = [f"数据来源：{provider}" if provider else "", "未复权"]
            if result.source and result.source != preferred:
                pieces.append("已使用备用数据源")
            if result.cached:
                pieces.append("缓存（更新失败）" if result.stale else "本地缓存")
            if result.bars:
                visible = self.dialog.chart.bars
                pieces.append(f"{visible[0].time} — {visible[-1].time}")
                if view == "daily" and len(result.bars) < 89:
                    pieces.append("历史不足时，部分均线从满足周期处开始显示")
            pieces.append(result.message)
            self.dialog.status.setText(" · ".join(piece for piece in pieces if piece))
        if self._pending is not None:
            self._start(self._pending)
