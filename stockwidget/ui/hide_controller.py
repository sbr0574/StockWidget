"""Timed and stale-data hiding share the widget's visibility entry point."""

from datetime import datetime
import math
import time

from PySide6.QtCore import QObject, QTimer, Qt

from stockwidget.core.hide_rules import all_quotes_stale, due_hide_times


class HideController(QObject):
    def __init__(self, source):
        super().__init__(source)
        self.source = source
        self._last_check = datetime.now().astimezone()
        self._fired = set()
        self._manual_show = False
        self._deadline = None
        self.schedule_timer = QTimer(self)
        self.schedule_timer.setInterval(1000)
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.countdown_timer = QTimer(self)
        self.countdown_timer.setTimerType(Qt.PreciseTimer)
        self.countdown_timer.setInterval(100)
        self.countdown_timer.timeout.connect(self.check_countdown)
        source.widget_visibility_changed.connect(self._visibility_changed)
        self.configure()

    def configure(self):
        if self.source.scheduled_hide_enabled and self.source.scheduled_hide_times:
            if not self.schedule_timer.isActive():
                self._last_check = datetime.now().astimezone()
                self.schedule_timer.start()
        else:
            self.schedule_timer.stop()
        if not self.source.auto_hide_enabled:
            self.cancel_countdown()

    def check_schedule(self, now=None):
        now = now or datetime.now().astimezone()
        due = due_hide_times(self.source.scheduled_hide_times, self._last_check, now)
        self._last_check = now
        self._fired = {event for event in self._fired if event[0] == now.date()}
        if not self.source.scheduled_hide_enabled:
            return
        pending = {(now.date(), value) for value in due} - self._fired
        self._fired.update(pending)  # 隐藏时也消耗当天事件，呼出后不会重复隐藏。
        if pending and self.source.widget_visible:
            self.source.hide_widget()

    def _visibility_changed(self):
        self.cancel_countdown()
        self._manual_show = self.source.widget_visible

    def quotes_refreshed(self, data):
        source = self.source
        if (not source.auto_hide_enabled or not source.widget_visible
                or not all_quotes_stale(data, source.checked_codes, time.time())):
            self.cancel_countdown()
            return
        if self._deadline is not None:
            return  # 连续收到旧行情不会重置倒计时。
        if self._manual_show:
            self._deadline = time.monotonic() + 5
            self.check_countdown()
            self.countdown_timer.start()
        else:
            source.hide_widget()

    def check_countdown(self):
        if self._deadline is None:
            return
        if not self.source.auto_hide_enabled or not self.source.widget_visible:
            self.cancel_countdown()
            return
        remaining = self._deadline - time.monotonic()
        if remaining <= 0:
            self.source.hide_widget()
            return
        label = self.source.hide_notice
        text = f"行情已超过30秒未更新，{math.ceil(remaining)}秒后隐藏"
        if label.text() != text:
            label.setText(text)
            label.show()
            self.source._defer_fit()
            self.source.presentation_changed.emit()

    def cancel_countdown(self):
        self._deadline = None
        self.countdown_timer.stop()
        label = self.source.hide_notice
        if label.text():
            label.clear()
            label.hide()
            self.source._defer_fit()
            self.source.presentation_changed.emit()
