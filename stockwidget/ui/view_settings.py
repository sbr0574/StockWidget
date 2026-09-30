"""Behavior for settings groups whose controls and layout live in settings.ui."""

import sys
from contextlib import ExitStack

from PySide6.QtCore import QEvent, QObject, QSignalBlocker, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox, QColorDialog, QComboBox, QFontComboBox, QGroupBox,
    QLabel, QPushButton, QSlider, QSpinBox, QWidget,
)

from stockwidget.core.view_options import PAGE_MODES
from stockwidget.ui.metric_pool import MetricPoolWidget
from stockwidget.ui.settings_style import color_swatch_icon


class _SettingsGroup(QObject):
    """Bind a native Designer QGroupBox without replacing its Qt behavior."""

    def __init__(self, group):
        super().__init__(group)
        self.group = group

    def _publish(self, **controls):
        for name, control in controls.items():
            setattr(self, name, control)
            setattr(self.group, name, control)

    def bind(self, source):
        self.source = source
        self._publish(body=self._control(QWidget, self.group.objectName() + "_body"))

    def _control(self, kind, name):
        control = self.group.findChild(kind, name)
        if control is None:
            raise ValueError(f"settings.ui 缺少控件：{name}")
        return control

    def _value_control(self, key, kind=QSpinBox):
        control = self._control(kind, key)
        control.valueChanged.connect(lambda value: self.source.set_view_options(**{key: value}))
        return control

    def _connect_sync(self):
        self.source.view_options_changed.connect(self.sync)
        self.source.taskbar_options_changed.connect(self.sync)
        self.source.presentation_changed.connect(self.sync)
        self.source.display_flags_changed.connect(self.sync)
        self.sync()

    def _page_controls(self, prefix):
        self._publish(mode=self._control(QComboBox, f"{prefix}_page_mode"))
        for index, (value, _label) in enumerate(PAGE_MODES):
            self.mode.setItemData(index, value)
        self.mode.currentIndexChanged.connect(
            lambda _: self.source.set_view_options(**{f"{prefix}_page_mode": self.mode.currentData()}))
        self._publish(interval=self._value_control(f"{prefix}_page_interval"))

    def sync(self):
        with ExitStack() as stack:
            stack.enter_context(QSignalBlocker(self.group))
            for control in self.group.findChildren(QWidget):
                stack.enter_context(QSignalBlocker(control))
            self._sync()


class _SyncedGroup(_SettingsGroup):
    def bind(self, source, key):
        super().bind(source)
        self.sync_key = key
        self.group.toggled.connect(lambda checked: source.set_view_options(**{key: checked}))
        # Native unchecked groups disable their body after toggled; sync means
        # the opposite, so restore editing after the native click/polish ends.
        self.group.clicked.connect(self.sync)
        self.group.installEventFilter(self)

    def eventFilter(self, watched, event):
        if event.type() in (QEvent.Polish, QEvent.StyleChange, QEvent.Show, QEvent.EnabledChange):
            # Apply after Qt has finished its native group enable/polish pass.
            QTimer.singleShot(0, self, self.sync)
        return False

    def _sync_switch(self):
        self.group.setChecked(getattr(self.source.view_options, self.sync_key))
        self.body.setEnabled(not self.group.isChecked())


class FloatRowSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self.group.toggled.connect(lambda value: source.set_view_options(float_paging_enabled=value))
        self._publish(rows=self._value_control("float_max_rows"))
        self._connect_sync()

    def _sync(self):
        options = self.source.view_options
        self.group.setChecked(options.float_paging_enabled)
        self.body.setEnabled(options.float_paging_enabled)
        self.rows.setValue(options.float_max_rows)


class PagingSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self._page_controls("float")
        self._connect_sync()

    def _sync(self):
        options = self.source.view_options
        self.body.setEnabled(options.float_paging_enabled)
        self.mode.setCurrentIndex(self.mode.findData(options.float_page_mode))
        self.interval.setValue(options.float_page_interval)
        self.interval.setEnabled(options.float_page_mode == "auto")


class TaskbarMetricSettings(_SyncedGroup):
    def bind(self, source):
        super().bind(source, "taskbar_sync_metrics")
        self._publish(metric_pool=self._control(MetricPoolWidget, "taskbar_metric_pool"))
        self.metric_pool.visible_metrics_changed.connect(self._metrics_changed)
        self._connect_sync()

    def _metrics_changed(self, metrics):
        if not self.source.view_options.taskbar_sync_metrics:
            self.source.set_view_options(taskbar_metrics=metrics)

    def _sync(self):
        self._sync_switch()
        self.metric_pool.set_visible_metrics(self.source.get_surface_metrics("taskbar"))


class TaskbarStyleSettings(_SyncedGroup):
    def bind(self, source):
        super().bind(source, "taskbar_sync_appearance")
        self._publish(font_family=self._control(QFontComboBox, "taskbar_font_family"))
        self.font_family.currentFontChanged.connect(
            lambda font: source.set_view_options(taskbar_font_family=font.family()))
        self._publish(font_size=self._value_control("taskbar_font_size", QSlider),
                      font_size_label=self._control(QLabel, "taskbar_font_size_label"),
                      color=self._control(QPushButton, "btn_taskbar_color"))
        self.color.clicked.connect(self._pick_color)
        self._publish(unicolor=self._control(QCheckBox, "taskbar_unicolor"))
        self.unicolor.toggled.connect(lambda value: source.set_view_options(taskbar_unicolor=value))
        self._publish(opacity=self._value_control("taskbar_opacity_pct", QSlider),
                      opacity_label=self._control(QLabel, "taskbar_opacity_pct_label"))
        self._connect_sync()

    def _pick_color(self):
        color = QColorDialog.getColor(QColor(self.source.view_options.taskbar_color), self.group, "任务栏文字颜色")
        if color.isValid():
            self.source.set_view_options(taskbar_color=color.name())

    def _sync(self):
        self._sync_switch()
        font, color, opacity, unicolor = self.source.get_taskbar_appearance()
        self.font_family.setCurrentFont(font)
        self.font_size.setValue(font.pointSize())
        self.font_size_label.setText(f"{font.pointSize()} pt")
        self.color.setToolTip(f"文字颜色: {color.name()}")
        self.color.setIcon(color_swatch_icon(color, self.group.devicePixelRatioF()))
        self.opacity.setValue(opacity)
        self.opacity_label.setText(f"{opacity}%")
        self.unicolor.setChecked(unicolor)


class TaskbarPagingSettings(_SyncedGroup):
    def bind(self, source):
        super().bind(source, "taskbar_sync_paging")
        self._page_controls("taskbar")
        self._connect_sync()

    def _sync(self):
        self._sync_switch()
        mode, seconds = self.source.get_page_settings("taskbar")
        self.mode.setCurrentIndex(self.mode.findData(mode))
        self.interval.setValue(seconds)
        self.interval.setEnabled(mode == "auto")


class FloatSplitSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self.group.toggled.connect(lambda value: source.set_view_options(float_split_enabled=value))
        self._publish(separator=self._control(QCheckBox, "float_split_separator"))
        self.separator.toggled.connect(lambda value: source.set_view_options(float_split_separator=value))
        self._connect_sync()

    def _sync(self):
        options = self.source.view_options
        self.group.setChecked(options.float_split_enabled)
        self.body.setEnabled(options.float_split_enabled)
        self.separator.setChecked(options.float_split_separator)


class TaskbarSplitSettings(_SyncedGroup):
    def bind(self, source):
        super().bind(source, "taskbar_sync_split")
        self._publish(enabled=self._control(QCheckBox, "taskbar_split_enabled"),
                      separator=self._control(QCheckBox, "taskbar_split_separator"))
        self.enabled.toggled.connect(lambda value: source.set_view_options(taskbar_split_enabled=value))
        self.separator.toggled.connect(lambda value: source.set_view_options(taskbar_split_separator=value))
        self._connect_sync()

    def _sync(self):
        self._sync_switch()
        enabled, separator = self.source.get_split_settings("taskbar")
        self.enabled.setChecked(enabled)
        self.separator.setChecked(separator)
        self.separator.setEnabled(enabled)


class TaskbarSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self._publish(left=self._control(QWidget, "taskbar_controls"),
                      dual=self._control(QCheckBox, "taskbar_dual_open"))
        self.dual.toggled.connect(lambda value: source.set_view_options(taskbar_dual_open=value))
        self._publish(rows=self._value_control("taskbar_rows"),
                      offset=self._control(QSpinBox, "taskbar_offset"))
        self.offset.valueChanged.connect(source.set_taskbar_offset)
        self.bindings = []
        for name, kind in (("metrics", TaskbarMetricSettings), ("style", TaskbarStyleSettings),
                           ("paging", TaskbarPagingSettings), ("split", TaskbarSplitSettings)):
            object_name = "taskbar_metric_settings" if name == "metrics" else f"taskbar_{name}_settings"
            group = self._control(QGroupBox, object_name)
            self._publish(**{name: group})
            binding = kind(group)
            binding.bind(source)
            self.bindings.append(binding)
        self._publish(metric_pool=self.metrics.metric_pool)
        self.group.toggled.connect(self._enable_taskbar)
        source.taskbar_status_changed.connect(self.group.setToolTip)
        self._connect_sync()

    def _enable_taskbar(self, enabled):
        self.source.set_view_options(taskbar_enabled=enabled)
        self.source.set_display_mode(("both" if self.source.view_options.taskbar_dual_open else "taskbar") if enabled else "float")

    def _sync(self):
        source = self.source
        self.group.setChecked(source.view_options.taskbar_enabled)
        self.body.setEnabled(source.view_options.taskbar_enabled)
        self.dual.setChecked(source.view_options.taskbar_dual_open)
        self.rows.setValue(source.view_options.taskbar_rows)
        self.offset.setValue(source.taskbar_offset)
        self.group.setEnabled(sys.platform == "win32")
        self.group.setToolTip(source.taskbar_status)
        for binding in self.bindings:
            binding.sync()
