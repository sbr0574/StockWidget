"""Designer 原生组框的参数绑定、同步及编辑状态；不创建静态布局。"""

from contextlib import ExitStack
import sys

from PySide6.QtCore import QObject, QSignalBlocker, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox,
    QColorDialog,
    QComboBox,
    QFontComboBox,
    QGroupBox,
    QLabel,
    QPushButton,
    QSlider,
    QSpinBox,
    QWidget,
)

from stockwidget.core.view_options import PAGE_MODES, taskbar_content_height, taskbar_font_size_limit
from stockwidget.platform.taskbar import find_taskbar
from stockwidget.ui.controls.metrics import MetricPoolWidget
from stockwidget.ui.controls.style import set_color_button


class _SettingsGroup(QObject):
    """Bind Designer controls without creating or replacing their layouts."""

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
        self.source.configuration_changed.connect(self.sync)
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
    def _connect_sync(self):
        # TaskbarSettings owns source subscriptions and refreshes all children.
        # Child groups only own their controls (and the style geometry timer).
        self.sync()

    def bind(self, source, key):
        super().bind(source)
        self.sync_key = key
        self._publish(sync_toggle=self._control(QCheckBox, key))
        self.sync_toggle.toggled.connect(lambda checked: source.set_view_options(**{key: checked}))

    def _sync_switch(self):
        checked = getattr(self.source.view_options, self.sync_key)
        self.sync_toggle.setChecked(checked)
        self.body.setVisible(not checked)
        self.body.setEnabled(not checked)


class FloatRowSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self.group.toggled.connect(lambda value: source.set_view_options(float_paging_enabled=value))
        self._publish(rows=self._value_control("float_max_rows"))
        self._publish(paging=self._control(QCheckBox, "float_paging_enabled"),
                      paging_body=self._control(QWidget, "float_paging_settings_body"),
                      auto=self._control(QCheckBox, "float_page_mode"),
                      interval=self._value_control("float_page_interval"))
        self.paging.toggled.connect(self._set_page_mode)
        self.auto.toggled.connect(self._set_page_mode)
        self._connect_sync()

    def _set_page_mode(self, _checked):
        mode = "auto" if self.auto.isChecked() else "manual"
        self.source.set_view_options(float_page_mode=mode if self.paging.isChecked() else "first")

    def _sync(self):
        options = self.source.view_options
        self.group.setChecked(options.float_paging_enabled)
        self.body.setEnabled(options.float_paging_enabled)
        self.rows.setValue(options.float_max_rows)
        paging = options.float_page_mode != "first"
        self.paging.setChecked(paging)
        self.paging_body.setVisible(paging)
        self.paging_body.setEnabled(paging)
        self.auto.setChecked(options.float_page_mode == "auto")
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
        self._publish(unicolor=self._control(QCheckBox, "taskbar_unicolor"),
                      auto_color=self._control(QCheckBox, "taskbar_auto_color"))
        self.unicolor.toggled.connect(lambda value: source.set_view_options(taskbar_unicolor=value))
        self.auto_color.toggled.connect(lambda value: source.set_view_options(taskbar_auto_color=value))
        self._publish(opacity=self._value_control("taskbar_opacity_pct", QSlider),
                      opacity_label=self._control(QLabel, "taskbar_opacity_pct_label"))
        self._taskbar_geometry = None
        self._geometry_timer = QTimer(self)
        self._geometry_timer.setInterval(1000)
        self._geometry_timer.timeout.connect(lambda: self.sync() if self.group.isVisible() else None)
        self._geometry_timer.start()
        self._connect_sync()

    def _font_size_limit(self):
        try:
            area = find_taskbar() if sys.platform == "win32" else None
        except OSError:
            area = None
        if area:
            self._taskbar_geometry = (taskbar_content_height(area.height, area.dpi), area.dpi)
        if self._taskbar_geometry is None:
            dpi = max(96, round(96 * self.group.devicePixelRatioF()))
            self._taskbar_geometry = (round(44 * dpi / 96), dpi)
        height, dpi = self._taskbar_geometry
        return taskbar_font_size_limit(height, dpi, self.source.view_options.taskbar_rows)

    def _pick_color(self):
        color = QColorDialog.getColor(QColor(self.source.view_options.taskbar_color), self.group, "任务栏文字颜色")
        if color.isValid():
            self.source.set_view_options(taskbar_color=color.name())

    def _sync(self):
        self._sync_switch()
        font, color, opacity, unicolor = self.source.get_taskbar_appearance()
        self.font_family.setCurrentFont(font)
        limit = self._font_size_limit()
        self.font_size.setMaximum(limit)
        self.font_size.setValue(min(font.pointSize(), limit))
        self.font_size_label.setText(f"{self.font_size.value()} pt")
        self.font_size.setToolTip(f"当前行数与任务栏高度下，最大字号为 {limit} pt；缩放变化后自动更新。")
        set_color_button(self.color, color, "任务栏文字颜色")
        self.opacity.setValue(opacity)
        self.opacity_label.setText(f"{opacity}%")
        self.unicolor.setChecked(unicolor)
        automatic = self.source.view_options.taskbar_auto_color
        self.auto_color.setChecked(automatic)
        self.color.setEnabled(not automatic)
        self.unicolor.setEnabled(not automatic)


class TaskbarPagingSettings(_SyncedGroup):
    def bind(self, source):
        super().bind(source, "taskbar_sync_paging")
        self._page_controls("taskbar")
        self._connect_sync()

    def _sync(self):
        self._sync_switch()
        mode, seconds = self.source.view_options.page_settings("taskbar")
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
        enabled, separator = self.source.view_options.split_settings("taskbar")
        self.enabled.setChecked(enabled)
        self.separator.setChecked(separator)
        self.separator.setEnabled(enabled)


class TaskbarSettings(_SettingsGroup):
    def bind(self, source):
        super().bind(source)
        self._publish(enabled=self._control(QCheckBox, "taskbar_enabled"),
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
        self.enabled.toggled.connect(lambda value: source.set_view_options(taskbar_enabled=value))
        source.taskbar_status_changed.connect(self.enabled.setToolTip)
        self._connect_sync()

    def _sync(self):
        source = self.source
        self.enabled.setChecked(source.view_options.taskbar_enabled)
        self.body.setVisible(source.view_options.taskbar_enabled)
        self.body.setEnabled(source.view_options.taskbar_enabled)
        self.dual.setChecked(source.view_options.taskbar_dual_open)
        self.rows.setValue(source.view_options.taskbar_rows)
        self.offset.setValue(source.taskbar_offset)
        self.group.setEnabled(sys.platform == "win32")
        self.enabled.setToolTip(source.taskbar_status)
        for binding in self.bindings:
            binding.sync()
