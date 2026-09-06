"""成交量/成交额的单位设置悬浮面板。"""

from PySide6.QtCore import QPoint, Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QRadioButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

# 单位模式: cn=中文, en=英文, auto=自动（美股/国际指数用英文，其余用中文）
UNIT_OPTIONS = (
    ("cn", "中文"),
    ("en", "英文"),
    ("auto", "自动"),
)


class UnitSettingsPanel(QFrame):
    """单击“成交量/成交额”指标块时弹出的单位选择面板。"""

    unit_mode_changed = Signal(str)
    # 面板关闭（含点击外部空白关闭）时发出，用于清除指标块的选中状态
    panel_closed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setObjectName("unit_settings_panel")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(7)

        self.title_label = QLabel("数值单位：", self)
        layout.addWidget(self.title_label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        self.radio_buttons = {}
        for value, text in UNIT_OPTIONS:
            radio = QRadioButton(text, self)
            radio.setObjectName(f"unit_option_{value}")
            radio.toggled.connect(
                lambda checked, v=value: self._on_toggled(v, checked)
            )
            self.radio_buttons[value] = radio
            row.addWidget(radio)
        row.addStretch(1)
        layout.addLayout(row)

        self.radio_buttons["auto"].setToolTip(
            "美股、国际指数使用英文单位，国内、港股等使用中文单位"
        )

        self.set_theme(False)
        self.sync_from(None)

    def sync_from(self, win):
        """按浮窗当前单位模式同步单选按钮；win 为 None 时使用默认值。"""
        mode = getattr(win, "unit_mode", "auto")
        if mode not in self.radio_buttons:
            mode = "auto"
        for value, radio in self.radio_buttons.items():
            radio.blockSignals(True)
            radio.setChecked(value == mode)
            radio.blockSignals(False)

    def _on_toggled(self, value: str, checked: bool):
        if checked:
            self.unit_mode_changed.emit(value)

    def set_theme(self, dark: bool):
        if dark:
            background = "rgb(44, 44, 46)"
            foreground = "rgb(242, 242, 247)"
            border = "rgba(255, 255, 255, 0.28)"
        else:
            background = "rgb(250, 250, 250)"
            foreground = "rgb(28, 28, 30)"
            border = "rgba(0, 0, 0, 0.28)"
        self.setStyleSheet(f"""
            QFrame#unit_settings_panel {{
                background-color: {background};
                color: {foreground};
                border: 1px solid {border};
                border-radius: 8px;
            }}
            QFrame#unit_settings_panel QLabel,
            QFrame#unit_settings_panel QRadioButton {{
                color: {foreground};
                border: none;
                background: transparent;
            }}
        """)

    def show_for(self, anchor: QWidget):
        """在锚点控件下方弹出并限制在屏幕范围内。"""
        self.adjustSize()
        position = anchor.mapToGlobal(QPoint(0, anchor.height() + 2))
        screen = anchor.screen()
        available = screen.availableGeometry() if screen is not None else None
        if available is not None:
            max_x = max(available.left(), available.right() - self.width() + 1)
            x = min(max(position.x(), available.left()), max_x)
            below_y = position.y()
            above_y = anchor.mapToGlobal(QPoint(0, -self.height() - 2)).y()
            y = below_y if below_y + self.height() <= available.bottom() + 1 else above_y
            max_y = max(available.top(), available.bottom() - self.height() + 1)
            y = min(max(y, available.top()), max_y)
            position = QPoint(x, y)
        self.move(position)
        self.show()
        self.raise_()

    def hideEvent(self, event):
        super().hideEvent(event)
        self.panel_closed.emit()
