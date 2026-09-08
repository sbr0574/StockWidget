"""从应用调色板读取系统强调色，供自绘控件和样式表共用。"""

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


def accent_color() -> QColor:
    palette = QApplication.palette()
    return palette.color(QPalette.ColorGroup.Active, QPalette.ColorRole.Accent)


def accent_rgba(alpha: float) -> str:
    color = accent_color()
    return f"rgba({color.red()}, {color.green()}, {color.blue()}, {alpha:.2f})"
