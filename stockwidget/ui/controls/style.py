"""系统调色板、设置页样式与颜色预览图标，供自绘控件共用。"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import QApplication


def accent_color() -> QColor:
    palette = QApplication.palette()
    return palette.color(QPalette.ColorGroup.Active, QPalette.ColorRole.Accent)


def accent_rgba(alpha: float) -> str:
    color = accent_color()
    return f"rgba({color.red()}, {color.green()}, {color.blue()}, {alpha:.2f})"


COLOR_SWATCH_SIZE = 12
LINUX_FONT_RULES = """
QWidget { font-size: 12px; }
"""

def color_swatch_icon(color: QColor, device_pixel_ratio: float = 1.0) -> QIcon:
    """按屏幕像素比绘制无描边圆形色标，避免高 DPI 缩放发糊。"""
    ratio = max(1.0, float(device_pixel_ratio))
    pixel_size = max(COLOR_SWATCH_SIZE, round(COLOR_SWATCH_SIZE * ratio))
    ratio = pixel_size / COLOR_SWATCH_SIZE
    pixmap = QPixmap(pixel_size, pixel_size)
    pixmap.setDevicePixelRatio(ratio)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(color)
    painter.drawEllipse(1, 1, COLOR_SWATCH_SIZE - 2, COLOR_SWATCH_SIZE - 2)
    painter.end()

    icon = QIcon()
    icon.addPixmap(pixmap, QIcon.Mode.Normal, QIcon.State.Off)
    icon.addPixmap(pixmap, QIcon.Mode.Disabled, QIcon.State.Off)
    return icon


def build_settings_stylesheet(dark: bool, *, linux_fonts: bool = False) -> str:
    """按系统深浅色生成设置窗口样式表"""
    # 在设置窗口的样式表中直接匹配每个子控件，覆盖 Linux 桌面主题的字号。
    # 仅对父窗口 setFont 无法覆盖子控件的样式字体；后创建的编辑器也要匹配。
    # 保留字体家族与字重，Qt 继续按屏幕像素比缩放这些逻辑像素。
    font_rules = LINUX_FONT_RULES if linux_fonts else ""
    if dark:
        sep = "rgba(255, 255, 255, 0.35)"
        header_bg, header_line = "rgba(255, 255, 255, 0.10)", "rgba(255, 255, 255, 0.30)"
        empty_hint = "rgba(255, 255, 255, 0.28)"
        selected_bg = accent_rgba(0.20)
        selected_hover_bg = accent_rgba(0.26)
        selected_pressed_bg = accent_rgba(0.34)
        color_bg = "rgba(255, 255, 255, 0.10)"
        color_hover_bg = accent_rgba(0.12)
        color_pressed_bg = accent_rgba(0.22)
        color_disabled_bg = "rgba(255, 255, 255, 0.05)"
        color_disabled_text = "rgba(255, 255, 255, 0.35)"
    else:
        sep = "rgba(0, 0, 0, 0.25)"
        header_bg, header_line = "rgba(0, 0, 0, 0.06)", "rgba(0, 0, 0, 0.20)"
        empty_hint = "rgba(0, 0, 0, 0.28)"
        selected_bg = accent_rgba(0.12)
        selected_hover_bg = accent_rgba(0.18)
        selected_pressed_bg = accent_rgba(0.24)
        color_bg = "rgba(0, 0, 0, 0.07)"
        color_hover_bg = accent_rgba(0.08)
        color_pressed_bg = accent_rgba(0.16)
        color_disabled_bg = "rgba(0, 0, 0, 0.04)"
        color_disabled_text = "rgba(0, 0, 0, 0.35)"

    buttons = f"""
QPushButton {{
    background-color: {color_bg};
    border: 1px solid transparent;
    border-radius: 6px;
    padding: 3px;
}}
QPushButton:hover {{
    background-color: {color_hover_bg};
}}
QPushButton:pressed {{
    background-color: {color_pressed_bg};
}}
QPushButton:checked {{
    background-color: {selected_bg};
    border: 2px solid {accent_color().name()};
}}
QPushButton:checked:hover {{
    background-color: {selected_hover_bg};
}}
QPushButton:checked:pressed {{
    background-color: {selected_pressed_bg};
}}
QPushButton:disabled {{
    background-color: {color_disabled_bg};
    color: {color_disabled_text};
}}
"""

    # Icon choices keep their original transparent tiles and selection outline.
    icon_selectors = tuple(f"QPushButton#{name}" for name in (
        "btn_icon_default", "btn_icon_lightG", "btn_icon_dark", "btn_icon_darkG", "btn_icon_custom"))
    icon_buttons = ",\n".join(icon_selectors)
    icon_states = "\n".join(
        f"{', '.join(selector + state for selector in icon_selectors)} {{ {rules} }}"
        for state, rules in (
            (":hover", f"background-color: {color_hover_bg};"),
            (":pressed", f"background-color: {color_pressed_bg};"),
            (":checked", f"background-color: {selected_bg}; border: 2px solid {accent_color().name()};"),
            (":checked:hover", f"background-color: {selected_hover_bg};"),
            (":checked:pressed", f"background-color: {selected_pressed_bg};"),
        )
    )
    icon_buttons = f"""
{icon_buttons} {{
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 3px;
}}
QPushButton#btn_icon_custom {{
    font-size: 22px;
    font-weight: 300;
}}
{icon_states}
"""

    return f"""
{font_rules}
QGroupBox[flat="true"] {{
    border: none;
    border-top: 1px solid {sep};
    margin-top: 9px;
    padding-top: 0px;
}}
QGroupBox[flat="true"]::title {{
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 2 8px;
}}
QTableWidget#list_codes QHeaderView::section {{
    background-color: {header_bg};
    border: none;
    border-bottom: 1px solid {header_line};
    padding: 4px 8px;
    font-weight: 600;
}}
QLabel#empty_watchlist_hint {{
    color: {empty_hint};
    background: transparent;
    font-size: 18px;
    font-weight: 500;
}}
{buttons}
{icon_buttons}
"""
