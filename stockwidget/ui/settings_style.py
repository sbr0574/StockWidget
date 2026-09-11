"""设置窗口的主题样式和颜色预览图标。"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap

from stockwidget.ui.theme import accent_color, accent_rgba

COLOR_SWATCH_SIZE = 12
_FLAT_GROUPS = ("gb_data", "gb_data_setting", "gb_icon", "gb_fcn", "gb_opacity",
                "gb_color", "gb_text", "gb_tabel", "gb_hotkeys")

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



def build_settings_stylesheet(dark: bool) -> str:
    """按系统深浅色生成设置窗口样式表"""
    if dark:
        sep = "rgba(255, 255, 255, 0.35)"
        header_bg, header_line = "rgba(255, 255, 255, 0.10)", "rgba(255, 255, 255, 0.30)"
        empty_hint = "rgba(255, 255, 255, 0.28)"
        icon_hover_bg = accent_rgba(0.12)
        icon_pressed_bg = accent_rgba(0.22)
        icon_selected = accent_rgba(0.20)
        icon_selected_hover = accent_rgba(0.26)
        icon_selected_pressed = accent_rgba(0.34)
        color_bg = "rgba(255, 255, 255, 0.10)"
        color_hover_bg = accent_rgba(0.12)
        color_pressed_bg = accent_rgba(0.22)
        color_disabled_bg = "rgba(255, 255, 255, 0.05)"
        color_disabled_text = "rgba(255, 255, 255, 0.35)"
    else:
        sep = "rgba(0, 0, 0, 0.25)"
        header_bg, header_line = "rgba(0, 0, 0, 0.06)", "rgba(0, 0, 0, 0.20)"
        empty_hint = "rgba(0, 0, 0, 0.28)"
        icon_hover_bg = accent_rgba(0.08)
        icon_pressed_bg = accent_rgba(0.16)
        icon_selected = accent_rgba(0.12)
        icon_selected_hover = accent_rgba(0.18)
        icon_selected_pressed = accent_rgba(0.24)
        color_bg = "rgba(0, 0, 0, 0.07)"
        color_hover_bg = accent_rgba(0.08)
        color_pressed_bg = accent_rgba(0.16)
        color_disabled_bg = "rgba(0, 0, 0, 0.04)"
        color_disabled_text = "rgba(0, 0, 0, 0.35)"

    flat_boxes = ",\n".join(f"QGroupBox#{n}" for n in _FLAT_GROUPS)
    flat_titles = ",\n".join(f"QGroupBox#{n}::title" for n in _FLAT_GROUPS)

    icon_selectors = (
        "QPushButton#btn_icon_default",
        "QPushButton#btn_icon_lightG",
        "QPushButton#btn_icon_dark",
        "QPushButton#btn_icon_darkG",
        "QPushButton#btn_icon_custom",
    )
    color_selectors = (
        "QPushButton#btn_fg_color",
        "QPushButton#btn_bg_color",
        "QPushButton#btn_up_color",
        "QPushButton#btn_down_color",
        "QPushButton#btn_neutral_color",
        "QPushButton#btn_add",
        "QPushButton#btn_del",
        "QPushButton#btn_top",
        "QPushButton#btn_check_update",
        "QPushButton#btn_clear_watchlist",
        "QPushButton#btn_reset_appearance",
        "QPushButton#btn_reset_settings",
        "QPushButton#btn_open_cache_dir",
    )
    icon_buttons = ",\n".join(icon_selectors)
    icon_hover = ",\n".join(f"{selector}:hover" for selector in icon_selectors)
    icon_pressed = ",\n".join(f"{selector}:pressed" for selector in icon_selectors)
    icon_checked = ",\n".join(f"{selector}:checked" for selector in icon_selectors)
    icon_checked_hover = ",\n".join(
        f"{selector}:checked:hover" for selector in icon_selectors
    )
    icon_checked_pressed = ",\n".join(
        f"{selector}:checked:pressed" for selector in icon_selectors
    )
    color_buttons = ",\n".join(color_selectors)
    color_hover = ",\n".join(f"{selector}:hover" for selector in color_selectors)
    color_pressed = ",\n".join(f"{selector}:pressed" for selector in color_selectors)
    color_disabled = ",\n".join(f"{selector}:disabled" for selector in color_selectors)

    choice_buttons = f"""
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
{icon_hover} {{
    background-color: {icon_hover_bg};
}}
{icon_pressed} {{
    background-color: {icon_pressed_bg};
}}
{icon_checked} {{
    background-color: {icon_selected};
    border: 2px solid {accent_color().name()};
}}
{icon_checked_hover} {{
    background-color: {icon_selected_hover};
}}
{icon_checked_pressed} {{
    background-color: {icon_selected_pressed};
}}
{color_buttons} {{
    background-color: {color_bg};
    border: none;
    border-radius: 6px;
    padding: 3px;
}}
{color_hover} {{
    background-color: {color_hover_bg};
}}
{color_pressed} {{
    background-color: {color_pressed_bg};
}}
{color_disabled} {{
    background-color: {color_disabled_bg};
    color: {color_disabled_text};
}}
"""

    return f"""
{flat_boxes} {{
    border: none;
    border-top: 1px solid {sep};
    margin-top: 9px;
    padding-top: 0px;
}}
{flat_titles} {{
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
{choice_buttons}
"""

