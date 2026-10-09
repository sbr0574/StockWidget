"""系统调色板、设置页样式与颜色预览图标，供自绘控件共用。"""

from PySide6.QtCore import QByteArray, QObject, QRectF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QPalette, QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication, QWidget


def accent_color() -> QColor:
    palette = QApplication.palette()
    return palette.color(QPalette.ColorGroup.Active, QPalette.ColorRole.Accent)


def is_dark_theme() -> bool:
    scheme = QGuiApplication.styleHints().colorScheme()
    if scheme != Qt.ColorScheme.Unknown:
        return scheme == Qt.ColorScheme.Dark
    app = QApplication.instance()
    theme = getattr(app, "theme", None)
    palette = theme.system_palette if theme and theme.mode != "system" else app.palette()
    return palette.color(QPalette.ColorRole.Window).lightness() < 128


def theme_palette(dark: bool, base: QPalette) -> QPalette:
    palette = QPalette(base)
    colors = ("#222222", "#eeeeee", "#292929", "#303030") if dark else ("#f4f4f4", "#202020", "#ffffff", "#eeeeee")
    for role, color in ((QPalette.Window, colors[0]), (QPalette.WindowText, colors[1]),
                        (QPalette.Base, colors[2]), (QPalette.Text, colors[1]),
                        (QPalette.Button, colors[3]), (QPalette.ButtonText, colors[1]),
                        (QPalette.ToolTipBase, colors[2]), (QPalette.ToolTipText, colors[1])):
        palette.setColor(role, QColor(color))
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        palette.setColor(QPalette.Disabled, role, QColor("#888888"))
    return palette


class ApplicationTheme(QObject):
    """Keep an explicit application theme separate from the system taskbar theme."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.system_palette = QPalette(app.palette())
        self.mode = "system"
        self._applying = False
        app.paletteChanged.connect(self._system_palette_changed)
        QGuiApplication.styleHints().colorSchemeChanged.connect(self._system_scheme_changed)

    def set_mode(self, mode):
        if mode != self.mode:
            self.mode = mode
            self._apply()

    def _system_palette_changed(self, palette):
        if not self._applying:
            self.system_palette = QPalette(palette)
            if self.mode != "system":
                self._apply()

    def _system_scheme_changed(self, *_args):
        self._applying = True
        try:
            self.app.setPalette(QPalette())
            self.system_palette = QPalette(self.app.palette())
        finally:
            self._applying = False
        self._apply()

    def _apply(self, *_args):
        if self._applying:
            return
        self._applying = True
        try:
            palette = self.system_palette if self.mode == "system" else theme_palette(self.mode == "dark", self.system_palette)
            self.app.setPalette(palette)
        finally:
            self._applying = False


def accent_rgba(alpha: float) -> str:
    color = accent_color()
    return f"rgba({color.red()}, {color.green()}, {color.blue()}, {alpha:.2f})"


COLOR_SWATCH_SIZE = 12
LINUX_FONT_RULES = """
QWidget { font-size: 12px; }
"""

_SETTINGS_ICON_PATHS = (
    'M8 5h11M8 12h11M8 19h11M4 5h.1M4 12h.1M4 19h.1',
    'M4 19V5m0 14h16M8 15l4-5 4 2 4-6',
    'M4 6h16M4 12h16M4 18h16M8 3v6m8 0v6m-7 0v6',
    'M4 5h16v14H4zM7 9h.1m4 0h.1m4 0h.1M7 12h.1m4 0h.1m4 0h.1M8 16h8',
    'M3 4h18v16H3zM3 8h18M7 16l9-5m-5 0h5v5',
    'M3 4h18v16H3zM3 15h18M7 17.5h2m2 0h2m2 0h2',
    'M12 3a9 9 0 1 0 0 18 9 9 0 1 0 0-18M12 11v6m0-10h.1',
)


def settings_navigation_icon(index: int, device_pixel_ratio: float = 1.0, palette=None) -> QIcon:
    """Render the sidebar's line icons with the settings text color."""
    color = (palette or QApplication.palette()).color(QPalette.ColorRole.WindowText).name()
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">'
           f'<path d="{_SETTINGS_ICON_PATHS[index]}" fill="none" stroke="{color}" '
           'stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>')
    ratio = max(1.0, float(device_pixel_ratio))
    pixmap = QPixmap(round(20 * ratio), round(20 * ratio))
    pixmap.setDevicePixelRatio(ratio)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    QSvgRenderer(QByteArray(svg.encode())).render(painter, QRectF(0, 0, 20, 20))
    painter.end()
    return QIcon(pixmap)


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


def set_color_button(button, color: QColor, title: str):
    name = color.name(QColor.NameFormat.HexRgb).upper()
    button.setText("  " + name)
    button.setToolTip(f"{title}: {name}")
    button.setAccessibleName(f"{title}: {name}")
    button.setIcon(color_swatch_icon(color, button.devicePixelRatioF()))


def build_settings_stylesheet(dark: bool, *, linux_fonts: bool = False, palette: QPalette | None = None) -> str:
    """按系统深浅色生成设置窗口样式表"""
    # 在设置窗口的样式表中直接匹配每个子控件，覆盖 Linux 桌面主题的字号。
    # 仅对父窗口 setFont 无法覆盖子控件的样式字体；后创建的编辑器也要匹配。
    # 保留字体家族与字重，Qt 继续按屏幕像素比缩放这些逻辑像素。
    font_rules = LINUX_FONT_RULES if linux_fonts else ""
    foreground = (palette or QApplication.palette()).color(QPalette.WindowText).name()
    sidebar_bg = "rgba(255, 255, 255, 0.04)" if dark else "rgba(0, 0, 0, 0.03)"
    nav_hover = "rgba(255, 255, 255, 0.06)" if dark else "rgba(0, 0, 0, 0.05)"
    nav_selected = "rgba(255, 255, 255, 0.12)" if dark else "rgba(0, 0, 0, 0.09)"
    scrollbar = "rgba(255, 255, 255, 0.30)" if dark else "rgba(0, 0, 0, 0.25)"
    description = "rgba(255, 255, 255, 0.55)" if dark else "rgba(0, 0, 0, 0.52)"
    slider_track = "rgba(255, 255, 255, 0.20)" if dark else "rgba(0, 0, 0, 0.18)"
    if dark:
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
QFrame#settings_sidebar {{
    background-color: {sidebar_bg};
    border: none;
}}
QLabel#settings_sidebar_title {{
    font-size: 16px;
    font-weight: 600;
    padding-left: 8px;
}}
QListWidget#settings_navigation {{
    background: transparent;
    border: none;
    outline: none;
    font-size: 13px;
}}
QListWidget#settings_navigation::item {{
    padding: 3px 8px;
    border-radius: 6px;
    color: {foreground};
}}
QListWidget#settings_navigation::item:hover {{
    background-color: {nav_hover};
}}
QListWidget#settings_navigation::item:selected {{
    background-color: {nav_selected};
    color: {foreground};
}}
QLabel#watchlist_title, QLabel#data_title, QLabel#general_title,
QLabel#shortcuts_title, QLabel#floating_title, QLabel#taskbar_title, QLabel#about_title {{
    font-size: 20px;
    font-weight: 600;
}}
QLabel[settingTitle="true"] {{
    font-weight: 600;
}}
QLabel[settingDescription="true"] {{
    color: {description};
    font-size: 12px;
}}
QPushButton[destructive="true"] {{
    background-color: rgba(220, 60, 60, 0.16);
}}
QPushButton[destructive="true"]:hover {{
    background-color: rgba(220, 60, 60, 0.24);
}}
QSlider::groove:horizontal {{
    height: 4px;
    border-radius: 2px;
    background: {slider_track};
}}
QSlider::sub-page:horizontal {{
    background: {accent_color().name()};
    border-radius: 2px;
}}
QSlider::handle:horizontal {{
    width: 14px;
    margin: -5px 0;
    border-radius: 7px;
    background: {accent_color().name()};
}}
QSlider::handle:horizontal:disabled, QSlider::sub-page:horizontal:disabled {{
    background: palette(mid);
}}
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollArea QScrollBar:vertical {{
    width: 8px;
    background: transparent;
    margin: 2px;
}}
QScrollArea QScrollBar::handle:vertical {{
    background: {scrollbar};
    border-radius: 2px;
    min-height: 24px;
}}
QScrollArea QScrollBar::add-line:vertical, QScrollArea QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollArea QScrollBar::add-page:vertical, QScrollArea QScrollBar::sub-page:vertical {{
    background: transparent;
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


def apply_settings_theme(widget: QWidget, dark: bool, palette: QPalette, *,
                         linux_fonts: bool = False, extra_stylesheet: str = "") -> None:
    """Apply shared dialog styles and palettes to native and styled controls."""
    widgets = (widget, *widget.findChildren(QWidget))
    # CSS palette references need the selected palette before polish. Native
    # controls can be reset by polish, so restore their palettes afterwards.
    for control in widgets:
        control.setPalette(palette)
    widget.setStyleSheet(build_settings_stylesheet(dark, linux_fonts=linux_fonts, palette=palette)
                        + "\n" + extra_stylesheet)
    for control in widgets:
        control.setPalette(palette)
