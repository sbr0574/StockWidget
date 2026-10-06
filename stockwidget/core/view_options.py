"""Validated view preferences and paging math, independent of Qt."""

from dataclasses import asdict, dataclass, field

from stockwidget.core.quote_presentation import DEFAULT_VISIBLE_METRICS, normalize_visible_metrics


PAGE_MODES = (("first", "只显示前 N 行"), ("manual", "手动翻页"), ("auto", "自动翻页"))
TASKBAR_STYLE_KEYS = (
    "taskbar_sync_appearance", "taskbar_font_family", "taskbar_font_size",
    "taskbar_color", "taskbar_auto_color", "taskbar_opacity_pct", "taskbar_unicolor",
)
APPEARANCE_OPTION_KEYS = ("color_mode", *TASKBAR_STYLE_KEYS)


def bounded_int(value, default, low, high):
    try:
        return max(low, min(high, int(value)))
    except (ValueError, TypeError, OverflowError):
        return default


def taskbar_content_height(height, dpi):
    return min(height, round(44 * dpi / 96))


def taskbar_font_pixels(height, dpi, rows, point_size):
    row_height = max(1, (height - 4) // rows)
    return max(7, min(round(point_size * dpi / 72), row_height - 2))


def taskbar_font_size_limit(height, dpi, rows):
    """Expose the renderer's row/DPI limit in whole point sizes."""
    pixels = max(7, (height - 4) // rows - 2)
    return max(5, min(30, round(pixels * 72 / dpi)))


@dataclass
class ViewOptions:
    chart_enabled: bool = False
    chart_display_mode: str = "window"
    color_mode: str = "system"
    hide_tray_icon: bool = False
    float_split_enabled: bool = False
    float_split_separator: bool = True
    float_paging_enabled: bool = False
    float_max_rows: int = 3
    float_page_mode: str = "manual"
    float_page_interval: int = 5
    taskbar_rows: int = 2
    taskbar_page_mode: str = "first"
    taskbar_page_interval: int = 5
    taskbar_sync_paging: bool = True
    taskbar_sync_split: bool = True
    taskbar_split_enabled: bool = False
    taskbar_split_separator: bool = True
    taskbar_metrics: list[str] = field(default_factory=lambda: list(DEFAULT_VISIBLE_METRICS))
    taskbar_sync_metrics: bool = True
    taskbar_enabled: bool = False
    taskbar_sync_appearance: bool = True
    taskbar_font_family: str = "Microsoft YaHei"
    taskbar_font_size: int = 10
    taskbar_color: str = "#FFFFFF"
    taskbar_auto_color: bool = False
    taskbar_opacity_pct: int = 100
    taskbar_unicolor: bool = True
    taskbar_dual_open: bool = False

    @classmethod
    def from_config(cls, cfg, font_family="Microsoft YaHei", metrics=None):
        defaults = cls(taskbar_font_family=font_family)
        result = cls(**{key: cfg.get(key, value) for key, value in asdict(defaults).items()})
        for key, low, high in (
            ("float_max_rows", 1, 20), ("float_page_interval", 1, 60),
            ("taskbar_rows", 1, 4), ("taskbar_page_interval", 1, 3600),
            ("taskbar_font_size", 5, 30), ("taskbar_opacity_pct", 0, 100),
        ):
            setattr(result, key, bounded_int(getattr(result, key), getattr(defaults, key), low, high))
        for key in ("float_page_mode", "taskbar_page_mode"):
            if getattr(result, key) not in ("first", "manual", "auto"):
                setattr(result, key, getattr(defaults, key))
        if result.color_mode not in ("system", "light", "dark"):
            result.color_mode = "system"
        if result.chart_display_mode not in ("window", "floating"):
            result.chart_display_mode = "window"
        if "taskbar_metrics" not in cfg and metrics is not None:
            result.taskbar_metrics = metrics
        result.taskbar_metrics = normalize_visible_metrics(result.taskbar_metrics)
        if "float_paging_enabled" not in cfg:
            result.float_paging_enabled = bounded_int(cfg.get("float_max_rows", 0), 0, 0, 20) > 0
        if "taskbar_enabled" not in cfg:
            result.taskbar_enabled = cfg.get("display_mode", "float") in ("taskbar", "both")
        if "taskbar_sync_metrics" not in cfg:
            result.taskbar_sync_metrics = "taskbar_metrics" not in cfg
        if "taskbar_sync_appearance" not in cfg:
            result.taskbar_sync_appearance = cfg.get("taskbar_sync_font", True) and cfg.get("taskbar_sync_color", True)
        if "taskbar_sync_paging" not in cfg:
            result.taskbar_sync_paging = "taskbar_page_mode" not in cfg
        if not isinstance(result.taskbar_font_family, str) or not result.taskbar_font_family.strip():
            result.taskbar_font_family = font_family
        if not isinstance(result.taskbar_color, str):
            result.taskbar_color = defaults.taskbar_color
        for key in ("chart_enabled", "hide_tray_icon", "float_paging_enabled", "taskbar_enabled", "taskbar_sync_metrics",
                    "taskbar_sync_appearance", "taskbar_sync_paging", "taskbar_unicolor", "taskbar_auto_color", "taskbar_dual_open",
                    "float_split_enabled", "float_split_separator", "taskbar_sync_split",
                    "taskbar_split_enabled", "taskbar_split_separator"):
            setattr(result, key, bool(getattr(result, key)))
        if "taskbar_dual_open" not in cfg:
            result.taskbar_dual_open = cfg.get("display_mode") == "both"
        return result

    def to_config(self):
        return asdict(self)

    def page_settings(self, surface):
        """读取指定显示面的有效翻页配置，保留未启用的独立偏好。"""
        if surface == "float" or self.taskbar_sync_paging:
            return (self.float_page_mode if self.float_paging_enabled else "first",
                    self.float_page_interval)
        return self.taskbar_page_mode, self.taskbar_page_interval

    def split_settings(self, surface):
        prefix = "float" if surface == "float" or self.taskbar_sync_split else "taskbar"
        return getattr(self, f"{prefix}_split_enabled"), getattr(self, f"{prefix}_split_separator")


@dataclass(frozen=True)
class Page:
    index: int
    count: int
    start: int
    stop: int
    controls: bool


def page_slice(total, limit, mode, index=0):
    total = max(0, total)
    if limit <= 0:
        return Page(0, 1, 0, total, False)
    count = max(1, (total + limit - 1) // limit)
    index = 0 if mode == "first" else max(0, min(index, count - 1))
    start = index * limit
    return Page(index, count, start, min(total, start + limit), mode != "first" and count > 1)


def column_ranges(total, split):
    """Split the current page in reading order; an odd extra item stays left."""
    midpoint = (total + 1) // 2 if split else total
    return ((0, midpoint), (midpoint, total)) if split else ((0, total),)
