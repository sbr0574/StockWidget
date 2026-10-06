"""指标定义、旧配置迁移、数值格式化及行情展示投影；不依赖 Qt。"""

from dataclasses import dataclass
from typing import Mapping, Sequence


@dataclass(frozen=True)
class MetricSpec:
    metric_id: str
    label: str
    header: str
    legacy_attr: str
    default_visible: bool = False


# 名称作为可排序、可显示/隐藏的普通指标参与布局。
NAME_METRIC_ID = "name"

METRIC_SPECS = (
    MetricSpec(NAME_METRIC_ID, "名称", "名称", "name_visible", True),
    MetricSpec("price", "现价", "现价", "price_visible", True),
    MetricSpec("change", "涨跌", "涨跌", "change_visible"),
    MetricSpec("change_pct", "涨幅", "涨幅", "change_pct_visible", True),
    MetricSpec("profit", "浮盈", "浮盈", "profit_visible"),
    MetricSpec("b1s1", "买一/卖一", "买一/卖一", "b1s1_visible"),
    MetricSpec("commi", "委比", "委比", "commi_visible"),
    MetricSpec("volume", "成交量", "成交量", "vol_visible"),
    MetricSpec("amount", "成交额", "成交额", "amount_visible"),
    MetricSpec("average", "均价", "均价", "avg_visible"),
    MetricSpec("kline", "日K线", "K线", "kline_visible"),
)

METRIC_BY_ID = {spec.metric_id: spec for spec in METRIC_SPECS}
METRIC_IDS = tuple(METRIC_BY_ID)
DEFAULT_VISIBLE_METRICS = tuple(
    spec.metric_id for spec in METRIC_SPECS if spec.default_visible
)


def normalize_visible_metrics(value: Sequence[str] | None) -> list[str]:
    """过滤未知项和重复项，同时保留合法指标的输入顺序。"""
    if not isinstance(value, (list, tuple)):
        return []

    normalized = []
    seen = set()
    for raw_metric_id in value:
        metric_id = str(raw_metric_id or "").strip()
        if metric_id in METRIC_BY_ID and metric_id not in seen:
            normalized.append(metric_id)
            seen.add(metric_id)
    return normalized


def visible_metrics_from_config(cfg: Mapping | None) -> list[str]:
    """读取新有序列表；字段缺失或类型错误时从旧可见布尔值迁移。"""
    cfg = cfg if isinstance(cfg, Mapping) else {}
    raw_metrics = cfg.get("visible_metrics")
    if isinstance(raw_metrics, (list, tuple)):
        normalized = normalize_visible_metrics(raw_metrics)
        # 旧版本配置的 visible_metrics 不含名称指标：名称仍受独立的
        # name_visible 开关控制，未关闭时补在首位，保持原有显示效果。
        if NAME_METRIC_ID not in normalized and bool(cfg.get("name_visible", True)):
            normalized.insert(0, NAME_METRIC_ID)
        return normalized

    return [
        spec.metric_id
        for spec in METRIC_SPECS
        if bool(cfg.get(spec.legacy_attr, spec.default_visible))
    ]


def metric_headers(metric_ids: Sequence[str] | None) -> list[str]:
    """每个指标对应一列，按有效指标的配置顺序取得表头。"""
    return [METRIC_BY_ID[metric_id].header for metric_id in normalize_visible_metrics(metric_ids)]


def legacy_visibility(metric_ids: Sequence[str] | None) -> dict[str, bool]:
    """仅在保存配置时生成旧布尔字段；运行时以有序指标列表为准。"""
    visible = set(normalize_visible_metrics(metric_ids))
    return {
        spec.legacy_attr: spec.metric_id in visible
        for spec in METRIC_SPECS
    }


# 自动单位模式下使用英文单位（k/M/B/T）的市场：美股与国际指数。
ENGLISH_UNIT_MARKETS = frozenset({"us", "gb"})


def format_value(value: float, lot_size: int = 1, unit_cn: bool = True) -> str:
    """统一格式化成交量/成交额等数值。

    unit_cn=True 时按 万/亿/万亿 缩写，否则按 k/M/B/T 缩写；
    lot_size 用于把成交量按每手股数换算后再缩写。
    """
    value = int(value / lot_size)
    if unit_cn:
        if value < 1e4:
            return f"{value}"
        if value < 1e8:
            return f"{value / 1e4:.2f}万"
        if value < 1e12:
            return f"{value / 1e8:.2f}亿"
        return f"{value / 1e12:.2f}万亿"
    if value < 1e3:
        return f"{value}"
    if value < 1e6:
        return f"{value / 1e3:.2f}k"
    if value < 1e9:
        return f"{value / 1e6:.2f}M"
    if value < 1e12:
        return f"{value / 1e9:.2f}B"
    return f"{value / 1e12:.2f}T"


def should_use_english_units(unit_mode: str, market: str = "") -> bool:
    """按单位模式决定是否使用英文单位。

    模式为 中文/英文 时按模式返回；“自动”时美股与国际指数
    （market 为 us/gb）使用英文单位，国内、港股等使用中文单位。
    """
    mode = str(unit_mode or "").strip().lower()
    if mode in ("en", "english", "英文"):
        return True
    if mode in ("cn", "zh", "chinese", "中文"):
        return False
    return str(market or "").strip().lower() in ENGLISH_UNIT_MARKETS


COLOR_ROLE_TEXT = "text"
COLOR_ROLE_UP = "up"
COLOR_ROLE_DOWN = "down"
COLOR_ROLE_NEUTRAL = "neutral"


@dataclass(frozen=True)
class BidAskCell:
    """One level-one metric, with independently colored values around its center."""
    buy: str
    sell: str
    buy_role: str = COLOR_ROLE_NEUTRAL
    sell_role: str = COLOR_ROLE_NEUTRAL
    marker: str = ""

    def __str__(self):
        return f"{self.buy} {self.marker or ' '} {self.sell}"


def _format_book_volume(value, lot_size, *, signed=False):
    """盘口数量按市场换算，超过一万时固定使用两位小数的万单位。"""
    volume = int(value / lot_size)
    if abs(volume) > 10000:
        return f"{volume / 10000:+.2f}万" if signed else f"{volume / 10000:.2f}万"
    return f"{volume:+d}" if signed else str(volume)


def direction_color_role(value) -> str:
    if value > 0:
        return COLOR_ROLE_UP
    if value < 0:
        return COLOR_ROLE_DOWN
    return COLOR_ROLE_NEUTRAL


@dataclass(frozen=True)
class QuoteDisplayOptions:
    name_length: int = -1
    code_visible: bool = False
    type_visible: bool = False
    unit_mode: str = "auto"


def price_precision(security_type, market=""):
    """Shared price precision for quotes and historical charts."""
    return 3 if security_type == "基" or market == "us" else 2


def format_quote(
    data: dict,
    security_type: str | None,
    display_code: str,
    *,
    market: str = "",
    cost: float | None = None,
    options: QuoteDisplayOptions = QuoteDisplayOptions(),
):
    data = dict(data)
    lot_size = 100 if market in {"sh", "sz", "bj"} and security_type != "期" else 1

    # 名称显示
    name = f"({security_type})" if security_type is not None and options.type_visible else ""
    name += f"{display_code} " if options.code_visible else ""
    if options.name_length == -1:
        name += data["name"]
    else:
        name += data["name"][:options.name_length]

    # 一档盘口数据
    pur_1 = data["purchaser_price"][0]
    sell_1 = data["seller_price"][0]
    marker = ""
    if pur_1 == sell_1 > 0:
        # 集合竞价阶段
        data["current_price"] = sell_1
        paired = int(data["seller_vol"][0] / lot_size)
        unpaired = int(
            (data["purchaser_vol"][1] or (-data["seller_vol"][1]))
            / lot_size
        )
        b1_label = _format_book_volume(paired, 1)
        s1_label = _format_book_volume(unpaired, 1, signed=True)
        b1_color_sign = (unpaired > 0) - (unpaired < 0)
        s1_color_sign = b1_color_sign
    else:
        # 连续交易阶段（有买/卖盘口量时才显示，否则"-"）
        pur_v1 = data["purchaser_vol"][0]
        sell_v1 = data["seller_vol"][0]
        if pur_1 and pur_v1 and data["current_price"] == pur_1:
            marker = "◀"
        elif sell_1 and sell_v1 and data["current_price"] == sell_1:
            marker = "▶"
        b1_label = _format_book_volume(pur_v1, lot_size) if (pur_1 and pur_v1) else "-"
        s1_label = _format_book_volume(sell_v1, lot_size) if (sell_1 and sell_v1) else "-"
        b1_color_sign = 1 if (pur_1 and pur_v1) else 0
        s1_color_sign = -1 if (sell_1 and sell_v1) else 0

    # 盘前数据填充
    if data["current_price"] == 0:
        data["current_price"] = data["prev_close"]
    if data["opening_price"] == 0:
        data["opening_price"] = data["current_price"]
        data["high_price"] = data["current_price"]
        data["low_price"] = data["current_price"]

    # 指标计算
    change = data["current_price"] - data["prev_close"] if data["prev_close"] else 0.0
    change_pct = (data["current_price"] / data["prev_close"] - 1) * 100 if data["prev_close"] else 0.0
    avg = (data["deals_amt"] / data["deals_vol"]) if data["deals_vol"] > 0 else data["prev_close"]
    p_sum, s_sum = sum(data["purchaser_vol"]), sum(data["seller_vol"])
    committee = (100 * (p_sum - s_sum) / (p_sum + s_sum)) if (p_sum + s_sum) > 0 else 0.0
    arrow = " "
    if data["high_price"] > data["low_price"]:
        if data["current_price"] == data["high_price"]:
            arrow = "↑"
        elif data["current_price"] == data["low_price"]:
            arrow = "↓"
    k_payload = {"k": (data["opening_price"], data["current_price"], data["high_price"], data["low_price"], data["prev_close"])}

    precision = price_precision(security_type, market)

    # 浮盈计算（与成本价比较），仅显示百分比
    profit_pct = None
    if cost is not None and cost > 0:
        profit_pct = (data["current_price"] / cost - 1) * 100
        profit_label = f"{profit_pct:+.2f}%"
        profit_sign = (profit_pct > 0) - (profit_pct < 0)
    else:
        profit_label = "-"
        profit_sign = 0

    is_index = security_type == "指"
    english_units = should_use_english_units(options.unit_mode, market)
    bid_ask = (BidAskCell("-", "-") if is_index else BidAskCell(
        b1_label, s1_label, direction_color_role(b1_color_sign), direction_color_role(s1_color_sign), marker))
    format_data = {
        "名称": name,
        "现价": f"{data['current_price']:.{precision}f}{arrow}",
        "涨跌": f"{change:+.{precision}f}",
        "涨幅": f"{change_pct:+.2f}%",
        "浮盈": profit_label,
        "买一/卖一": bid_ask,
        "委比": f"{committee:+.2f}%" if (p_sum + s_sum) > 0 else "-",
        "成交量": (
            "-" if is_index and not data["deals_vol"]
            else format_value(
                data["deals_vol"],
                lot_size=lot_size,
                unit_cn=not english_units,
            )
        ),
        "成交额": (
            "-" if is_index and not data["deals_amt"]
            else format_value(
                data["deals_amt"],
                lot_size=1,
                unit_cn=not english_units,
            )
        ),
        "均价": f"{avg:.{precision}f}",
        "K线": k_payload,
    }
    color_roles = {
        "名称": COLOR_ROLE_TEXT,
        "现价": direction_color_role(change),
        "涨跌": direction_color_role(change),
        "涨幅": direction_color_role(change),
        "浮盈": direction_color_role(profit_sign),
        "买一/卖一": COLOR_ROLE_TEXT,
        "委比": direction_color_role(committee),
        "成交量": COLOR_ROLE_TEXT,
        "成交额": COLOR_ROLE_TEXT,
        "均价": direction_color_role(avg - data["prev_close"]),
        "K线": COLOR_ROLE_TEXT,
    }
    sort_values = {
        "现价": data["current_price"],
        "涨跌": change,
        "涨幅": change_pct,
        "浮盈": profit_pct,
        "委比": committee if (p_sum + s_sum) > 0 else None,
        "成交量": data["deals_vol"],
        "成交额": data["deals_amt"],
        "均价": avg,
    }

    # 指数不显示浮盈/买一卖一/委比/均价（均置为"-"）
    if is_index:
        for key in ("浮盈", "委比", "均价"):
            format_data[key] = "-"
            color_roles[key] = direction_color_role(0)
        sort_values["浮盈"] = None
        sort_values["委比"] = None
        sort_values["均价"] = None
        if not data["deals_vol"]:
            sort_values["成交量"] = None
        if not data["deals_amt"]:
            sort_values["成交额"] = None

    return format_data, color_roles, sort_values


SORTABLE_HEADERS = ("现价", "涨跌", "涨幅", "浮盈", "委比", "成交量", "成交额", "均价")


def sorted_quote_indices(sort_values, header, *, descending=True):
    """按原始数值稳定排序，缺失值始终放最后；不修改自选顺序。"""
    indices = list(range(len(sort_values)))
    if header not in SORTABLE_HEADERS:
        return indices
    valid = [i for i in indices if sort_values[i].get(header) is not None]
    missing = [i for i in indices if sort_values[i].get(header) is None]
    valid.sort(key=lambda i: sort_values[i][header], reverse=descending)
    return valid + missing


def project_quote_columns(rows, color_roles, headers):
    """浮窗和任务栏共用相同的缺省文字、颜色及对齐规则。"""
    values = [[row.get(h, "-") for h in headers] for row in rows]
    roles = [[row.get(h, COLOR_ROLE_TEXT) for h in headers] for row in color_roles]
    right_columns = [i for i, h in enumerate(headers) if h not in ("名称", "K线", "买一/卖一")]
    return values, roles, right_columns
