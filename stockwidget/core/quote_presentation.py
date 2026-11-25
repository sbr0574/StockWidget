"""指标定义、旧配置迁移、数值格式化及行情展示投影；不依赖 Qt。"""

from dataclasses import dataclass
import math
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
    MetricSpec("average", "均值", "均值", "avg_visible"),
    MetricSpec("volume_ratio", "量比", "量比", "volume_ratio_visible"),
    # Keep the stored ID / legacy visibility flag so existing layouts migrate intact.
    MetricSpec("kline", "迷你图", "迷你图", "kline_visible"),
)

METRIC_BY_ID = {spec.metric_id: spec for spec in METRIC_SPECS}
METRIC_IDS = tuple(METRIC_BY_ID)
DEFAULT_VISIBLE_METRICS = tuple(
    spec.metric_id for spec in METRIC_SPECS if spec.default_visible
)


def quote_detail_values(quote, instrument, details=None):
    """Numeric chart-header metrics; unavailable ratios remain None.

    Domestic earnings multiples use current market cap and the matching profit,
    never weighted historical EPS. US TTM EPS comes from its basic quote.
    """
    data = {**(quote or {}), **(details or {})}

    def number(key):
        value = data.get(key)
        return value if isinstance(value, (int, float)) and math.isfinite(value) else None

    def ratio(numerator, denominator):
        return numerator / denominator if numerator is not None and denominator is not None and denominator > 0 else None

    price = number("current_price") or number("prev_close")
    total, floating = number("total_shares"), number("float_shares")
    total = total if total and total > 0 else None
    floating = floating if floating and floating > 0 else None
    volume = number("deals_vol")
    result = {key: number(key) for key in ("inside_volume", "outside_volume", "total_shares", "float_shares")}
    result["total_shares"], result["float_shares"] = total, floating
    capitalization = price * total if price and total else (number("market_cap") or None)
    result.update({
        "pe_ttm": (ratio(price, number("eps_ttm")) if instrument.get("market") == "us"
                   else ratio(capitalization, number("ttm_profit"))),
        "pe_static": ratio(capitalization, number("annual_profit")),
        "pe_dynamic": ratio(capitalization, number("annualized_profit")),
        "pb": ratio(price, number("book_value_per_share")),
        "market_cap": capitalization,
        "float_market_cap": price * floating if price and floating and floating > 0 else None,
        "turnover_rate": ratio(volume * 100 if volume is not None else None, floating),
        "amplitude": ratio(((number("high_price") or 0) - (number("low_price") or 0)) * 100,
                           number("prev_close")) if number("opening_price") else None,
    })
    if volume and not (result["inside_volume"] or result["outside_volume"]):
        result["inside_volume"] = result["outside_volume"] = None
    if instrument.get("type") in {"指", "基", "期"}:
        for key in ("pe_ttm", "pe_static", "pe_dynamic", "pb", "market_cap", "total_shares",
                    "float_market_cap", "float_shares", "turnover_rate", "inside_volume", "outside_volume"):
            result[key] = None
    return result


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


def format_signed_value(value: float, lot_size: int = 1, unit_cn: bool = True) -> str:
    """带正负号格式化成交量/成交额等增减差值。"""
    if not isinstance(value, (int, float)) or not math.isfinite(value):
        return "—"
    magnitude = abs(value)
    if int(magnitude / lot_size) == 0:
        return "0"
    sign = "+" if value > 0 else "-"
    return sign + format_value(magnitude, lot_size=lot_size, unit_cn=unit_cn)


def volume_lot_size(security_type: str | None, market: str = "") -> int:
    return 100 if market in {"sh", "sz", "bj"} and security_type != "期" else 1


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

CN_INDEX_MARKETS = frozenset({"sh", "sz", "bj"})
MARKET_BREADTH_BUCKETS = (
    (">10%", COLOR_ROLE_UP),
    ("7~10%", COLOR_ROLE_UP),
    ("5~7%", COLOR_ROLE_UP),
    ("2~5%", COLOR_ROLE_UP),
    ("0~2%", COLOR_ROLE_UP),
    ("0%", COLOR_ROLE_NEUTRAL),
    ("0~-2%", COLOR_ROLE_DOWN),
    ("-2~-5%", COLOR_ROLE_DOWN),
    ("-5~-7%", COLOR_ROLE_DOWN),
    ("-7~-10%", COLOR_ROLE_DOWN),
    ("<-10%", COLOR_ROLE_DOWN),
)

MARKET_OVERVIEW_INDICES = (
    ("sh000001", {"code": "000001", "market": "sh", "type": "指", "name": "上证指数"}),
    ("sz399001", {"code": "399001", "market": "sz", "type": "指", "name": "深证成指"}),
    ("sh000300", {"code": "000300", "market": "sh", "type": "指", "name": "沪深300"}),
    ("sh000680", {"code": "000680", "market": "sh", "type": "指", "name": "科创综指"}),
    ("sz399006", {"code": "399006", "market": "sz", "type": "指", "name": "创业板指"}),
    ("bj899050", {"code": "899050", "market": "bj", "type": "指", "name": "北证50"}),
)

# 沪深京全市场成交额口径：上证指数(沪市全市场) + 深证综指(深市全市场) + 北证50(北交所全市场)。
# 注意：深证成指(sz399001)的分钟K线仅统计成分股成交额，因此统一以深证综指(sz399106)对齐实时与历史分钟基准。
MARKET_TURNOVER_INDICES = (
    ("sh000001", {"code": "000001", "market": "sh", "type": "指", "name": "上证指数"}),
    ("sz399106", {"code": "399106", "market": "sz", "type": "指", "name": "深证综指"}),
    ("bj899050", {"code": "899050", "market": "bj", "type": "指", "name": "北证50"}),
)

MARKET_OVERVIEW_QUOTES = {
    key: dict(instrument)
    for key, instrument in (*MARKET_OVERVIEW_INDICES, *MARKET_TURNOVER_INDICES)
}


def is_cn_index(instrument) -> bool:
    """Return True for domestic indices listed in index_cn.json."""
    if not isinstance(instrument, Mapping):
        return False
    return (
        str(instrument.get("market", "") or "").strip().lower() in CN_INDEX_MARKETS
        and str(instrument.get("type", "") or "").strip() == "指"
    )


def format_overview_index(instrument: Mapping, quote: Mapping | None) -> dict:
    """Format one index card for the market overview panel."""
    name = str(instrument.get("name") or (quote or {}).get("name") or instrument.get("code") or "")
    if not isinstance(quote, Mapping):
        return {"name": name, "price": "—", "change": "—", "change_pct": "—", "role": COLOR_ROLE_TEXT}
    precision = price_precision(instrument.get("type"), instrument.get("market", ""))
    try:
        current = float(quote.get("current_price") or 0)
        prev_close = float(quote.get("prev_close") or 0)
    except (TypeError, ValueError):
        return {"name": name, "price": "—", "change": "—", "change_pct": "—", "role": COLOR_ROLE_TEXT}
    if not (math.isfinite(current) and math.isfinite(prev_close)):
        return {"name": name, "price": "—", "change": "—", "change_pct": "—", "role": COLOR_ROLE_TEXT}
    if current == 0 and prev_close > 0:
        current = prev_close
    if current <= 0:
        return {"name": name, "price": "—", "change": "—", "change_pct": "—", "role": COLOR_ROLE_TEXT}
    price_text = f"{current:.{precision}f}"
    if prev_close <= 0:
        return {"name": name, "price": price_text, "change": "—", "change_pct": "—", "role": COLOR_ROLE_NEUTRAL}
    change = current - prev_close
    change_pct = (current / prev_close - 1) * 100
    return {
        "name": name,
        "price": price_text,
        "change": f"{change:+.{precision}f}",
        "change_pct": f"{change_pct:+.2f}%",
        "role": direction_color_role(change),
    }


def _parse_quote_date_clock(quote: Mapping | None):
    if not isinstance(quote, Mapping):
        return None, None
    raw_date = str(quote.get("date") or "").strip().replace("/", "-")
    raw_time = str(quote.get("time") or "").strip()
    if len(raw_date) != 10 or raw_date[4] != "-" or raw_date[7] != "-":
        return None, None
    try:
        year, month, day = int(raw_date[:4]), int(raw_date[5:7]), int(raw_date[8:10])
        if not (1900 <= year <= 2200 and 1 <= month <= 12 and 1 <= day <= 31):
            return None, None
    except ValueError:
        return None, None
    clock = None
    if len(raw_time) >= 8 and raw_time[2] == ":" and raw_time[5] == ":":
        try:
            hh, mm, ss = int(raw_time[:2]), int(raw_time[3:5]), int(raw_time[6:8])
            if 0 <= hh <= 23 and 0 <= mm <= 59 and 0 <= ss <= 59:
                clock = (hh, mm, ss)
        except ValueError:
            clock = None
    return raw_date, clock


def _bar_minute(bar) -> int | None:
    stamp = str(getattr(bar, "time", "") or "")
    if len(stamp) < 16 or stamp[13] != ":":
        return None
    try:
        hh, mm = int(stamp[11:13]), int(stamp[14:16])
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return hh * 60 + mm
    except ValueError:
        return None
    return None


def _cumulative_amount_at_clock(day_bars: Sequence[tuple[int, float]], clock) -> float | None:
    """Cumulative turnover on a completed CN trading day at wall clock (hh, mm, ss)."""
    if clock is None or not day_bars:
        return None
    hh, mm, ss = clock
    minutes = hh * 60 + mm + ss / 60.0
    if minutes <= 570.0:  # At or before 09:30:00 continuous open.
        return None
    if minutes >= 900.0:  # At or after 15:00:00 close.
        return sum(amount for _, amount in day_bars)
    if 690.0 <= minutes <= 780.0:  # Lunch break 11:30:00 .. 13:00:00.
        return sum(amount for bar_min, amount in day_bars if bar_min <= 690)
    whole = int(minutes)
    frac = minutes - whole
    total = sum(amount for bar_min, amount in day_bars if bar_min <= whole)
    if frac > 0:
        next_minute = whole + 1
        # Only interpolate inside an active continuous minute; never bridge breaks or closing auction.
        if (571 <= next_minute <= 690) or (781 <= next_minute <= 898):
            for bar_min, amount in day_bars:
                if bar_min == next_minute:
                    total += amount * frac
                    break
    return total


def compute_market_turnover(
    overview_quotes: Mapping | None,
    turnover_bars: Mapping | None = None,
) -> dict:
    """Compute full A-share (SH + SZ + BJ) live turnover, yesterday total, and same-time delta.

    Requires complete, same-date market quotes across sh000001, sz399106, and bj899050;
    never fabricates missing markets or extrapolates partial history.
    """
    keys = tuple(key for key, _ in MARKET_TURNOVER_INDICES)
    quotes = overview_quotes if isinstance(overview_quotes, Mapping) else {}
    bars_map = turnover_bars if isinstance(turnover_bars, Mapping) else {}

    parsed_live = {}
    for key in keys:
        quote = quotes.get(key)
        date_str, clock = _parse_quote_date_clock(quote)
        if date_str is None:
            continue
        try:
            amt = float(quote.get("deals_amt"))
        except (TypeError, ValueError):
            continue
        if not math.isfinite(amt) or amt < 0:
            continue
        parsed_live[key] = (date_str, clock, amt)

    today_date = None
    today_amount = None
    if len(parsed_live) == len(keys):
        dates = {entry[0] for entry in parsed_live.values()}
        if len(dates) == 1:
            today_date = next(iter(dates))
            today_amount = sum(entry[2] for entry in parsed_live.values())

    yesterday_amount = None
    delta_amount = None
    prev_date = None
    if today_date is not None:
        per_market_prev = {}
        for key in keys:
            bars = bars_map.get(key) or ()
            grouped: dict[str, list[tuple[int, float]]] = {}
            invalid_dates: set[str] = set()
            for bar in bars:
                stamp = str(getattr(bar, "time", "") or "")
                if len(stamp) < 16:
                    continue
                bar_date = stamp[:10]
                if bar_date >= today_date:
                    continue
                bar_min = _bar_minute(bar)
                amt = getattr(bar, "amount", None)
                if (bar_min is None or amt is None
                        or not isinstance(amt, (int, float))
                        or not math.isfinite(amt) or amt < 0):
                    invalid_dates.add(bar_date)
                    continue
                grouped.setdefault(bar_date, []).append((bar_min, float(amt)))
            valid_dates = [d for d in grouped if d not in invalid_dates and grouped[d]]
            if not valid_dates:
                break
            latest_prev = max(valid_dates)
            per_market_prev[key] = (latest_prev, grouped[latest_prev])

        if len(per_market_prev) == len(keys):
            prev_dates = {entry[0] for entry in per_market_prev.values()}
            if len(prev_dates) == 1:
                prev_date = next(iter(prev_dates))
                yesterday_amount = sum(
                    sum(amount for _, amount in per_market_prev[key][1])
                    for key in keys
                )
                same_time_parts = []
                for key in keys:
                    part = _cumulative_amount_at_clock(per_market_prev[key][1], parsed_live[key][1])
                    if part is None:
                        same_time_parts = None
                        break
                    same_time_parts.append(part)
                if same_time_parts is not None:
                    delta_amount = today_amount - sum(same_time_parts)

    return {
        "today": today_amount,
        "yesterday": yesterday_amount,
        "delta": delta_amount,
        "date": today_date,
        "prev_date": prev_date,
    }


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


@dataclass(frozen=True)
class MiniChartCell:
    """Shared chart payload; amplitude limits belong to the visible column."""
    kind: str = "kline"
    candle: tuple[float, ...] = ()
    points: tuple[tuple[float, float], ...] = ()
    limits: tuple[float, float] | None = None

    @property
    def reference(self):
        return self.candle[4] if len(self.candle) == 5 else 0

    @property
    def amplitude(self):
        if len(self.candle) != 5 or not all(math.isfinite(v) for v in self.candle):
            return None
        _opening, current, high, low, previous = self.candle
        if min(current, high, low, previous) <= 0:
            return None
        return tuple((v / previous - 1) * 100 for v in (min(low, current), current, max(high, current)))


def amplitude_limits(cells):
    ranges = [cell.amplitude for cell in cells if isinstance(cell, MiniChartCell) and cell.amplitude is not None]
    return (min(v[0] for v in ranges), max(v[2] for v in ranges)) if ranges else None


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
    average_period: int = 0


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
    average_price: float | None = None,
    volume_ratio: float | None = None,
):
    data = dict(data)
    lot_size = volume_lot_size(security_type, market)

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
    avg = (average_price if options.average_period else
           data["deals_amt"] / data["deals_vol"] if data["deals_vol"] > 0 and data["deals_amt"] > 0 else None)
    if avg is not None and not math.isfinite(avg):
        avg = None
    if volume_ratio is not None and (not math.isfinite(volume_ratio) or volume_ratio < 0):
        volume_ratio = None
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
        "均值": f"{avg:.{precision}f}" if avg is not None else "-",
        "量比": f"{volume_ratio:.2f}" if volume_ratio is not None else "-",
        "迷你图": k_payload,
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
        "均值": direction_color_role(avg - data["prev_close"]) if avg is not None else COLOR_ROLE_NEUTRAL,
        "量比": direction_color_role(volume_ratio - 1) if volume_ratio is not None else COLOR_ROLE_NEUTRAL,
        "迷你图": COLOR_ROLE_TEXT,
    }
    sort_values = {
        "现价": data["current_price"],
        "涨跌": change,
        "涨幅": change_pct,
        "浮盈": profit_pct,
        "委比": committee if (p_sum + s_sum) > 0 else None,
        "成交量": data["deals_vol"],
        "成交额": data["deals_amt"],
        "均值": avg,
        "量比": volume_ratio,
    }

    # 指数的成交额/成交量不是指数点位；日K的 MA 点位仍然有效。
    if is_index:
        for key in ("浮盈", "委比", *(("均值",) if not options.average_period else ())):
            format_data[key] = "-"
            color_roles[key] = direction_color_role(0)
        sort_values["浮盈"] = None
        sort_values["委比"] = None
        if not options.average_period:
            sort_values["均值"] = None
        if not data["deals_vol"]:
            sort_values["成交量"] = None
        if not data["deals_amt"]:
            sort_values["成交额"] = None

    return format_data, color_roles, sort_values


SORTABLE_HEADERS = ("现价", "涨跌", "涨幅", "浮盈", "委比", "成交量", "成交额", "均值", "量比")


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
    right_columns = [i for i, h in enumerate(headers) if h not in ("名称", "迷你图", "买一/卖一")]
    return values, roles, right_columns
