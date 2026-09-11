"""行情指标计算和展示投影，不依赖 Qt，也不修改原始行情。"""

from dataclasses import dataclass

from stockwidget.core.formatters import format_value, should_use_english_units

COLOR_ROLE_TEXT = "text"
COLOR_ROLE_UP = "up"
COLOR_ROLE_DOWN = "down"
COLOR_ROLE_NEUTRAL = "neutral"


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


def format_quote(
    data: dict,
    security_type: str | None,
    display_code: str,
    *,
    market: str = "",
    cost: float | None = None,
    options: QuoteDisplayOptions = QuoteDisplayOptions(),
) -> tuple[dict, dict]:
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
    if pur_1 == sell_1 > 0:
        # 集合竞价阶段
        data["current_price"] = sell_1
        paired = int(data["seller_vol"][0] / lot_size)
        unpaired = int(
            (data["purchaser_vol"][1] or (-data["seller_vol"][1]))
            / lot_size
        )
        b1_label = f"{paired:d}"
        s1_label = f"{unpaired:+d}"
        b1_color_sign = (unpaired > 0) - (unpaired < 0)
        s1_color_sign = b1_color_sign
    else:
        # 连续交易阶段（有买/卖盘口量时才显示，否则"-"）
        pur_v1 = data["purchaser_vol"][0]
        sell_v1 = data["seller_vol"][0]
        buy_marker = "<" if pur_1 and pur_v1 and data["current_price"] == pur_1 else " "
        sell_marker = ">" if sell_1 and sell_v1 and data["current_price"] == sell_1 else " "
        b1_label = f"{int(pur_v1 / lot_size)}{buy_marker}" if (pur_1 and pur_v1) else "-"
        s1_label = f"{sell_marker}{int(sell_v1 / lot_size)}" if (sell_1 and sell_v1) else "-"
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

    precision = 3 if security_type == "基" or market == "us" else 2

    # 浮盈计算（与成本价比较），仅显示百分比
    if cost is not None and cost > 0:
        profit_pct = (data["current_price"] / cost - 1) * 100
        profit_label = f"{profit_pct:+.2f}%"
        profit_sign = (profit_pct > 0) - (profit_pct < 0)
    else:
        profit_label = "-"
        profit_sign = 0

    is_index = security_type == "指"
    english_units = should_use_english_units(options.unit_mode, market)
    format_data = {
        "名称": name,
        "现价": f"{data["current_price"]:.{precision}f}{arrow}",
        "涨跌": f"{change:+.{precision}f}",
        "涨幅": f"{change_pct:+.2f}%",
        "浮盈": profit_label,
        "买一": b1_label,
        "卖一": s1_label,
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
        "买一": direction_color_role(b1_color_sign),
        "卖一": direction_color_role(s1_color_sign),
        "委比": direction_color_role(committee),
        "成交量": COLOR_ROLE_TEXT,
        "成交额": COLOR_ROLE_TEXT,
        "均价": direction_color_role(avg - data["prev_close"]),
        "K线": COLOR_ROLE_TEXT,
    }
    # 指数不显示浮盈/买一卖一/委比/均价（均置为"-"）
    if is_index:
        for key in ("浮盈", "买一", "卖一", "委比", "均价"):
            format_data[key] = "-"
            color_roles[key] = direction_color_role(0)
    return format_data, color_roles
