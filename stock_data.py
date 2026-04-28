from typing import Any

try:
    import akshare as ak
except Exception:
    ak = None


def _format_volume(value: float) -> str:
    if value < 1e4:
        return f"{value:.0f}"
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    return f"{value / 1e8:.2f}亿"


def _format_amount(value: float) -> str:
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    if value < 1e12:
        return f"{value / 1e8:.2f}亿"
    return f"{value / 1e12:.2f}万亿"


def _to_float(v: Any, default: float = 0.0) -> float:
    try:
        text = str(v).replace(",", "").replace("%", "").strip()
        return float(text)
    except Exception:
        return default


def _first_value(row: dict, *keys: str, default: float = 0.0) -> float:
    for key in keys:
        if key in row:
            value = _to_float(row.get(key), default=float("nan"))
            if value == value:
                return value
    return default


def _index_by_code(df, code_keys: tuple[str, ...], normalize) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for _, row in df.iterrows():
        data = row.to_dict()
        raw = ""
        for key in code_keys:
            if key in data and str(data.get(key, "")).strip():
                raw = str(data.get(key)).strip()
                break
        if not raw:
            continue
        code = normalize(raw)
        if code:
            out[code] = data
    return out


def _norm_cn(code: str) -> str:
    text = str(code or "").strip().lower()
    if len(text) == 8 and text[:2] in ("sh", "sz", "bj"):
        return text
    if len(text) == 6 and text.isdigit():
        if text[0] in ("6", "5", "9"):
            return f"sh{text}"
        if text[0] in ("0", "1", "2", "3"):
            return f"sz{text}"
        if text[0] in ("4", "8"):
            return f"bj{text}"
    return text


def _norm_hk(code: str) -> str:
    text = str(code or "").strip().lower()
    if text.startswith("hk"):
        text = text[2:]
    digits = "".join(ch for ch in text if ch.isdigit())
    return f"hk{digits.zfill(5)}" if digits else ""


def _norm_us(code: str) -> str:
    text = str(code or "").strip().lower()
    if text.startswith("us"):
        text = text[2:]
    text = text.lstrip(".").replace(" ", "")
    return f"us{text}" if text else ""


def _fetch_market_snapshots() -> dict[str, dict[str, dict]]:
    if ak is None:
        raise RuntimeError("akshare unavailable")

    cn = _index_by_code(ak.stock_zh_a_spot_em(), ("代码", "code"), _norm_cn)
    fund = _index_by_code(ak.fund_etf_spot_em(), ("代码", "code"), _norm_cn)
    hk = _index_by_code(ak.stock_hk_spot_em(), ("代码", "symbol"), _norm_hk)
    us = _index_by_code(ak.stock_us_spot_em(), ("代码", "symbol"), _norm_us)
    return {"cn": cn, "fund": fund, "hk": hk, "us": us}


def _row_from_quote(code: str, row: dict, short_code: bool, name_length: int):
    name = str(row.get("名称") or row.get("name") or row.get("股票名称") or "").strip()
    category = "沪A" if code.startswith("sh") else "深A"
    precision = 2
    if code.startswith("bj"):
        category = "京A"
    elif code.startswith("hk"):
        category = "港股"
    elif code.startswith("us"):
        category = "美股"
    elif code.startswith(("sh5", "sz1")):
        category = "基金"
        precision = 3

    current_price = _first_value(row, "最新价", "最新", "现价", "最新价(美元)")
    prev_close = _first_value(row, "昨收", "昨收价", "昨收盘")
    opening_price = _first_value(row, "今开", "开盘")
    high_price = _first_value(row, "最高", "最高价")
    low_price = _first_value(row, "最低", "最低价")
    change = _first_value(row, "涨跌额", "涨跌")
    change_pct = _first_value(row, "涨跌幅")
    deals_vol = _first_value(row, "成交量", "成交量(手)")
    deals_amt = _first_value(row, "成交额", "成交金额")

    if prev_close <= 0 and current_price and change:
        prev_close = current_price - change
    if change == 0 and current_price and prev_close:
        change = current_price - prev_close
    if change_pct == 0 and prev_close:
        change_pct = (current_price / prev_close - 1) * 100 if prev_close else 0

    if current_price == 0:
        current_price = prev_close
    if opening_price == 0:
        opening_price = current_price
    if high_price == 0:
        high_price = current_price
    if low_price == 0:
        low_price = current_price

    avg = (deals_amt / deals_vol) if deals_vol > 0 else prev_close
    arrow = " "
    if high_price > low_price:
        if abs(current_price - high_price) < 10 ** (-precision):
            arrow = "↑"
        elif abs(current_price - low_price) < 10 ** (-precision):
            arrow = "↓"

    value_prefix = code[2:] if short_code and code[:2] in ("sh", "sz", "bj", "hk", "us") else code
    name_prefix = name if name_length == 0 else name[:name_length]
    label_name = f"{name_prefix} [{category}]" if category else name_prefix

    return [
        value_prefix,
        label_name,
        f"{current_price:.{precision}f}{arrow}",
        f"{change:+.{precision}f}",
        f"{change_pct:+.2f}%",
        "-",
        "-",
        "-",
        _format_volume(deals_vol),
        _format_amount(deals_amt),
        f"{avg:.{precision}f}",
        {"k": (opening_price, current_price, high_price, low_price, prev_close)},
    ], {
        "delta": (change > 0) - (change < 0),
        "commi": 0,
        "avg": (avg > prev_close) - (avg < prev_close),
        "b1": 0,
        "s1": 0,
    }


def fetch_stock_rows(codes: list[str], short_code: bool, name_length: int, b1s1_display: str):
    _ = b1s1_display
    symbols = [str(c).strip().lower() for c in codes if str(c).strip()]
    if not symbols:
        raise Exception("暂无数据，请添加自选")

    snapshots = _fetch_market_snapshots()
    combined = {}
    combined.update(snapshots["cn"])
    combined.update(snapshots["fund"])
    combined.update(snapshots["hk"])
    combined.update(snapshots["us"])

    price_data = []
    sign_data = []
    for code in symbols:
        row = combined.get(code)
        if not row:
            continue
        one_row, one_sign = _row_from_quote(code, row, short_code, name_length)
        price_data.append(one_row)
        sign_data.append(one_sign)

    if not price_data:
        raise Exception("所选代码暂无可用行情")
    return price_data, sign_data
