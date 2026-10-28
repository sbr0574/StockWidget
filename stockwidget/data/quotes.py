import requests

from stockwidget.data.network_errors import request_error_message


# 新浪行情统一为价格、成交量额、买卖盘和北京时间字段。
_SINA_HEADERS = {"Referer": "https://finance.sina.com.cn", "User-Agent": "Mozilla/5.0"}
_Z5 = (0, 0, 0, 0, 0)   # 空五档（港美股/期货只有一档或无盘口）


def _as_float(value) -> float:
    """把行情数字转为 float；盘前的 '-'、空值等占位符按 0 处理。"""
    try:
        return float(value) if value not in (None, "", "-", "--") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _as_int(value) -> int:
    try:
        return int(_as_float(value))
    except (OverflowError, ValueError):
        return 0

def _new_entry(name, opening, prev_close, current, high, low,
               vol, amt, pur_vol, pur_price, sell_vol, sell_price,
               date, time) -> dict:
    """按统一 schema 构造行情条目。"""
    return {
        "name": str(name or ""),
        "opening_price": _as_float(opening),
        "prev_close": _as_float(prev_close),
        "current_price": _as_float(current),
        "high_price": _as_float(high),
        "low_price": _as_float(low),
        "deals_vol": _as_int(vol),
        "deals_amt": _as_float(amt),
        "purchaser_vol": [_as_int(value) for value in pur_vol],
        "purchaser_price": [_as_float(value) for value in pur_price],
        "seller_vol": [_as_int(value) for value in sell_vol],
        "seller_price": [_as_float(value) for value in sell_price],
        "date": str(date or ""),
        "time": str(time or ""),
    }


# ---------------- 统一代码 <-> 数据源代码转换 ----------------
# 统一代码格式：
#   A股    sh600000 / sz000001 / bj920000
#   港股   hk00700
#   美股   usaapl（us + 小写代码）
#   全球指数 gbnky
#   期货   au2512（具体合约）/ au0（主力连续）

def _sina_code(instrument: dict) -> str:
    """根据显式 market/code 元数据生成新浪请求代码。"""
    market = str(instrument.get("market", "") or "").strip().lower()
    code = str(instrument.get("code", "") or "").strip().lower()
    if not code:
        return ""
    if instrument.get("type") == "期":
        return "nf_" + code.upper()
    if market in {"sh", "sz", "bj"}:
        return market + code
    if market == "hk":
        return "rt_hk" + code.upper()
    if market == "us":
        return "gb_" + code
    if market == "gb":
        return "b_" + code.upper()
    if not market:
        return "nf_" + code.upper()
    return ""


# ---------------- 新浪解析 ----------------

def _parse_sina_a(parts: list, is_index: bool = False, market: str = "") -> dict:
    """A股: 0名称 1今开 2昨收 3最新 4最高 5最低 6买一 7卖一 8量 9额
    10~29 五档(量,价...) 30日期 31时间。指数无五档；上证指数的量为手，
    深证指数的量为股，统一转换为股后写入 schema。"""
    if is_index:
        volume = _as_int(parts[8])
        if market == "sh":
            volume *= 100
        return _new_entry(
            name=parts[0],
            opening=parts[1], prev_close=parts[2], current=parts[3],
            high=parts[4], low=parts[5],
            vol=volume, amt=parts[9],
            pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
            date=parts[30] if len(parts) > 30 else "",
            time=parts[31] if len(parts) > 31 else "",
        )
    return _new_entry(
        name=parts[0],
        opening=parts[1], prev_close=parts[2], current=parts[3],
        high=parts[4], low=parts[5],
        vol=parts[8], amt=parts[9],
        pur_vol=[int(x or 0) for x in parts[10:20:2]],
        pur_price=[float(x or 0) for x in parts[11:20:2]],
        sell_vol=[int(x or 0) for x in parts[20:30:2]],
        sell_price=[float(x or 0) for x in parts[21:30:2]],
        date=parts[30] if len(parts) > 30 else "",
        time=parts[31] if len(parts) > 31 else "",
    )


def _parse_sina_hk(parts: list, is_index: bool = False) -> dict:
    """港股: 0英文名 1中文名 2今开 3昨收 4最高 5最低 6最新
    7涨跌额 8涨跌幅 9买一价 10卖一价 11成交额 12成交量 ... 17日期 18时间
    港股指数无盘口，成交量已是股，成交额为千元(×1000)。"""
    if is_index:
        return _new_entry(
            name=parts[1],
            opening=parts[2], prev_close=parts[3], current=parts[6],
            high=parts[4], low=parts[5],
            vol=parts[12],
            amt=float(parts[11] or 0) * 1000,
            pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
            date=parts[17] if len(parts) > 17 else "",
            time=parts[18] if len(parts) > 18 else "",
        )
    return _new_entry(
        name=parts[1],
        opening=parts[2], prev_close=parts[3], current=parts[6],
        high=parts[4], low=parts[5],
        vol=parts[12], amt=parts[11],
        pur_vol=_Z5, pur_price=[float(parts[9] or 0), 0, 0, 0, 0],
        sell_vol=_Z5, sell_price=[float(parts[10] or 0), 0, 0, 0, 0],
        date=parts[17] if len(parts) > 17 else "",
        time=parts[18] if len(parts) > 18 else "",
    )


def _parse_sina_us(parts: list, is_index: bool = False) -> dict:
    """美股: 0名称 1最新 2涨跌幅 3时间 4涨跌额 5今开 6最高 7最低
    10成交量 ... 26昨收 ... 30成交额；成交量已是股，美股指数无成交额。"""
    t = (parts[3].split() + ["", ""])[:2] if len(parts) > 3 else ["", ""]
    if is_index:
        return _new_entry(
            name=parts[0],
            opening=parts[5], prev_close=parts[26] if len(parts) > 26 else 0,
            current=parts[1], high=parts[6], low=parts[7],
            vol=parts[10],
            amt=0,   # 美股指数无成交额
            pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
            date=t[0], time=t[1],
        )
    return _new_entry(
        name=parts[0],
        opening=parts[5], prev_close=parts[26] if len(parts) > 26 else 0,
        current=parts[1], high=parts[6], low=parts[7],
        vol=parts[10] if len(parts) > 10 else 0,
        amt=parts[30] if len(parts) > 30 else 0,
        pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
        date=t[0], time=t[1],
    )


def _parse_sina_futures(parts: list) -> dict:
    """上期所期货: 0名称 1时间(HHMMSS) 2今开 3最高 4最低 5买价 6卖价
    7最新 8结算价 9昨收(主连为0) 10昨结算 11买量 12卖量 13持仓量
    14成交量 15交易所 16品种 17日期 ... 27均价"""
    tt = str(parts[1] or "")
    time_str = f"{tt[0:2]}:{tt[2:4]}:{tt[4:6]}" if len(tt) >= 6 else ""
    return _new_entry(
        name=parts[0],
        opening=parts[2], prev_close=parts[10] if len(parts) > 10 else 0,
        current=parts[7], high=parts[3], low=parts[4],
        vol=parts[14] if len(parts) > 14 else 0,
        amt=0,   # 新浪期货响应不含成交额
        pur_vol=_Z5, pur_price=[float(parts[5] or 0), 0, 0, 0, 0],
        sell_vol=_Z5, sell_price=[float(parts[6] or 0), 0, 0, 0, 0],
        date=parts[17] if len(parts) > 17 else "",
        time=time_str,
    )


def _parse_sina_global(parts: list) -> dict:
    """全球指数 b_；部分市场第5项是旧日期，第7项才是最新北京时间。"""
    clock = parts[5] if len(parts) > 5 else ""
    if len(clock) != 8 or clock[2:3] != ":" or clock[5:6] != ":":
        clock = parts[7] if len(parts) > 7 else ""
    return _new_entry(
        name=parts[0],
        opening=parts[8], prev_close=parts[9], current=parts[1],
        high=parts[10], low=parts[11],
        vol=0, amt=0,
        pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
        date=parts[6] if len(parts) > 6 else "",
        time=clock,
    )


def request_quote(instruments: dict[str, dict]) -> dict:
    """新浪财经实时行情。输入和输出均以代码表 key 为索引。"""
    data = {}
    if not instruments:
        return {}
    if len(instruments) > 120:
        items = list(instruments.items())
        for offset in range(0, len(items), 120):
            data.update(request_quote(dict(items[offset:offset + 120])))
        return data
    requested = {}
    labels = []
    for key, instrument in instruments.items():
        label = _sina_code(instrument)
        if label:
            requested[label.lower()] = (str(key).lower(), instrument)
            labels.append(label)
    if not requested:
        return {}
    url = "https://hq.sinajs.cn/list=" + ",".join(labels)
    response = requests.get(url, headers=_SINA_HEADERS, timeout=3)
    try:
        response.raise_for_status()
        response.encoding = "gbk"
        text = response.text
    finally:
        response.close()
    matched = False
    for line in text.split("\n"):
        if not line or '"' not in line:
            continue
        key = line.split('="')[0]
        body = line.split('="')[1]
        if '"' not in body:
            continue
        parts = body.split('"', 1)[0].split(",")
        if "hq_str_" not in key:
            continue
        sname = key.split("hq_str_", 1)[1].strip()
        request_info = requested.get(sname.lower())
        if request_info is None:
            continue
        matched = True
        if len(parts) < 3:
            continue
        canonical_key, instrument = request_info
        market = str(instrument.get("market", "") or "").strip().lower()
        is_index = str(instrument.get("type", "") or "").strip() == "指"
        if instrument.get("type") == "期":
            entry = _parse_sina_futures(parts)
        elif market == "hk":
            entry = _parse_sina_hk(parts, is_index=is_index)
        elif market == "us":
            entry = _parse_sina_us(parts, is_index=is_index)
        elif market == "gb":
            entry = _parse_sina_global(parts)
        elif market in {"sh", "sz", "bj"}:
            entry = _parse_sina_a(parts, is_index=is_index, market=market)
        elif not market:
            entry = _parse_sina_futures(parts)
        else:
            continue
        data[canonical_key] = entry
    if not matched:
        raise ValueError("Invalid Sina quote response")
    return data


def fetch_quote_result(codes: dict):
    """Return (ok, quotes, safe_error) for consumers without Qt dependencies."""
    try:
        data = request_quote(codes)
        if codes and not data:
            return False, None, "新浪：未返回有效行情"
        return True, data, None
    except requests.exceptions.RequestException as exc:
        return False, None, f"新浪：{request_error_message(exc)}"
    except Exception as exc:
        return False, None, f"新浪：行情处理失败（{type(exc).__name__}）"
