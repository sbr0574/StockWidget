from typing import Tuple

import requests

# =====================================================================
# 统一行情字典 schema（新浪 / 东财 两套数据源返回格式一致）：
#   data[code] = {
#       "name": str,  "opening_price": float, "prev_close": float,
#       "current_price": float, "high_price": float, "low_price": float,
#       "deals_vol": int, "deals_amt": float,
#       "purchaser_vol": [买1~买5 量], "purchaser_price": [买1~买5 价],
#       "seller_vol": [卖1~卖5 量], "seller_price": [卖1~卖5 价],
#       "date": str, "time": str,
#   }
#   注：换手率 / 量比 是东财独有字段（新浪没有），已在 _new_entry 中预留（注释），
#       启用东财数据源后取消注释即可填充。
# =====================================================================

# 当前启用的实时行情数据源："sina"（默认）/ "eastmoney"
DATA_SOURCE = "sina"

_SINA_HEADERS = {"Referer": "https://finance.sina.com.cn", "User-Agent": "Mozilla/5.0"}
_EM_HEADERS = {"Referer": "https://quote.eastmoney.com", "User-Agent": "Mozilla/5.0"}
_EM_QUOTE_URL = "https://push2delay.eastmoney.com/api/qt/ulist.np/get"
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
        # 东财独有字段（新浪无），预留
        # "turnover_rate": None,   # 换手率（%）
        # "volume_ratio": None,    # 量比
    }


# ---------------- 统一代码 <-> 数据源代码转换 ----------------
# 统一代码格式：
#   A股    sh600000 / sz000001 / bj920000
#   港股   hk00700
#   美股   usaapl（us + 小写代码）
#   全球指数 gbnky
#   期货   au2512（具体合约）/ au0（主力连续，对应东财 m:113 的 aum）

def _sina_code(instrument: dict) -> str:
    """根据显式 market/code 元数据生成新浪请求代码。"""
    market = str(instrument.get("market", "") or "").strip().lower()
    code = str(instrument.get("code", "") or "").strip().lower()
    if not code:
        return ""
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


# 上期能源（INE）品种代码前缀，东财 secid 市场码为 142（上期所为 113）
_INE_PRODUCTS = {"sc", "nr", "lu", "bc", "ec"}


def _em_secid(instrument: dict) -> str:
    """根据显式 market/code 元数据生成东财 secid。"""
    m = str(instrument.get("market", "") or "").strip().lower()
    c = str(instrument.get("code", "") or "").strip().lower()
    if not c:
        return ""
    if m == "sh":
        return "1." + c
    if m == "sz":
        return "0." + c
    if m == "bj":
        return "0." + c                 # 北交所东财 secid 未验证，暂用深市前缀
    if m == "hk":
        return "116." + c
    if m == "us":
        return "105." + c.upper()       # 105=纳斯达克（106=纽交所/107=美交所）
    if m == "gb":
        return "100." + c              # 全球指数（东财 secid 未验证，暂用 100 前缀）
    if m:
        return ""
    mkt = "142" if c[:2] in _INE_PRODUCTS else "113"  # 上期能源142 / 上期所113
    if len(c) == 3 and c.endswith("0"):
        return f"{mkt}.{c[:-1]}m"      # 主力连续 au0 -> aum
    return f"{mkt}.{c}"                # 具体合约 au2512


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
    """全球指数(b_): 0名称 1最新 2涨跌额 3涨跌幅 5北京时间 6日期 8今开 9昨收 10最高 11最低"""
    return _new_entry(
        name=parts[0],
        opening=parts[8], prev_close=parts[9], current=parts[1],
        high=parts[10], low=parts[11],
        vol=0, amt=0,
        pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
        date=parts[6] if len(parts) > 6 else "",
        time=parts[5] if len(parts) > 5 else "",
    )


def request_sina(instruments: dict[str, dict]) -> dict:
    """新浪财经实时行情。输入和输出均以代码表 key 为索引。"""
    data = {}
    if not instruments:
        return {}
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
    response.encoding = "gbk"
    for line in response.text.split("\n"):
        if not line or '"' not in line:
            continue
        key = line.split('="')[0]
        body = line.split('="')[1]
        if '"' not in body:
            continue
        parts = body.split(",")
        if len(parts) < 3 or "hq_str_" not in key:
            continue
        sname = key.split("hq_str_", 1)[1].strip()
        request_info = requested.get(sname.lower())
        if request_info is None:
            continue
        canonical_key, instrument = request_info
        market = str(instrument.get("market", "") or "").strip().lower()
        is_index = str(instrument.get("type", "") or "").strip() == "指"
        if market == "hk":
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
    return data


# ---------------- 东财解析 ----------------

def request_eastmoney(instruments: dict[str, dict]) -> Tuple[list, dict]:
    """东方财富实时行情（A股/港股/美股/上期所期货），字段与新浪统一。
    东财独有字段（换手率 f8 / 量比 f10）已随请求拉取，字典中预留（注释）。"""
    data = {}
    if not instruments:
        return [], {}
    requests_meta = []
    for key, instrument in instruments.items():
        secid = _em_secid(instrument)
        if secid:
            requests_meta.append((str(key).lower(), instrument, secid))
    if not requests_meta:
        return [], {}
    secids = [item[2] for item in requests_meta]
    params = {
        "secids": ",".join(secids),
        "fields": "f12,f13,f14,f2,f3,f4,f5,f6,f15,f16,f17,f18,f8,f10",
        "fltt": 2,
        "invt": 2,
    }
    response = requests.get(
        _EM_QUOTE_URL,
        params=params,
        headers=_EM_HEADERS,
        timeout=3,
    )
    response.raise_for_status()
    diff = ((response.json() or {}).get("data") or {}).get("diff") or []
    secid_to_request = {s.lower(): (key, instrument) for key, instrument, s in requests_meta}
    raw_to_requests = {}
    for key, instrument, secid in requests_meta:
        raw_to_requests.setdefault(secid.split(".", 1)[1].lower(), []).append((key, instrument))
    for d in diff:
        raw_code = str(d.get("f12", "") or "").strip().lower()
        response_secid = f"{d.get('f13')}.{raw_code}".lower()
        request_info = secid_to_request.get(response_secid)
        if request_info is None:
            candidates = raw_to_requests.get(raw_code, [])
            request_info = candidates[0] if len(candidates) == 1 else None
        if request_info is None:
            continue
        key, instrument = request_info
        vol = _as_float(d.get("f5"))
        if str(instrument.get("market", "") or "").strip().lower() in {"sh", "sz", "bj"}:
            vol = vol * 100   # 东财 A股 f5 单位是“手”，新浪为“股”，统一为股
        entry = _new_entry(
            name=d.get("f14"),
            opening=d.get("f17"), prev_close=d.get("f18"), current=d.get("f2"),
            high=d.get("f15"), low=d.get("f16"),
            vol=vol, amt=d.get("f6"),
            pur_vol=_Z5, pur_price=_Z5,
            sell_vol=_Z5, sell_price=_Z5,
            date="", time="",
        )
        # 东财独有字段（新浪无），预留，取消注释即可填充：
        # entry["turnover_rate"] = d.get("f8")   # 换手率（%）
        # entry["volume_ratio"] = d.get("f10")   # 量比
        data[key] = entry
    return [key in data for key in instruments], data


def request_quote(instruments: dict[str, dict], source: str = DATA_SOURCE) -> dict:
    """统一行情入口，使用 watchlist 中的显式 code/market 元数据。"""
    if source == "eastmoney":
        _, data = request_eastmoney(instruments)
        return data
    return request_sina(instruments)
