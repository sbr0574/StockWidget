import requests
from typing import Tuple

from stockwidget.core.markets import market_of

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
_EM_QUOTE_URL = "https://push2delay.eastmoney.com/api/qt/ulist.np/get"   # 东财延迟行情主机（本机代理下可达更稳）
_Z5 = (0, 0, 0, 0, 0)   # 空五档（港美股/期货只有一档或无盘口）

def _new_entry(name, opening, prev_close, current, high, low,
               vol, amt, pur_vol, pur_price, sell_vol, sell_price,
               date, time) -> dict:
    """按统一 schema 构造行情条目。"""
    return {
        "name": str(name or ""),
        "opening_price": float(opening or 0),
        "prev_close": float(prev_close or 0),
        "current_price": float(current or 0),
        "high_price": float(high or 0),
        "low_price": float(low or 0),
        "deals_vol": int(vol or 0),
        "deals_amt": float(amt or 0),
        "purchaser_vol": list(pur_vol),
        "purchaser_price": list(pur_price),
        "seller_vol": list(sell_vol),
        "seller_price": list(sell_price),
        "date": str(date or ""),
        "time": str(time or ""),
        # 东财独有字段（新浪无），预留，取消注释即可启用：
        # "turnover_rate": None,   # 换手率（%）
        # "volume_ratio": None,    # 量比
    }


# ---------------- 统一代码 <-> 数据源代码转换 ----------------
# 统一代码格式：
#   A股    sh600519 / sz000001 / bj430047
#   港股   hk00700
#   美股   usaapl（us + 小写代码）
#   期货   au2512（具体合约）/ au0（主力连续，对应东财 m:113 的 aum）

def _sina_code(code: str) -> str:
    """统一代码 -> 新浪请求代码"""
    c = str(code).strip().lower()
    if c.startswith(("sh", "sz", "bj")):
        return c
    if c.startswith("hk"):
        return "rt_hk" + c[2:].upper()   # rt_hk00700 / rt_hkHSI（代码部分大写，兼容字母代码）
    if c.startswith("us"):
        return "gb_" + c[2:]
    if c.startswith("g") and len(c) > 1:
        return "b_" + c[1:].upper()      # 全球指数 gnky -> b_NKY
    return "nf_" + c.upper()             # 期货: au2512 -> nf_AU2512, au0 -> nf_AU0


def _canonical_from_sina(sname: str) -> str:
    """新浪返回 key（hq_str_ 之后）-> 统一代码（统一为小写）"""
    if sname.startswith("rt_"):
        return sname[3:].lower()         # rt_hk00700 -> hk00700 / rt_hkHSI -> hkhsi
    if sname.startswith("gb_"):
        return "us" + sname[3:].lower()  # gb_aapl -> usaapl
    if sname.startswith("b_"):
        return "g" + sname[2:].lower()   # b_NKY -> gnky
    if sname.startswith("nf_"):
        return sname[3:].lower()          # nf_AU2512 -> au2512
    return sname                          # sh600519


# 上期能源（INE）品种代码前缀，东财 secid 市场码为 142（上期所为 113）
_INE_PRODUCTS = {"sc", "nr", "lu", "bc", "ec"}


def _em_secid(code: str) -> str:
    """统一代码 -> 东财 secid"""
    c = str(code).strip().lower()
    m = market_of(c)
    if m == "sh":
        return "1." + c[2:]
    if m == "sz":
        return "0." + c[2:]
    if m == "bj":
        return "0." + c[2:]            # 北交所东财 secid 未验证，暂用深市前缀
    if m == "hk":
        return "116." + c[2:]
    if m == "us":
        return "105." + c[2:].upper()  # 105=纳斯达克（106=纽交所/107=美交所）
    if m == "g":
        return "100." + c[1:]          # 全球指数（东财 secid 未验证，暂用 100 前缀）
    mkt = "142" if c[:2] in _INE_PRODUCTS else "113"  # 上期能源142 / 上期所113
    if len(c) == 3 and c.endswith("0"):
        return f"{mkt}.{c[:-1]}m"      # 主力连续 au0 -> aum
    return f"{mkt}.{c}"                # 具体合约 au2512


# ---------------- 新浪解析 ----------------

def _is_index_sina(sname: str) -> bool:
    """新浪返回 key（hq_str_ 之后）是否为指数。"""
    if sname.startswith("rt_hk"):
        return not sname[5:].isdigit()                          # 港股指数：字母代码
    if sname.startswith("gb_"):
        return sname[3:].lower() in ("ixic", "dji", "inx", "ndx")  # 美股指数
    if sname.startswith(("sh", "sz")):
        return sname.startswith("sh000") or sname.startswith("sz399")  # A股指数
    return False


def _parse_sina_a(parts: list, is_index: bool = False) -> dict:
    """A股: 0名称 1今开 2昨收 3最新 4最高 5最低 6买一 7卖一 8量 9额
    10~29 五档(量,价...) 30日期 31时间；指数无五档且量单位为手(×100转股)。"""
    if is_index:
        return _new_entry(
            name=parts[0],
            opening=parts[1], prev_close=parts[2], current=parts[3],
            high=parts[4], low=parts[5],
            vol=int(parts[8] or 0) * 100, amt=parts[9],
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
    港股指数无盘口、量单位为手(×100)、额为千元(×1000)。"""
    if is_index:
        return _new_entry(
            name=parts[1],
            opening=parts[2], prev_close=parts[3], current=parts[6],
            high=parts[4], low=parts[5],
            vol=int(parts[12] or 0) * 100,
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
    10成交量 ... 26昨收 ... 30成交额；美股指数量单位为手(×100)、无成交额。"""
    t = (parts[3].split() + ["", ""])[:2] if len(parts) > 3 else ["", ""]
    if is_index:
        return _new_entry(
            name=parts[0],
            opening=parts[5], prev_close=parts[26] if len(parts) > 26 else 0,
            current=parts[1], high=parts[6], low=parts[7],
            vol=int(parts[10] or 0) * 100,
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


def request_sina(req_codes: list[str]) -> dict:
    """新浪财经实时行情（A股/港股/美股/上期所期货）。
    返回 (ret, data)；ret[i] 表示 req_codes[i] 是否成功。"""
    data = {}
    if not req_codes:
        return [], {}
    label = ",".join(_sina_code(c) for c in req_codes if str(c).strip())
    url = "https://hq.sinajs.cn/list=" + label
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
        if sname.startswith("rt_"):
            entry = _parse_sina_hk(parts, is_index=_is_index_sina(sname))      # 港股股票 + 港股指数
        elif sname.startswith("gb_"):
            entry = _parse_sina_us(parts, is_index=_is_index_sina(sname))      # 美股 + 美股指数
        elif sname.startswith("nf_"):
            entry = _parse_sina_futures(parts)
        elif sname.startswith("b_"):
            entry = _parse_sina_global(parts)  # 全球指数(日经/KOSPI/DAX等)
        elif sname.startswith(("sh", "sz", "bj")):
            entry = _parse_sina_a(parts, is_index=_is_index_sina(sname))
        else:
            continue
        data[_canonical_from_sina(sname)] = entry
    return data


# ---------------- 东财解析 ----------------

def request_eastmoney(req_codes: list[str]) -> Tuple[list, dict]:
    """东方财富实时行情（A股/港股/美股/上期所期货），字段与新浪统一。
    东财独有字段（换手率 f8 / 量比 f10）已随请求拉取，字典中预留（注释）。"""
    data = {}
    if not req_codes:
        return [], {}
    secids = [_em_secid(c) for c in req_codes]
    params = {
        "secids": ",".join(secids),
        "fields": "f12,f14,f2,f3,f4,f5,f6,f15,f16,f17,f18,f8,f10",
        "fltt": 2,
        "invt": 2,
    }
    response = requests.get(_EM_QUOTE_URL, params=params, timeout=3)
    diff = ((response.json() or {}).get("data") or {}).get("diff") or []
    raw_to_code = {s.split(".", 1)[1]: c for c, s in zip(req_codes, secids)}
    for d in diff:
        code = raw_to_code.get(d.get("f12"))
        if code is None:
            continue
        vol = d.get("f5") or 0
        if market_of(code) in ("sh", "sz", "bj"):
            vol = vol * 100   # 东财 A股 f5 单位是“手”，新浪为“股”，统一为股
        entry = _new_entry(
            name=d.get("f14"),
            opening=d.get("f17"), prev_close=d.get("f18"), current=d.get("f2"),
            high=d.get("f15"), low=d.get("f16"),
            vol=vol, amt=d.get("f6"),
            pur_vol=_Z5, pur_price=_Z5, sell_vol=_Z5, sell_price=_Z5,
            date="", time="",
        )
        # 东财独有字段（新浪无），预留，取消注释即可填充：
        # entry["turnover_rate"] = d.get("f8")   # 换手率（%）
        # entry["volume_ratio"] = d.get("f10")   # 量比
        data[code] = entry
    return [c in data for c in req_codes], data


def request_quote(req_codes: list[str], source: str = DATA_SOURCE) -> Tuple[list, dict]:
    """统一行情入口。默认使用 DATA_SOURCE（当前为新浪），切换东财改 DATA_SOURCE 即可。"""
    if source == "eastmoney":
        return request_eastmoney(req_codes)
    return request_sina(req_codes)

