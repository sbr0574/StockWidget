"""
GitHub Action 专用：拉取全市场代码并写入 resources/ 下 JSON
本脚本由 .github/workflows/update-codes.yml 调用, 保持自包含, 不依赖项目运行时包
沪深股票和基金直接使用交易所官方接口，其余市场沿用东财或新浪接口
"""
import json
import os
import re
import sys
import traceback
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests

from pypinyin import Style, pinyin

# 脚本位于 <root>/.github/scripts/
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

STOCK_FILE = "stock_codes_list.json"      # 沪深京、港美股及全球主要指数列表
FUTURES_FILE = "futures_codes_list.json"  # 上期所期货列表

_EM_CLIST_URL = "https://push2delay.eastmoney.com/api/qt/clist/get"
_EM_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://quote.eastmoney.com/center/gridlist.html",
}
_SINA_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://vip.stock.finance.sina.com.cn/",
}
_DF_COLUMNS = ["code", "name", "name_en", "type", "market"]


# ----------------- 工具函数 -----------------

def _has_cjk(text: str) -> bool:
    """判断字符串是否包含 CJK 中日韩字符。"""
    return any("\u4e00" <= ch <= "\u9fff" for ch in str(text or ""))

def _concat_dedup(frames: list[pd.DataFrame]) -> pd.DataFrame:
    df = pd.concat(frames, ignore_index=True)
    return df.drop_duplicates(subset=["market", "code"]).reset_index(drop=True)


# ----------------- 交易所数据 -----------------

def _exchange_code(value) -> str:
    """把交易所工作簿中的数字代码规范为六位字符串。"""
    text = str(value or "").strip()
    if not text or text.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        return text.split(".", 1)[0].zfill(6)
    return ""

def _rows_frame(rows: list[tuple[str, str, str, str, str]]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=_DF_COLUMNS).drop_duplicates(
        subset=["market", "code"]
    ).reset_index(drop=True)

def _stock_sh_name_code(symbol: str) -> pd.DataFrame:
    """获取上交所股票列表。"""
    symbol_map = {
        "主板A股": ("1", "A_STOCK_CODE"),
        "主板B股": ("2", "B_STOCK_CODE"),
        "科创板": ("8", "A_STOCK_CODE"),
    }
    stock_type, code_field = symbol_map[symbol]
    response = requests.get(
        "https://query.sse.com.cn/sseQuery/commonQuery.do",
        params={
            "STOCK_TYPE": stock_type,
            "REG_PROVINCE": "",
            "CSRC_CODE": "",
            "STOCK_CODE": "",
            "sqlId": "COMMON_SSE_CP_GPJCTPZ_GPLB_GP_L",
            "COMPANY_STATUS": "2,4,5,7,8",
            "type": "inParams",
            "isPagination": "true",
            "pageHelp.cacheSize": "1",
            "pageHelp.beginPage": "1",
            "pageHelp.pageSize": "10000",
            "pageHelp.pageNo": "1",
            "pageHelp.endPage": "1",
        },
        headers={
            "Host": "query.sse.com.cn",
            "Pragma": "no-cache",
            "Referer": "https://www.sse.com.cn/assortment/stock/list/share/",
            "User-Agent": _EM_HEADERS["User-Agent"],
        },
        timeout=15,
    )
    response.raise_for_status()
    rows = []
    for item in (response.json() or {}).get("result") or []:
        code = _exchange_code(item.get(code_field))
        name = str(item.get("SEC_NAME_CN") or "").strip()
        mtype = "沪" if stock_type in {"1", "2"} else "科"
        if code and name:
            rows.append((code, name, "", mtype, "sh"))
    return _rows_frame(rows)

def _szse_xlsx(catalog_id: str, tab_key: str, referer: str) -> pd.DataFrame:
    """下载并解析深交所 XLSX"""
    response = requests.get(
        "https://www.szse.cn/api/report/ShowReport"
        if catalog_id == "1110"
        else "https://fund.szse.cn/api/report/ShowReport",
        params={
            "SHOWTYPE": "xlsx",
            "CATALOGID": catalog_id,
            "TABKEY": tab_key,
            "random": "0.6935816432433362",
        },
        headers={
            "Referer": referer,
            "User-Agent": _EM_HEADERS["User-Agent"],
        },
        timeout=15,
    )
    response.raise_for_status()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        return pd.read_excel(BytesIO(response.content), engine="openpyxl", dtype=str)

def _stock_sz_a_name_code() -> pd.DataFrame:
    """下载深交所 A 股表，并按交易所板块字段拆分主板和创业板。"""
    table = _szse_xlsx(
        "1110", "tab1", "https://www.szse.cn/market/product/stock/list/index.html"
    )
    frames = []
    for board in ("主板", "创业板"):
        rows = []
        board_rows = table[table["板块"].astype(str).str.strip() == board]
        for _, item in board_rows.iterrows():
            code = _exchange_code(item.get("A股代码"))
            name = str(item.get("A股简称") or "").strip()
            mtype = "深" if board == "主板" else "创"
            if code and name:
                rows.append((code, name, "", mtype, "sz"))
        frame = _rows_frame(rows)
        print(f"深证{board}: {len(frame)} 条", flush=True)
        frames.append(frame)
    return _concat_dedup(frames)

def _stock_sz_b_name_code() -> pd.DataFrame:
    """获取深交所 B 股列表。"""
    table = _szse_xlsx(
        "1110", "tab2", "https://www.szse.cn/market/product/stock/list/index.html"
    )
    rows = []
    for _, item in table.iterrows():
        code = _exchange_code(item.get("B股代码"))
        name = str(item.get("B股简称") or "").strip()
        if code and name:
            rows.append((code, name, "", "深", "sz"))
    return _rows_frame(rows)

def _fund_sse_rows() -> dict[str, list[tuple[str, str, str, str, str]]]:
    """获取上交所上市交易基金全表并按官网 fundType 分类。"""
    result = {"ETF基金": [], "LOF基金": [], "封闭式基金": []}
    page = 1
    while True:
        response = requests.get(
            "https://query.sse.com.cn/commonSoaQuery.do",
            params={
                "isPagination": "true",
                "pageHelp.pageSize": "2000",
                "pageHelp.pageNo": str(page),
                "pageHelp.beginPage": str(page),
                "pageHelp.cacheSize": "1",
                "pageHelp.endPage": str(page),
                "pagecache": "false",
                "sqlId": "FUND_LIST",
                "order": "fundCode|asc",
            },
            headers={
                "Referer": "https://www.sse.com.cn/assortment/fund/list/",
                "User-Agent": _EM_HEADERS["User-Agent"],
            },
            timeout=15,
        )
        response.raise_for_status()
        payload = response.json() or {}
        for item in payload.get("result") or []:
            fund_type = str(item.get("fundType") or "").strip()
            if fund_type == "00":
                category = "ETF基金"
            elif fund_type == "10":
                category = "LOF基金"
            elif fund_type in {"40", "50"}:
                category = "封闭式基金"
            else:
                continue
            code = _exchange_code(item.get("fundCode"))
            name = str(item.get("secNameFull") or item.get("fundAbbr") or "").strip()
            if code and name:
                result[category].append((code, name, "", "基", "sh"))
        page_help = payload.get("pageHelp") or {}
        if page >= int(page_help.get("pageCount") or 1):
            break
        page += 1
    return result

def _fund_szse_rows() -> dict[str, list[tuple[str, str, str, str, str]]]:
    """获取深交所基金全表，并使用工作簿的基金类别字段分类。"""
    table = _szse_xlsx(
        "1000_lf",
        "tab1",
        "https://fund.szse.cn/marketdata/fundslist/index.html",
    )
    category_map = {
        "ETF": "ETF基金",
        "LOF": "LOF基金",
        "不动产基金": "封闭式基金",
        "基础设施基金": "封闭式基金",
        "REITs": "封闭式基金",
        "REITS": "封闭式基金",
    }
    result = {"ETF基金": [], "LOF基金": [], "封闭式基金": []}
    for _, item in table.iterrows():
        category = category_map.get(str(item.get("基金类别") or "").strip())
        code = _exchange_code(item.get("基金代码"))
        name = str(item.get("基金简称") or "").strip()
        if category and code and name:
            result[category].append((code, name, "", "基", "sz"))
    return result

def _fund_exchange_all() -> pd.DataFrame:
    """合并沪深交易所官方基金列表"""
    sh_rows = _fund_sse_rows()
    sz_rows = _fund_szse_rows()
    frames = []
    for category in ("ETF基金", "LOF基金", "封闭式基金"):
        frame = _rows_frame(sh_rows[category] + sz_rows[category])
        print(f"{category}: {len(frame)} 条", flush=True)
        frames.append(frame)
    return _concat_dedup(frames)


# ----------------- 新浪数据 -----------------

def _stock_hk_name_code() -> pd.DataFrame:
    """港股全部股票"""
    rows = []
    page = 1
    while True:
        diff = None
        for _ in range(3):
            try:
                r = requests.get(
                    "https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/Market_Center.getHKStockData",
                    params={"page": page, "num": 60, "sort": "symbol", "asc": 1,
                            "node": "qbgg_hk", "_s_r_a": "init"},
                    headers=_SINA_HEADERS,
                    timeout=10,
                )
                diff = r.json()
                if not isinstance(diff, list):
                    diff = []
                break
            except Exception:
                diff = None
        if not diff:
            break
        for d in diff:
            code = str(d.get("symbol") or "").strip().zfill(5)
            name = str(d.get("name") or "").strip()
            eng = str(d.get("engname") or "").strip()
            if code:
                rows.append((code, name, eng, "港", "hk"))
        if len(diff) < 60:
            break
        page += 1
    return pd.DataFrame(rows, columns=_DF_COLUMNS)


# ----------------- 东财数据 -----------------

def _em_clist_all(fs: str, fid: str = "f12", fields: str = "f12,f14") -> list[dict]:
    """东财 clist 分页拉全市场原始 dict 列表；单页失败重试 3 次，断连返回已收集部分。"""
    _em_page_size = 500
    rows = []
    page = 1
    total = None
    while True:
        diff = None
        for _ in range(3):
            try:
                r = requests.get(
                    _EM_CLIST_URL,
                    params={
                        "pn": page, "pz": _em_page_size, "po": 1, "np": 1, "fltt": 2, "invt": 2,
                        "ut": "bd1d9ddb04089700cf9c27f6f7426281",
                        "fid": fid, "fs": fs, "fields": fields,
                    },
                    headers=_EM_HEADERS,
                    timeout=10,
                )
                data = ((r.json() or {}).get("data") or {})
                diff = data.get("diff") or []
                if data.get("total") is not None:
                    total = int(data.get("total") or 0)
                break
            except Exception:
                diff = None
        if not diff:
            break
        rows.extend(diff)
        if total is not None and len(rows) >= total:
            break
        if total is None and len(diff) < _em_page_size:
            break
        page += 1
    return rows

def _em_stock_df(fs: str, mtype: str, market: str) -> pd.DataFrame:
    """从东财 clist 拉取一个分类的股票/基金/指数列表 返回df"""
    rows = []
    for d in _em_clist_all(fs, fields="f12,f14"):
        code = str(d.get("f12") or "").strip()
        name = str(d.get("f14") or "").strip()
        if not code or name.startswith("退市") or name.endswith("退"):
            continue
        rows.append((code, name, "", mtype, market))
    df = pd.DataFrame(rows, columns=_DF_COLUMNS)
    return df.drop_duplicates(subset=["market", "code"]).reset_index(drop=True)

def _index_cn_em() -> pd.DataFrame:
    """东财沪深指数列表。"""
    return _concat_dedup([
        _em_stock_df("m:1+t:1", "指", market="sh"),
        _em_stock_df("m:0+t:5", "指", market="sz"),
    ])

def _stock_us_name_code() -> pd.DataFrame:
    """美股全部股票"""
    rows = []
    for d in _em_clist_all("m:105,m:106,m:107", fid="f20"):
        code = str(d.get("f12") or "").strip().lower()
        name = str(d.get("f14") or "").strip()
        if not code:
            continue
        eng = name if not _has_cjk(name) else ""
        rows.append((code, name, eng, "美", "us"))
    return pd.DataFrame(rows, columns=_DF_COLUMNS)


# ----------------- 离线数据 -----------------

def _stock_hk_index_name_code() -> pd.DataFrame:
    """港股指数列表。"""
    _HK_INDEXES = (
        ("ces100", "中华港股通精选100指数"),
        ("ces120", "中华120指数"),
        ("ces280", "中华280指数"),
        ("ces300", "中华沪深港300指数"),
        ("cesa80", "中华A80指数"),
        ("cesg10", "中华博彩业指数"),
        ("ceshkm", "中华香港内地指数"),
        ("cscmc", "中证内地消费指数"),
        ("cshk100", "中证香港100指数"),
        ("cshkdiv", "中证香港红利港币指数"),
        ("cshklc", "中证香港上市可交易内地消费指数"),
        ("cshklre", "中证香港上市可交易内地地産指数"),
        ("cshkmcs", "中证香港中盘精选港币指数"),
        ("cshkme", "中证香港内地股港元指数"),
        ("cshkpe", "中证香港内地民营企业指数"),
        ("cshkse", "中证香港内地国有企业指数"),
        ("csrhk50", "中证锐联香港基本面50港币指数"),
        ("gem", "标普香港创业板指数"),
        ("hkl", "标普香港大型股指数"),
        ("hscci", "恒生香港中资企业指数"),
        ("hscei", "恒生中国企业指数"),
        ("hsi", "恒生指数"),
        ("hsmbi", "恒生中国内地银行指数"),
        ("hsmogi", "恒生中国内地石油及天然气指数"),
        ("hsmpi", "恒生中国内地地产指数"),
        ("hstech", "恒生科技指数"),
        ("vhsi", "恒指波幅指数"),
    )
    rows = [(code, name, "", "指", "hk") for code, name in _HK_INDEXES]
    return pd.DataFrame(rows, columns=_DF_COLUMNS)

def _stock_global_index_name_code() -> pd.DataFrame:
    """全球主要指数列表"""
    _GLOBAL_INDEXES = (
        ("aex", "荷兰AEX综合指数"),
        ("as51", "澳大利亚标准普尔200指数"),
        ("cac", "法CAC40指数"),
        ("case", "埃及CASE 30指数"),
        ("dax", "德国DAX 30种股价指数"),
        ("ftsemib", "富时意大利MIB指数"),
        ("gsptse", "加拿大S&P/TSX综合指数"),
        ("ibex", "西班牙IBEX指数"),
        ("ibov", "巴西BOVESPA股票指数"),
        ("indexcf", "俄罗斯MICEX指数"),
        ("jci", "印度尼西亚雅加达综合指数"),
        ("kospi", "首尔综合指数"),
        ("mxx", "墨西哥BOLSA指数"),
        ("nky", "日经225指数"),
        ("nz250", "新西兰NZSE 50指数"),
        ("sensex", "印度孟买SENSEX指数"),
        ("swi20", "瑞士股票指数"),
        ("sx5e", "欧洲Stoxx50指数"),
        ("twjq", "中国台湾加权指数"),
        ("ukx", "英国富时100指数"),
    )
    rows = [(code, name, "", "指", "gb") for code, name in _GLOBAL_INDEXES]
    return pd.DataFrame(rows, columns=_DF_COLUMNS)

def _stock_us_index_name_code() -> pd.DataFrame:
    """美股主要指数"""
    rows = [
        ("ixic", "纳斯达克综合指数", "", "指", "us"),
        ("dji", "道琼斯工业平均指数", "", "指", "us"),
        ("inx", "标普500指数", "", "指", "us"),
        ("ndx", "纳斯达克100指数", "", "指", "us"),
    ]
    return pd.DataFrame(rows, columns=_DF_COLUMNS)


# ----------------- 股指列表 -----------------

def _run_frame_tasks(tasks: list[tuple]) -> list[pd.DataFrame]:
    """并发拉取互不依赖的分类，按 tasks 原顺序返回，保证最终 JSON 稳定。"""
    if not tasks:
        return []
    results: list[pd.DataFrame | None] = [None] * len(tasks)
    workers = min(6, len(tasks))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {}
        for idx, task in enumerate(tasks):
            label, func, args, kwargs = task
            futures[pool.submit(func, *args, **kwargs)] = idx
        for future in as_completed(futures):
            idx = futures[future]
            label = tasks[idx][0]
            try:
                results[idx] = future.result()
            except Exception:
                traceback.print_exc()
                results[idx] = pd.DataFrame(columns=_DF_COLUMNS)
            print(f"{label}更新完成: {len(results[idx])} 条", flush=True)
    return [df if df is not None else pd.DataFrame(columns=_DF_COLUMNS) for df in results]

def _tasks() -> list[tuple]:
    return [
        ("沪A", _stock_sh_name_code, ("主板A股",), {}),
        ("科创板", _stock_sh_name_code, ("科创板",), {}),
        ("沪B", _stock_sh_name_code, ("主板B股",), {}),
        ("深A及创业板", _stock_sz_a_name_code, (), {}),
        ("深B", _stock_sz_b_name_code, (), {}),
        ("京市", _em_stock_df, ("m:0+t:81+s:2048", "京"), {"market": "bj"}),
        ("沪深基金", _fund_exchange_all, (), {}),
        ("国内指数", _index_cn_em, (), {}),
        ("港股", _stock_hk_name_code, (), {}),
        ("港股指数", _stock_hk_index_name_code, (), {}),
        ("美股", _stock_us_name_code, (), {}),
        ("全球股指", _stock_global_index_name_code, (), {}),
        ("美股指数", _stock_us_index_name_code, (), {}),
    ]

def stock_info_all() -> pd.DataFrame:
    """非期货证券"""
    frames = _run_frame_tasks(_tasks())
    return pd.concat(frames, ignore_index=True)


# ----------------- 上期所期货 -----------------

def _futures_nodes() -> list[str]:
    """从新浪节点脚本中提取上期所及上期能源节点名。"""
    url = "https://vip.stock.finance.sina.com.cn/quotes_service/view/js/qihuohangqing.js"
    r = requests.get(url, headers=_SINA_HEADERS, timeout=15)
    r.encoding = "gb2312"
    match = re.search(r"\bshfe\s*:\s*\[(.*?)\]\s*,\s*cffex\s*:", r.text, re.DOTALL)
    if not match:
        raise ValueError("新浪期货节点脚本中缺少 shfe 列表")
    return re.findall(r"\[\s*'[^']*'\s*,\s*'([^']+)'", match.group(1))

def _stock_shfe_futures() -> pd.DataFrame:
    """上期所全部期货合约（含上期能源；新浪 getHQFuturesData 按品种遍历）。"""
    rows = []
    try:
        for node in _futures_nodes():
            for page in range(1, 4):
                try:
                    r = requests.get(
                        "https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/Market_Center.getHQFuturesData",
                        params={"page": page, "sort": "position", "asc": "0", "node": node, "base": "futures"},
                        headers=_SINA_HEADERS,
                        timeout=8,
                    )
                    j = r.json()
                    if not isinstance(j, list) or not j:
                        break
                    for d in j:
                        symbol = str(d.get("symbol", "") or "").strip().lower()
                        name = str(d.get("name", "") or "").strip()
                        if symbol:
                            rows.append((symbol, name))
                    if len(j) < 20:
                        break
                except Exception:
                    break
    except Exception:
        rows = []
    seen, uniq = set(), []
    for c, n in rows:
        if c not in seen:
            seen.add(c)
            uniq.append((c, n))
    rows = [(code, name, "", "期", "") for code, name in uniq]
    return pd.DataFrame(rows, columns=_DF_COLUMNS)

def futures_info_all() -> pd.DataFrame:
    """上期所期货"""
    df = _stock_shfe_futures()
    print(f"期货: {len(df)} 条", flush=True)
    return df


# ---------------- 市场列表字典生成 ----------------

def _to_halfwidth(text: str) -> str:
    """全角字符转半角, 避免全角字母匹配不上搜索。"""
    out = []
    for ch in str(text or ""):
        o = ord(ch)
        if o == 0x3000:
            out.append(" ")                      # 全角空格
        elif 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))          # 全角 -> 半角
        else:
            out.append(ch)
    return "".join(out)

def _name_pinyin(name: str) -> tuple[str, str]:
    """中文名转拼音全拼和首字母缩写, 非中文返回空串"""
    text = _to_halfwidth(name).strip()
    if not text:
        return "", ""
    text = text.replace(" ", "")
    py_full = "".join(x[0] for x in pinyin(text, style=Style.NORMAL, strict=False))
    py_abbr = "".join(x[0] for x in pinyin(text, style=Style.FIRST_LETTER, strict=False))
    return py_full.lower(), py_abbr.lower()

def _normalize_code(code: str, market: str) -> str:
    """标准化证券代码：沪深京 6 位、港股 5 位、其他原样返回"""
    code = str(code or "").strip().lower()
    market = str(market or "").strip().lower()
    if market in {"sh", "sz", "bj"} and code.isdigit():
        return code.zfill(6)
    if market == "hk" and code.isdigit():
        return code.zfill(5)
    return code

def _code_sort_key(item: tuple[str, dict]) -> tuple:
    """按市场顺序、市场、代码排序"""
    _market_order = {"sh": 0, "sz": 1, "bj": 2, "hk": 3, "us": 4, "gb": 5, "": 6}
    key, entry = item
    market = str(entry.get("market", "") or "").strip().lower()
    code = str(entry.get("code", "") or "").strip().lower()
    return (_market_order.get(market, 99), market, code, key)

def _df_to_dict(df: pd.DataFrame) -> dict:
    """把 [code,name,name_en,type,market] 的 df 转成有序字典。"""
    codes = {}
    for _, row in df.iterrows():
        market = str(row.get("market", "") or "").strip().lower()
        code = _normalize_code(row.get("code", ""), market)
        if not code:
            continue
        name = _to_halfwidth(str(row.get("name", "") or "")).strip()
        name_en = _to_halfwidth(str(row.get("name_en", "") or "")).strip()
        if not name and name_en: name = name_en
        mtype = str(row.get("type", "") or "").strip()
        py_full, py_abbr = _name_pinyin(name) if _has_cjk(name) else ("", "")
        entry = {
            "code": code,
            "market": market,
            "type": mtype,
            "name": name,
            "name_en": name_en,
            "py": py_full,
            "abbr": py_abbr,
        }
        codes[market + code] = entry
    return dict(sorted(codes.items(), key=_code_sort_key))


# ---------------- 拉取市场列表并保存 ----------------

def fetch_codes_groups() -> dict[str, dict]:
    """拉取两组代码，返回 {文件名: {"last_update": "YYYY-MM-DD", "codes": {...}}}。"""
    now = datetime.now().strftime("%Y-%m-%d")
    return {
        STOCK_FILE: {
            "last_update": now,
            "codes": _df_to_dict(stock_info_all())
        },
        FUTURES_FILE: {
            "last_update": now,
            "codes": _df_to_dict(futures_info_all()),
        },
    }

def main() -> int:
    try:
        groups = fetch_codes_groups()
    except Exception as exc:
        traceback.print_exc()
        print(f"::error::拉取代码列表失败: {exc}")
        return 1

    resources_dir = os.path.join(ROOT, "resources")
    os.makedirs(resources_dir, exist_ok=True)
    for fname, data in groups.items():
        path = os.path.join(resources_dir, fname)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)

    for fname, data in groups.items():
        print(f"更新 {fname}: {len(data.get('codes', {}))} 条")
    return 0


if __name__ == "__main__":
    sys.exit(main())
