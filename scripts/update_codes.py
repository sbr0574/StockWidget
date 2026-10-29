"""拉取全市场代码并写入仓库 resources/data/ 下的 JSON。

直接运行本文件即可，不依赖 StockWidget 运行时模块。股票与基金使用交易所
官方名录，美股中文名称与期货使用新浪。指数 JSON 由维护者直接
在 codes-data 分支维护，本脚本只读取其状态，不生成或改写指数列表。
"""
import csv
import json
import os
import re
import sys
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from io import BytesIO, StringIO
from threading import Lock

import pandas as pd
import requests

from openpyxl import load_workbook
from pypinyin import Style, pinyin

# 脚本位于 <root>/scripts/。服务器可用 CODES_OUTPUT_DIR 指向 codes-data
# 工作树，确保失败分类保留数据分支中的上一版文件。
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.environ.get("CODES_OUTPUT_DIR", os.path.join(ROOT, "resources", "data"))

STATUS_FILE = "codes_update_status.json"
MANUAL_FILES = ("index_cn.json", "index_global.json")
US_CN_ALIAS_CACHE_FILE = "cache_us_cn_aliases.json"
TASK_ATTEMPTS = 3
TASK_RETRY_SECONDS = (2, 5)

_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
_HKEX_SECURITIES_EN_URL = "https://www.hkex.com.hk/eng/services/trading/securities/securitieslists/ListOfSecurities.xlsx"
_HKEX_SECURITIES_ZH_URL = "https://www.hkex.com.hk/chi/services/trading/securities/securitieslists/ListOfSecurities_c.xlsx"
_HKEX_HEADERS = {
    "User-Agent": _USER_AGENT,
    "Referer": "https://www.hkex.com.hk/Services/Trading/Securities/Securities-Lists",
}
_HKEX_CATEGORY_TYPES = {
    "Equity": "港",
    "Exchange Traded Products": "基",
    "Real Estate Investment Trusts": "基",
}
_NASDAQ_LISTED_URL = "https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqlisted.txt"
_NASDAQ_OTHER_LISTED_URL = (
    "https://www.nasdaqtrader.com/dynamic/SymDir/otherlisted.txt"
)
_NASDAQ_HEADERS = {
    "User-Agent": _USER_AGENT,
    "Referer": "https://www.nasdaqtrader.com/trader.aspx?id=symboldirdefs",
}
_SINA_HEADERS = {
    "User-Agent": _USER_AGENT,
    "Referer": "https://vip.stock.finance.sina.com.cn/",
}
_DF_COLUMNS = ["code", "name", "name_en", "type", "market"]
_SZSE_REQUEST_LOCK = Lock()


# ----------------- 分页请求与校验 -----------------

_REQUEST_INTERVAL = .5
_REQUEST_LOCK = Lock()
_last_request_started = None
_BSE_URL = "https://www.bse.cn/nqxxController/nqxxCnzq.do"
_BSE_HEADERS = {
    "User-Agent": _USER_AGENT,
    "Referer": "https://www.bse.cn/nq/listedcompany.html",
}


class CodeListError(ValueError):
    """已在当前页完成有限重试，或列表未通过完整性检查。"""


def retryable_error(error: Exception) -> bool:
    """仅重试临时传输或服务端错误，不重试拒绝、限流或证书错误。"""
    if isinstance(error, requests.exceptions.SSLError):
        return False
    if isinstance(error, requests.HTTPError):
        return error.response is not None and error.response.status_code in {
            500, 502, 503, 504,
        }
    return isinstance(error, (
        requests.ConnectionError, requests.Timeout,
        requests.exceptions.ChunkedEncodingError,
    ))


def safe_error(error: Exception) -> str:
    """日志与发布状态不包含请求地址、代理凭据或原始异常栈。"""
    if isinstance(error, CodeListError):
        return str(error)
    if isinstance(error, requests.HTTPError):
        status = error.response.status_code if error.response is not None else None
        return f"HTTP {status}" if status is not None else "HTTP 请求失败"
    if isinstance(error, requests.exceptions.ProxyError):
        return "代理连接失败"
    if isinstance(error, requests.exceptions.SSLError):
        return "安全连接或证书验证失败"
    if isinstance(error, requests.Timeout):
        return "请求超时"
    if isinstance(error, requests.ConnectionError):
        return "连接失败或被远端中断"
    if isinstance(error, requests.exceptions.JSONDecodeError):
        return "响应不是有效 JSON"
    if isinstance(error, requests.exceptions.ChunkedEncodingError):
        return "响应传输中断"
    if isinstance(error, requests.RequestException):
        return "网络请求失败"
    return f"数据处理失败（{type(error).__name__}）"


def _request_page(url, label, *, headers, params=None, data=None, as_text=False):
    """分页请求共用限速、有限重试和 JSON/JSONP 解析；不重跑已完成页。"""
    global _last_request_started
    for attempt in range(3):
        try:
            with _REQUEST_LOCK:
                if _last_request_started is not None:
                    delay = _REQUEST_INTERVAL - (time.monotonic() - _last_request_started)
                    if delay > 0:
                        time.sleep(delay)
                _last_request_started = time.monotonic()
                request = requests.get if data is None else requests.post
                kwargs = {"params": params} if data is None else {"data": data}
                response = request(url, headers=headers, timeout=(10, 20), **kwargs)
                try:
                    response.raise_for_status()
                    if as_text:
                        response.encoding = "gbk"
                    text = response.text.lstrip("\ufeff").strip()
                finally:
                    response.close()
            break
        except requests.RequestException as error:
            reason = safe_error(error)
            if attempt == 2 or not retryable_error(error):
                raise CodeListError(f"{label}失败：{reason}") from None
            delay = TASK_RETRY_SECONDS[attempt]
            print(f"{label}失败，{delay} 秒后重试当前页：{reason}", flush=True)
            time.sleep(delay)
    if as_text:
        return text
    if not text.startswith(("[", "{")):
        match = re.fullmatch(r"[\w$]+\s*\((.*)\)\s*;?", text, re.DOTALL)
        if not match:
            raise CodeListError(f"{label}不是有效列表响应")
        text = match.group(1)
    try:
        return json.loads(text)
    except ValueError:
        raise CodeListError(f"{label}不是有效 JSON") from None




def fetch_bse_list() -> list[dict]:
    """从北交所官网拉齐上市股票；缺页、重复或范围错误均保留旧列表。"""
    rows, seen = [], set()
    expected = None
    page = 0
    while True:
        label = f"北交所官方名录第 {page + 1} 页"
        payload = _request_page(_BSE_URL, label, headers=_BSE_HEADERS, data={
            "page": str(page), "typejb": "T", "xxfcbj[]": "2", "xxzqdm": "",
            "sortfield": "xxzqdm", "sorttype": "asc",
        })
        if not isinstance(payload, list) or len(payload) != 1 or not isinstance(payload[0], dict):
            raise CodeListError(f"{label}格式错误")
        data = payload[0]
        fields = ("number", "numberOfElements", "size", "totalElements", "totalPages")
        if any(type(data.get(field)) is not int for field in fields):
            raise CodeListError(f"{label}分页信息无效")
        total, pages, size = data["totalElements"], data["totalPages"], data["size"]
        if (total <= 0 or size <= 0 or not 0 < pages <= 10000
                or pages != (total + size - 1) // size or data["number"] != page):
            raise CodeListError(f"{label}分页范围错误")
        if expected is None:
            expected = (total, pages, size)
        elif expected != (total, pages, size):
            raise CodeListError(f"{label}分页总数发生变化")
        content = data.get("content")
        count = min(size, total - page * size)
        if (not isinstance(content, list) or len(content) != count
                or data["numberOfElements"] != count):
            raise CodeListError(f"{label}数量不完整")
        for item in content:
            if not isinstance(item, dict):
                raise CodeListError(f"{label}证券数据无效")
            code = str(item.get("xxzqdm") or "").strip()
            if (not re.fullmatch(r"\d{6}", code) or not str(item.get("xxzqjc") or "").strip()
                    or item.get("xxfcbj") != "2" or item.get("xxzqjb") != "T"):
                raise CodeListError(f"{label}缺少证券信息或包含非上市股票")
            if code in seen:
                raise CodeListError(f"{label}证券代码重复")
            seen.add(code)
            rows.append(item)
        if page + 1 == pages:
            return rows
        page += 1


# ----------------- 工具函数 -----------------

def _has_cjk(text: str) -> bool:
    """判断字符串是否包含 CJK 中日韩字符。"""
    return any("\u4e00" <= ch <= "\u9fff" for ch in str(text or ""))


def _concat_dedup(frames: list[pd.DataFrame]) -> pd.DataFrame:
    df = pd.concat(frames, ignore_index=True)
    return df.drop_duplicates(subset=["market", "code"]).reset_index(drop=True)


# ----------------- 交易所数据 -----------------

def _exchange_code(value, width: int = 6) -> str:
    """把交易所工作簿中的数字代码规范为固定宽度字符串。"""
    text = str(value or "").strip()
    if not text or text.lower() == "nan":
        return ""
    if re.fullmatch(rf"\d{{1,{width}}}(?:\.0+)?", text):
        return text.split(".", 1)[0].zfill(width)
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
            "REG_PROVINCE": "", "CSRC_CODE": "", "STOCK_CODE": "",
            "sqlId": "COMMON_SSE_CP_GPJCTPZ_GPLB_GP_L",
            "COMPANY_STATUS": "2,4,5,7,8",
            "type": "inParams",
            "isPagination": "true",
            "pageHelp.cacheSize": "1", "pageHelp.beginPage": "1",
            "pageHelp.pageSize": "10000", "pageHelp.pageNo": "1", "pageHelp.endPage": "1",
        },
        headers={
            "Host": "query.sse.com.cn",
            "Pragma": "no-cache",
            "Referer": "https://www.sse.com.cn/assortment/stock/list/share/",
            "User-Agent": _USER_AGENT,
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

def _stock_sh_all() -> pd.DataFrame:
    """顺序获取上交所 A 股、B 股和科创板。"""
    frames = []
    for label, symbol in (
        ("沪A", "主板A股"),
        ("沪B", "主板B股"),
        ("科创板", "科创板"),
    ):
        frame = _stock_sh_name_code(symbol)
        print(f"{label}: {len(frame)} 条", flush=True)
        if frame.empty:
            raise ValueError(f"{label}返回列表为空")
        frames.append(frame)
    return _concat_dedup(frames)

def _szse_xlsx(catalog_id: str, tab_key: str, referer: str) -> pd.DataFrame:
    """下载并解析一次深交所 XLSX，失败后由分类任务统一重试。"""
    url = (
        "https://www.szse.cn/api/report/ShowReport"
        if catalog_id == "1110"
        else "https://fund.szse.cn/api/report/ShowReport"
    )
    with _SZSE_REQUEST_LOCK:
        response = requests.get(
            url,
            params={
                "SHOWTYPE": "xlsx",
                "CATALOGID": catalog_id,
                "TABKEY": tab_key,
                "random": f"{time.time():.16f}",
            },
            headers={
                "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
                "Connection": "close",
                "Referer": referer,
                "User-Agent": _USER_AGENT,
            },
            timeout=(15, 30),
        )
        response.raise_for_status()
        if not response.content.startswith(b"PK"):
            content_type = response.headers.get("Content-Type", "unknown")
            raise ValueError(f"深交所返回的不是 XLSX: {content_type}")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            table = pd.read_excel(
                BytesIO(response.content), engine="openpyxl", dtype=str
            )
        if table.empty:
            raise ValueError("深交所 XLSX 内容为空")
        return table

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
        if frame.empty:
            raise ValueError(f"深证{board}返回列表为空")
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
    print(f"深B: {len(rows)} 条", flush=True)
    if rows == []:
        raise ValueError("深B返回列表为空")
    return _rows_frame(rows)

def _stock_sz_all() -> pd.DataFrame:
    """顺序获取深交所 A 股、创业板和 B 股。"""
    return _concat_dedup([_stock_sz_a_name_code(), _stock_sz_b_name_code()])

def _stock_bj_name_code() -> pd.DataFrame:
    """以北交所官网在市股票名录为准，保留正式证券代码和中英文名称。"""
    rows = [
        (str(item["xxzqdm"]).strip(), str(item["xxzqjc"]).strip(),
         str(item.get("xxywjc") or "").strip(), "京", "bj")
        for item in fetch_bse_list()
    ]
    return _rows_frame(rows)


def _fund_sse_rows() -> dict[str, list[tuple[str, str, str, str, str]]]:
    """获取上交所上市交易基金全表并按官网 fundType 分类。"""
    result = {"ETF基金": [], "LOF基金": [], "封闭式基金": []}
    categories = {"00": "ETF基金", "10": "LOF基金", "40": "封闭式基金", "50": "封闭式基金"}
    page = 1
    while True:
        response = requests.get(
            "https://query.sse.com.cn/commonSoaQuery.do",
            params={
                "isPagination": "true",
                "pageHelp.pageSize": "2000",
                "pageHelp.pageNo": str(page), "pageHelp.beginPage": str(page),
                "pageHelp.cacheSize": "1", "pageHelp.endPage": str(page),
                "pagecache": "false",
                "sqlId": "FUND_LIST",
                "order": "fundCode|asc",
            },
            headers={
                "Referer": "https://www.sse.com.cn/assortment/fund/list/",
                "User-Agent": _USER_AGENT,
            },
            timeout=15,
        )
        response.raise_for_status()
        payload = response.json() or {}
        for item in payload.get("result") or []:
            category = categories.get(str(item.get("fundType") or "").strip())
            if not category:
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


# ----------------- 港交所数据 -----------------

def _hkex_xlsx_records(url: str, required_columns: tuple[str, ...]) -> list[dict]:
    """下载港交所完整证券名录，并返回指定列。"""
    response = requests.get(
        url,
        headers=_HKEX_HEADERS,
        timeout=(15, 60),
    )
    response.raise_for_status()
    if not response.content.startswith(b"PK"):
        content_type = response.headers.get("Content-Type", "unknown")
        raise ValueError(f"港交所返回的不是 XLSX: {content_type}")

    workbook = load_workbook(
        BytesIO(response.content), read_only=True, data_only=True
    )
    try:
        if "ListOfSecurities" not in workbook.sheetnames:
            raise ValueError("港交所工作簿缺少 ListOfSecurities 工作表")
        sheet = workbook["ListOfSecurities"]
        # 港交所文件的 dimension 当前仅声明到第 8 行；只读模式必须重置，
        # 否则 openpyxl 会忽略后续一万余条证券记录。
        sheet.reset_dimensions()
        row_iter = sheet.iter_rows(values_only=True)
        column_indexes = None
        for values in row_iter:
            cells = [str(value or "").strip() for value in values]
            if all(column in cells for column in required_columns):
                column_indexes = {
                    column: cells.index(column) for column in required_columns
                }
                break
        if column_indexes is None:
            raise ValueError(
                "港交所工作簿缺少必要列: " + ", ".join(required_columns)
            )

        records = []
        for values in row_iter:
            record = {
                column: values[index] if index < len(values) else None
                for column, index in column_indexes.items()
            }
            if any(value not in (None, "") for value in record.values()):
                records.append(record)
        return records
    finally:
        workbook.close()


def _stock_hk_name_code() -> pd.DataFrame:
    """从港交所官方名录获取港股股票、交易所买卖产品及 REIT。"""
    english_records = _hkex_xlsx_records(
        _HKEX_SECURITIES_EN_URL,
        ("Stock Code", "Name of Securities", "Category"),
    )
    chinese_records = _hkex_xlsx_records(
        _HKEX_SECURITIES_ZH_URL,
        ("股份代號", "股份名稱"),
    )
    from opencc import OpenCC

    to_simplified = OpenCC("hk2s").convert
    chinese_names = {
        code: to_simplified(str(item.get("股份名稱") or "").strip())
        for item in chinese_records
        if (code := _exchange_code(item.get("股份代號"), 5))
    }

    rows = []
    for item in english_records:
        category = str(item.get("Category") or "").strip()
        mtype = _HKEX_CATEGORY_TYPES.get(category)
        if not mtype:
            continue
        code = _exchange_code(item.get("Stock Code"), 5)
        name_en = str(item.get("Name of Securities") or "").strip()
        name = chinese_names.get(code, "")
        if not code or not (name or name_en):
            continue
        rows.append((code, name, name_en, mtype, "hk"))
    frame = _rows_frame(rows)
    print(f"港股股票: {(frame['type'] == '港').sum()} 条", flush=True)
    print(f"港股基金: {(frame['type'] == '基').sum()} 条", flush=True)
    return frame


# ----------------- 美股目录与中文名称 -----------------


def _clean_us_cn_aliases(value) -> dict[str, str]:
    """清理别名缓存，只保留带中文字符的 code -> name。"""
    aliases = {}
    for raw_code, raw_name in (value if isinstance(value, dict) else {}).items():
        code = str(raw_code or "").strip().lower()
        name = str(raw_name or "").strip()
        if code and _has_cjk(name):
            aliases[code] = name
    return aliases


def _fetch_us_cn_aliases(official_codes: set[str]) -> dict[str, str]:
    """按官方代码分批向新浪补充中文名，不使用行情列表决定证券范围。"""
    aliases = {}
    codes = sorted(code for code in official_codes if re.fullmatch(r"[a-z0-9_]+", code))
    for offset in range(0, len(codes), 150):
        batch = set(codes[offset:offset + 150])
        label = f"新浪美股名称第 {offset // 150 + 1} 批"
        text = _request_page("https://hq.sinajs.cn/list=" + ",".join("gb_" + code for code in sorted(batch)),
                             label, headers=_SINA_HEADERS, as_text=True)
        records = re.findall(r'var\s+hq_str_gb_([a-zA-Z0-9_]+)\s*=\s*"([^"\r\n]*)"', text)
        returned = {code.lower() for code, _ in records if code.lower() in batch}
        if returned != batch:
            raise CodeListError(f"{label}返回标的数量不符：{len(returned)}/{len(batch)}")
        for code, body in records:
            name = body.split(",", 1)[0].strip()
            if code.lower() in batch and _has_cjk(name):
                aliases[code.lower()] = name
    if not aliases:
        raise CodeListError("新浪美股未返回中文名称")
    return aliases


def _stock_us_name_code(
    today: date | None = None,
    output_dir: str | None = None,
    cache_path: str | None = None,
) -> pd.DataFrame:
    """每日以 Nasdaq 官方目录为准；跨月后的首次运行从新浪更新中文别名。"""
    output_dir = output_dir or OUTPUT_DIR
    cache_path = cache_path or os.path.join(output_dir, US_CN_ALIAS_CACHE_FILE)
    today = today or datetime.now(timezone(timedelta(hours=8))).date()

    official_rows = []
    sources = (
        (_NASDAQ_LISTED_URL, "Symbol"),
        # CQS Symbol 最接近现有行情接口使用的代码：类别股用点号、
        # 优先股用小写 p、权证用 .WS；下方统一转换为下划线形式。
        (_NASDAQ_OTHER_LISTED_URL, "CQS Symbol"),
    )
    for url, symbol_field in sources:
        response = requests.get(
            url,
            headers=_NASDAQ_HEADERS,
            timeout=(15, 60),
        )
        response.raise_for_status()
        text = response.content.decode("utf-8-sig")
        for item in csv.DictReader(StringIO(text), delimiter="|"):
            if str(item.get("Test Issue") or "").strip().upper() != "N":
                continue
            symbol = str(item.get(symbol_field) or "").strip()
            name = str(item.get("Security Name") or "").strip()
            if not symbol or not name or symbol.startswith("File Creation Time"):
                continue
            code = symbol.replace(".", "_").replace("p", "_").lower()
            official_rows.append((code, name))

    official_codes = {code for code, _ in official_rows}
    cache = _load_json(cache_path)
    aliases = _clean_us_cn_aliases(cache.get("aliases"))
    current_month = today.strftime("%Y-%m")
    cached_month = str(cache.get("last_update") or "")[:7]
    alias_status = {
        "last_attempt": str(cache.get("last_update") or ""),
        "last_checked": str(cache.get("last_update") or ""),
        "last_update": str(cache.get("last_update") or ""),
        "updated": False, "error": False,
    }
    if cached_month != current_month or not aliases:
        alias_status["last_attempt"] = today.isoformat()
        try:
            refreshed = _fetch_us_cn_aliases(official_codes)
            new_aliases = {
                code: name for code, name in refreshed.items()
                if code in official_codes
            }
            if not new_aliases:
                raise ValueError("新浪中文别名与 Nasdaq 官方代码无匹配")
            _save_json(
                cache_path,
                {"last_update": today.isoformat(), "aliases": new_aliases},
            )
            aliases = new_aliases
            alias_status.update(
                last_checked=today.isoformat(), last_update=today.isoformat(), updated=True,
            )
            print(f"美股中文别名: 新浪更新 {len(aliases)} 条", flush=True)
        except Exception as exc:
            alias_status.update(error=True, error_message=safe_error(exc))
            print(
                f"::warning::美股中文别名更新失败，继续使用缓存: {safe_error(exc)}",
                flush=True,
            )
    else:
        aliases = {
            code: name for code, name in aliases.items()
            if code in official_codes
        }
        print(f"美股中文别名: 使用缓存 {len(aliases)} 条", flush=True)

    frame = _rows_frame([
        (code, aliases.get(code, name), name, "美", "us") for code, name in official_rows
    ])
    frame.attrs["us_cn_aliases"] = alias_status
    return frame


# ----------------- 分类任务 -----------------

def _tasks(output_dir: str = OUTPUT_DIR) -> list[tuple]:
    """每个任务负责一个 JSON；耗时较长的任务优先提交到线程池。"""
    return [
        ("stock_us.json", "美股", _stock_us_name_code, (), {"output_dir": output_dir}),
        ("stock_hk.json", "港股", _stock_hk_name_code, (), {}),
        ("futures_sh.json", "上期所期货", futures_info_all, (), {}),
        ("stock_sh.json", "沪市股票", _stock_sh_all, (), {}),
        ("stock_sz.json", "深市股票", _stock_sz_all, (), {}),
        ("fund_cn.json", "沪深基金", _fund_exchange_all, (), {}),
        ("stock_bj.json", "北交所股票", _stock_bj_name_code, (), {}),
    ]


def _run_task(task: tuple) -> tuple[pd.DataFrame | None, str | None]:
    """分类间隔离；临时错误重试，已按页重试的任务不再整类重跑。"""
    _, label, func, args, kwargs = task
    last_error = None
    for attempt in range(1, TASK_ATTEMPTS + 1):
        try:
            result = func(*args, **kwargs)
            if result.empty:
                raise ValueError("返回列表为空")
            print(f"{label}更新完成: {len(result)} 条", flush=True)
            return result, None
        except Exception as exc:
            last_error = exc
            if attempt < TASK_ATTEMPTS and retryable_error(exc):
                delay = TASK_RETRY_SECONDS[attempt - 1]
                print(
                    f"{label}第 {attempt} 次拉取失败，{delay} 秒后重试: {safe_error(exc)}",
                    flush=True,
                )
                time.sleep(delay)
            else:
                break
    message = safe_error(last_error)
    print(f"::warning::{label}拉取失败，保留旧文件: {message}")
    return None, message


def _run_tasks(tasks: list[tuple]) -> dict[str, tuple[pd.DataFrame | None, str | None]]:
    """并发执行分类任务，返回每个文件的最终结果。"""
    results = {}
    with ThreadPoolExecutor(max_workers=min(6, len(tasks))) as pool:
        pending = {pool.submit(_run_task, task): task for task in tasks}
        for future in as_completed(pending):
            filename = pending[future][0]
            try:
                results[filename] = future.result()
            except Exception as exc:
                results[filename] = (None, safe_error(exc))
    return results


# ----------------- 上期所期货 -----------------

def _futures_nodes() -> list[str]:
    """从新浪节点脚本中提取上期所及上期能源节点名。"""
    url = "https://vip.stock.finance.sina.com.cn/quotes_service/view/js/qihuohangqing.js"
    r = requests.get(url, headers=_SINA_HEADERS, timeout=15)
    r.raise_for_status()
    r.encoding = "gb2312"
    match = re.search(r"\bshfe\s*:\s*\[(.*?)\]\s*,\s*cffex\s*:", r.text, re.DOTALL)
    if not match:
        raise ValueError("新浪期货节点脚本中缺少 shfe 列表")
    return re.findall(r"\[\s*'[^']*'\s*,\s*'([^']+)'", match.group(1))

def _stock_shfe_futures() -> pd.DataFrame:
    """上期所全部期货合约（含上期能源；新浪 getHQFuturesData 按品种遍历）。"""
    rows = []
    for node in _futures_nodes():
        for page in range(1, 4):
            r = requests.get(
                "https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/Market_Center.getHQFuturesData",
                params={"page": page, "sort": "position", "asc": "0", "node": node, "base": "futures"},
                headers=_SINA_HEADERS,
                timeout=8,
            )
            r.raise_for_status()
            contracts = r.json()
            if not isinstance(contracts, list):
                raise ValueError("新浪期货接口返回格式错误")
            if not contracts:
                break
            for contract in contracts:
                symbol = str(contract.get("symbol", "") or "").strip().lower()
                name = str(contract.get("name", "") or "").strip()
                if symbol:
                    rows.append((symbol, name))
            if len(contracts) < 20:
                break
    seen, uniq = set(), []
    for c, n in rows:
        if c not in seen:
            seen.add(c)
            uniq.append((c, n))
    rows = [(code, name, "", "期", "sh") for code, name in uniq]
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
        # 期货沿用原有合约键，避免更新代码表后丢失已有自选项。
        codes[code if mtype == "期" else market + code] = entry
    return dict(sorted(codes.items(), key=_code_sort_key))


# ---------------- 拉取市场列表并保存 ----------------

def _load_json(path: str) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def _save_json(path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
        file.write("\n")
    os.replace(tmp, path)


def _iso_now() -> str:
    return datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")


def _file_status_on_failure(previous: dict, existing: dict, run_date: str, error: str) -> dict:
    return {
        "last_checked": str(previous.get("last_checked") or ""),
        "last_update": str(existing.get("last_update") or previous.get("last_update") or ""),
        "updated": False,
        "error": True,
        "last_attempt": run_date,
        "error_message": error,
    }


def update_code_files(output_dir: str = OUTPUT_DIR) -> dict:
    """更新所有分类并写状态文件；某一分类失败时保留其现有 JSON。"""
    started_at = _iso_now()
    run_date = started_at[:10]
    status_path = os.path.join(output_dir, STATUS_FILE)
    previous_report = _load_json(status_path)
    previous_status = previous_report.get("files", {})
    tasks = _tasks(output_dir=output_dir)
    results = _run_tasks(tasks)
    file_states = {}
    auxiliary = dict(previous_report.get("auxiliary", {}))

    for filename, _, _, _, _ in tasks:
        path = os.path.join(output_dir, filename)
        existing = _load_json(path)
        frame, error = results.get(filename, (None, "任务未返回结果"))
        previous = previous_status.get(filename, {})
        if frame is None or error:
            file_states[filename] = _file_status_on_failure(previous, existing, run_date, error)
            continue

        if filename == "stock_us.json" and "us_cn_aliases" in frame.attrs:
            auxiliary[US_CN_ALIAS_CACHE_FILE] = frame.attrs["us_cn_aliases"]
        codes = _df_to_dict(frame)
        if not codes:
            file_states[filename] = _file_status_on_failure(previous, existing, run_date, "返回列表为空")
            continue

        changed = existing.get("codes") != codes or not existing.get("last_update")
        last_update = run_date if changed else str(existing.get("last_update") or run_date)
        if changed:
            _save_json(path, {"last_update": last_update, "codes": codes})
        file_states[filename] = {
            "last_checked": run_date,
            "last_update": last_update,
            "updated": changed,
            "error": False,
            "last_attempt": run_date,
        }
        action = "已更新" if changed else "无变化"
        print(f"{filename}: {len(codes)} 条，{action}", flush=True)

    for filename in MANUAL_FILES:
        existing = _load_json(os.path.join(output_dir, filename))
        if (isinstance(existing.get("codes"), dict) and existing["codes"]
                and existing.get("last_update")):
            file_states[filename] = {
                "last_checked": run_date, "last_update": existing["last_update"],
                "updated": False, "error": False,
            }
        else:
            file_states[filename] = _file_status_on_failure(
                previous_status.get(filename, {}), existing, run_date,
                "手工维护的指数 JSON 缺失或无效，请检查 codes-data 分支",
            )

    status = {
        "run_date": run_date,
        "started_at": started_at,
        "completed_at": _iso_now(),
        "files": file_states,
        "auxiliary": auxiliary,
    }
    failed = [name for name, state in file_states.items() if state["error"]]
    failed.extend(name for name, state in auxiliary.items() if state.get("error"))
    status["has_errors"] = bool(failed)
    _save_json(status_path, status)
    if failed:
        print(f"本次失败分类: {', '.join(failed)}；其他文件已正常处理", flush=True)
    return status


def main() -> int:
    try:
        update_code_files()
        return 0
    except Exception as exc:
        print(f"::error::代码列表更新流程失败: {safe_error(exc)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
