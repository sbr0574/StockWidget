"""
GitHub Action 专用：拉取全市场代码并写入 resources/ 下 JSON
本脚本由 .github/workflows/update-codes.yml 调用, 保持自包含, 不依赖项目运行时包
数据源优先使用东方财富 clist 接口
"""
import json
import os
import re
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

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

def _has_cjk(text: str) -> bool:
    """判断字符串是否包含 CJK 中日韩字符。"""
    return any("\u4e00" <= ch <= "\u9fff" for ch in str(text or ""))

def _concat_dedup(frames: list[pd.DataFrame]) -> pd.DataFrame:
    df = pd.concat(frames, ignore_index=True)
    return df.drop_duplicates(subset=["market", "code"]).reset_index(drop=True)

def _is_active_em_security(item: dict, today: int) -> bool:
    """依据东财名称、上市日期和行情字段排除退市及待上市证券。"""
    name = str(item.get("f14") or "").strip()
    if name.startswith("退市") or name.endswith("退"):
        return False

    listing_date = str(item.get("f26") or "").strip()
    if listing_date.isdigit():
        listing_day = int(listing_date)
        if listing_day > today:
            return False
        if listing_day == today:
            return True

    missing_values = {None, "", "-", 0, 0.0}
    return not (item.get("f2") in missing_values and item.get("f18") in missing_values)


# ----------------- 并发拉取分类 -----------------

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
            print(f"{label}: {len(results[idx])} 条", flush=True)
    return [df if df is not None else pd.DataFrame(columns=_DF_COLUMNS) for df in results]

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

def _em_stock_df(fs: str, mtype: str, market: str, active_only: bool = False) -> pd.DataFrame:
    """从东财 clist 拉取一个分类的股票/基金/指数列表 返回df"""
    rows = []
    fields = "f12,f14,f2,f18,f26" if active_only else "f12,f14"
    today = int(datetime.now().strftime("%Y%m%d"))
    for d in _em_clist_all(fs, fields=fields):
        code = str(d.get("f12") or "").strip()
        name = str(d.get("f14") or "").strip()
        if not code or name.startswith("退市") or name.endswith("退"):
            continue
        if active_only and not _is_active_em_security(d, today):
            continue
        rows.append((code, name, "", mtype, market))
    df = pd.DataFrame(rows, columns=_DF_COLUMNS)
    return df.drop_duplicates(subset=["market", "code"]).reset_index(drop=True)

_ETF_BOARDS = (
    "MK0021", "MK0022", "MK0023", "MK0024", "MK0827",
    "MK0400", "MK0401", "MK0402", "MK0403",
)
_LOF_BOARDS = ("MK0404", "MK0405", "MK0406", "MK0407", "MK0408")

def _market_board_fs(market_id: int, boards: tuple[str, ...]) -> str:
    """生成带明确沪深市场约束的东财板块查询。"""
    return ",".join(f"m:{market_id}+b:{board}" for board in boards)

def _fund_etf_em() -> pd.DataFrame:
    """东财 ETF 列表，分别按沪深市场拉取。"""
    return _concat_dedup([
        _em_stock_df(_market_board_fs(1, _ETF_BOARDS),"基",market="sh",active_only=True,),
        _em_stock_df(_market_board_fs(0, _ETF_BOARDS),"基",market="sz",active_only=True,),
    ])

def _fund_lof_em() -> pd.DataFrame:
    """东财 LOF 列表，分别按沪深市场拉取。"""
    return _concat_dedup([
        _em_stock_df(_market_board_fs(1, _LOF_BOARDS),"基",market="sh",active_only=True,),
        _em_stock_df(_market_board_fs(0, _LOF_BOARDS),"基",market="sz",active_only=True,),
    ])

def _fund_close_em() -> pd.DataFrame:
    """分别拉取沪深封闭式基金，避免混合 fs 丢失深市 market。"""
    return _concat_dedup([
        _em_stock_df("m:1+t:9+e:97", "基", market="sh", active_only=True),
        _em_stock_df("m:0+t:10+e:97", "基", market="sz", active_only=True),
    ])

def _index_cn_em() -> pd.DataFrame:
    """东财沪深指数列表。"""
    return _concat_dedup([
        _em_stock_df("m:1+t:1", "指", market="sh"),
        _em_stock_df("m:0+t:5", "指", market="sz"),
    ])

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


def _tasks() -> list[tuple]:
    return [
        ("沪A", _em_stock_df, ("m:1+t:2", "沪"), {"market": "sh", "active_only": True}),
        ("科创板", _em_stock_df, ("m:1+t:23", "沪"), {"market": "sh", "active_only": True}),
        ("沪B", _em_stock_df, ("m:1+t:3", "沪"), {"market": "sh", "active_only": True}),
        ("深A", _em_stock_df, ("m:0+t:6", "深"), {"market": "sz", "active_only": True}),
        ("创业板", _em_stock_df, ("m:0+t:80", "深"), {"market": "sz", "active_only": True}),
        ("深B", _em_stock_df, ("m:0+t:7", "深"), {"market": "sz", "active_only": True}),
        ("京市", _em_stock_df, ("m:0+t:81+s:2048", "京"), {"market": "bj", "active_only": True}),
        ("ETF基金", _fund_etf_em, (), {}),
        ("LOF基金", _fund_lof_em, (), {}),
        ("封闭式基金", _fund_close_em, (), {}),
        ("国内指数", _index_cn_em, (), {}),
        ("港股", _stock_hk_name_code, (), {}),
        ("港股指数", _stock_hk_index_name_code, (), {}),
        ("美股", _stock_us_name_code, (), {}),
        ("全球股指", _stock_global_index_name_code, (), {}),
        ("美股指数", _stock_us_index_name_code, (), {}),
    ]


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

def futures_info_all() -> pd.DataFrame:
    """上期所期货"""
    df = _stock_shfe_futures()
    print(f"期货: {len(df)} 条", flush=True)
    return df

def stock_info_all() -> pd.DataFrame:
    """合并全部非期货证券"""
    tasks = _tasks()
    frames = _run_frame_tasks(tasks)
    by_label = {task[0]: frame for task, frame in zip(tasks, frames)}
    fund_total = len(_concat_dedup([
        by_label["ETF基金"], by_label["LOF基金"], by_label["封闭式基金"],
    ]))
    index_total = len(_concat_dedup([
        by_label["国内指数"], by_label["港股指数"],
        by_label["全球股指"], by_label["美股指数"],
    ]))
    print(f"基金合计: {fund_total} 条", flush=True)
    print(f"指数合计: {index_total} 条", flush=True)
    return pd.concat(frames, ignore_index=True)

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
