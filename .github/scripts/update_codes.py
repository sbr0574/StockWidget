"""GitHub Action 专用：拉取全市场代码并写入 resources/ 下三个 JSON。

本脚本由 .github/workflows/update-codes.yml 调用，保持自包含，不依赖项目运行时包。
数据源优先使用东方财富 clist 接口（带 UA/Referer、超时与重试），
避免直接访问沪深交易所官网或新浪无浏览器头的接口，降低被反爬/重置连接的概率。
"""
import json
import os
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
import requests

import akshare as ak
from pypinyin import Style, pinyin

# 脚本位于 <项目根>/.github/scripts/，向上三层即项目根
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 全市场代码列表文件名（对应三个独立 JSON）
CN_FILE = "stock_codes_list.json"          # 沪深京个股、基金、国内指数、港股及港股指数
GLOBAL_FILE = "stock_codes_global.json"    # 美股个股、全球主要指数（gb 前缀）
FUTURES_FILE = "stock_codes_futures.json"  # 上期所期货

_EM_CLIST_URL = "https://push2delay.eastmoney.com/api/qt/clist/get"
_EM_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://quote.eastmoney.com/center/gridlist.html",
}
_SINA_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://vip.stock.finance.sina.com.cn/",
}


def _env_int(name: str, default: int) -> int:
    try:
        value = int(os.getenv(name, str(default)))
    except ValueError:
        return default
    return max(1, value)


_EM_PAGE_SIZE = _env_int("UPDATE_CODES_EM_PAGE_SIZE", 500)
_FETCH_WORKERS = _env_int("UPDATE_CODES_WORKERS", 6)
_MARKET_ORDER = {"sh": 0, "sz": 1, "bj": 2, "hk": 3, "us": 4, "gb": 5, "": 6}


def _to_halfwidth(text: str) -> str:
    """全角字符转半角（全角字母/数字/常用符号），避免全角字母匹配不上搜索。"""
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
    text = _to_halfwidth(name).strip()           # 先转半角（处理全角字母 S/W/R/A/B/I 等）
    if not text:
        return "", ""
    text = text.replace(" ", "")
    py_full = "".join(x[0] for x in pinyin(text, style=Style.NORMAL, strict=False))
    py_abbr = "".join(x[0] for x in pinyin(text, style=Style.FIRST_LETTER, strict=False))
    return py_full.lower(), py_abbr.lower()


def _has_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in str(text or ""))


def _normalize_code(code: str, market: str) -> str:
    code = str(code or "").strip().lower()
    market = str(market or "").strip().lower()
    if market in {"sh", "sz", "bj"} and code.isdigit():
        return code.zfill(6)
    if market == "hk" and code.isdigit():
        return code.zfill(5)
    return code


def _code_sort_key(item: tuple[str, dict]) -> tuple:
    key, entry = item
    market = str(entry.get("market", "") or "").strip().lower()
    code = str(entry.get("code", "") or "").strip().lower()
    return (_MARKET_ORDER.get(market, 99), market, code, key)


def _df_to_codes(df: pd.DataFrame) -> dict:
    """把 [code,name,name_en,type,market] 的 df 转成有序 codes 字典。"""
    codes = {}
    for _, row in df.iterrows():
        market = str(row.get("market", "") or "").strip().lower()
        code = _normalize_code(row.get("code", ""), market)
        if not code:
            continue
        name = _to_halfwidth(str(row.get("name", "") or "")).strip()
        name_en = _to_halfwidth(str(row.get("name_en", "") or row.get("engname", "") or "")).strip()
        if not name and name_en:
            name = name_en
        if name and not name_en and not _has_cjk(name):
            name_en = name
        mtype = str(row.get("type", "") or "").strip()
        py_full, py_abbr = _name_pinyin(name) if _has_cjk(name) else ("", "")
        entry = {
            "code": code,
            "type": mtype,
            "market": market,
            "name": name,
            "py": py_full,
            "abbr": py_abbr,
            "name_en": name_en,
        }
        codes[market + code] = entry
    return dict(sorted(codes.items(), key=_code_sort_key))


def _em_clist_all(fs: str, fid: str = "f12", fields: str = "f12,f14") -> list[dict]:
    """东财 clist 分页拉全市场原始 dict 列表；单页失败重试 3 次，断连返回已收集部分。"""
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
                        "pn": page, "pz": _EM_PAGE_SIZE, "po": 1, "np": 1, "fltt": 2, "invt": 2,
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
        if total is None and len(diff) < _EM_PAGE_SIZE:
            break
        page += 1
    return rows


def _market_from_fs(fs: str) -> str:
    has_sh = "m:1" in fs
    has_sz = "m:0" in fs
    if has_sh and not has_sz:
        return "sh"
    if has_sz and not has_sh:
        return "sz"
    return ""


def _market_from_code(code: str) -> str:
    code = str(code or "").strip()
    if code.startswith(("43", "83", "87", "88", "92")):
        return "bj"
    if code.startswith(("5", "6", "9")):
        return "sh"
    if code.startswith(("0", "1", "2", "3")):
        return "sz"
    return ""


def _market_from_em(market_id: str) -> str:
    if market_id == "1":
        return "sh"
    if market_id == "0":
        return "sz"
    return ""


def _em_stock_df(fs: str, mtype: str, market: str | None = None) -> pd.DataFrame:
    """从东财 clist 拉取一个分类，返回 [证券代码, 证券简称, 市场, 类型]。

    market 用于修正东财 f13 无法表达的市场（例如北交所 f13 返回 0，但实际市场为 bj）。
    """
    rows = []
    default_market = str(market or "").strip().lower() or _market_from_fs(fs)
    for d in _em_clist_all(fs, fields="f12,f13,f14"):
        code = str(d.get("f12") or "").strip()
        name = str(d.get("f14") or "").strip()
        market_id = str(d.get("f13") or "").strip()
        if not code or name.startswith("退市"):
            continue
        row_market = default_market or _market_from_em(market_id) or _market_from_code(code)
        rows.append((code, name, row_market))
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称", "市场"])
    df["类型"] = mtype
    return df.drop_duplicates(subset=["市场", "证券代码"]).reset_index(drop=True)


def _concat_dedup(frames: list[pd.DataFrame]) -> pd.DataFrame:
    df = pd.concat(frames, ignore_index=True)
    subset = [c for c in ("市场", "证券代码") if c in df.columns]
    if len(subset) < 2:
        subset = ["证券代码"] if "证券代码" in df.columns else None
    return df.drop_duplicates(subset=subset).reset_index(drop=True)


def _run_frame_tasks(tasks: list[tuple], tick) -> list[pd.DataFrame]:
    """并发拉取互不依赖的分类，按 tasks 原顺序返回，保证最终 JSON 稳定。"""
    if not tasks:
        return []
    results: list[pd.DataFrame | None] = [None] * len(tasks)
    workers = min(_FETCH_WORKERS, len(tasks))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {}
        for idx, task in enumerate(tasks):
            func, args, kwargs = task
            futures[pool.submit(func, *args, **kwargs)] = idx
        for future in as_completed(futures):
            idx = futures[future]
            try:
                results[idx] = future.result()
            except Exception:
                traceback.print_exc()
                results[idx] = pd.DataFrame()
            tick()
    return [df if df is not None else pd.DataFrame() for df in results]


def _fund_etf_em() -> pd.DataFrame:
    """东财 ETF 列表（合并两个 ETF 分类，覆盖股票/债券/货币/跨境/黄金等 ETF）。"""
    return _concat_dedup([
        _em_stock_df("b:MK0021,b:MK0022,b:MK0023,b:MK0024,b:MK0827", "基"),
        _em_stock_df("b:MK0400,b:MK0401,b:MK0402,b:MK0403", "基"),
    ])


def _fund_lof_em() -> pd.DataFrame:
    """东财 LOF 列表。"""
    return _em_stock_df("b:MK0404,b:MK0405,b:MK0406,b:MK0407,b:MK0408", "基")


def _fund_close_sina() -> pd.DataFrame:
    """新浪传统封闭式基金列表；被限流时返回空表，不阻塞整体更新。"""
    url = (
        "https://vip.stock.finance.sina.com.cn/quotes_service/api/jsonp.php/"
        "IO.XSRV2.CallbackList['da_yPT46_Ll7K6WD']/Market_Center.getHQNodeDataSimple"
    )
    params = {
        "page": "1", "num": "5000", "sort": "symbol", "asc": "0",
        "node": "close_fund", "[object HTMLDivElement]": "qvvne",
    }
    headers = {
        **_SINA_HEADERS,
        "Referer": "https://vip.stock.finance.sina.com.cn/fund_center/index.html",
    }
    columns = ["证券代码", "证券简称", "市场"]
    rows = []
    try:
        r = requests.get(url, params=params, headers=headers, timeout=15)
        text = r.text or ""
        if r.status_code != 200 or text.lstrip().startswith("<") or "拒绝访问" in text:
            return pd.DataFrame(rows, columns=columns)
        from akshare.utils import demjson
        payload = text[text.find("([") + 1 : -2]
        data = demjson.decode(payload)
        if not isinstance(data, list):
            return pd.DataFrame(rows, columns=columns)
        for d in data:
            symbol = str(d.get("symbol") or "").strip()
            name = str(d.get("name") or "").strip()
            if len(symbol) < 3:
                continue
            market, code = symbol[:2].lower(), symbol[2:]
            if market in ("sh", "sz", "bj") and code:
                rows.append((code, name, market))
    except Exception:
        # 新浪接口不稳定或已限流，回退为东财 REITs 分类（见 _fund_close_em）
        pass
    return pd.DataFrame(rows, columns=columns)


def _fund_close_em() -> pd.DataFrame:
    """封闭式基金：东财 REITs + 新浪传统封闭基金（新浪失败时保留 REITs）。"""
    df = _concat_dedup([
        _em_stock_df("m:1+t:9+e:97,m:0+t:10+e:97", "基"),
        _fund_close_sina(),
    ])
    df["类型"] = "基"
    return df


def _index_cn_em() -> pd.DataFrame:
    """东财沪深指数列表。"""
    return _concat_dedup([
        _em_stock_df("m:1+t:1", "指"),
        _em_stock_df("m:0+t:5", "指"),
    ])


def _stock_hk_name_code() -> pd.DataFrame:
    """港股全部股票（新浪 getHKStockData 分页，单页重试；含英文名）"""
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
                rows.append((code, name, eng))
        if len(diff) < 60:
            break
        page += 1
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称", "英文名称"])
    df["类型"] = "港"
    df["市场"] = "hk"
    return df


def _stock_hk_index_name_code() -> pd.DataFrame:
    """港股指数（新浪 rt_hk 前缀；过滤上证/沪深指数避免与国内指数重复）"""
    rows = []
    try:
        df = ak.stock_hk_index_spot_sina()[["代码", "名称"]]
        for _, row in df.iterrows():
            code = str(row.get("代码", "") or "").strip().lower()
            name = str(row.get("名称", "") or "").strip()
            if not code or code.startswith(("sse", "csi")):
                continue
            rows.append((code, name))
    except Exception:
        rows = []
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称"])
    df["类型"] = "指"
    df["市场"] = "hk"
    return df


def _is_ascii(text: str) -> bool:
    return bool(text) and all(ord(c) < 128 for c in str(text).strip())


def _stock_us_name_code() -> pd.DataFrame:
    """美股全部股票（东财延迟主机 clist，按总市值排序使热门股在前，附带英文名）。"""
    rows = []
    for d in _em_clist_all("m:105,m:106,m:107", fid="f20"):
        code = str(d.get("f12") or "").strip().lower()
        name = str(d.get("f14") or "").strip()
        if not code:
            continue
        eng = name if _is_ascii(name) else ""
        rows.append((code, name, eng))
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称", "英文名称"])
    df["类型"] = "美"
    df["市场"] = "us"
    return df


def _stock_global_index_name_code() -> pd.DataFrame:
    """全球主要指数（新浪 b_ 前缀国际指数：日经/KOSPI/DAX 等）"""
    rows = []
    try:
        df = ak.index_global_name_table()   # 列: 指数名称, 代码
        for _, row in df.iterrows():
            code = str(row.get("代码", "") or "").strip().lower()
            name = str(row.get("指数名称", "") or "").strip()
            if code:
                rows.append((code, name))
    except Exception:
        rows = []
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称"])
    df["类型"] = "指"
    df["市场"] = "gb"
    return df


def _stock_us_index_name_code() -> pd.DataFrame:
    """美股主要指数（新浪 gb_ 前缀）"""
    rows = [
        ("ixic", "纳斯达克综合指数"),
        ("dji", "道琼斯工业平均指数"),
        ("inx", "标普500指数"),
        ("ndx", "纳斯达克100指数"),
    ]
    df = pd.DataFrame(rows, columns=["证券代码", "证券简称"])
    df["类型"] = "指"
    df["市场"] = "us"
    return df


def _futures_symbol_mark() -> list[tuple[str, str]]:
    """上期所品种与新浪行情节点映射（带超时，避免 akshare 原接口无限等待）。"""
    url = "https://vip.stock.finance.sina.com.cn/quotes_service/view/js/qihuohangqing.js"
    r = requests.get(url, headers=_SINA_HEADERS, timeout=15)
    r.encoding = "gb2312"
    text = r.text
    raw = text[text.find("{") : text.find("}") + 1]
    from akshare.utils import demjson
    data = demjson.decode(raw)
    shfe = data.get("shfe", [])
    return [(str(item[0]).strip(), str(item[1]).strip()) for item in shfe[1:] if len(item) >= 2]


def _stock_shfe_futures() -> pd.DataFrame:
    """上期所全部期货合约（含上期能源；新浪 getHQFuturesData 按品种遍历）。"""
    rows = []
    try:
        for mark in _futures_symbol_mark():
            for page in range(1, 4):
                try:
                    r = requests.get(
                        "https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/Market_Center.getHQFuturesData",
                        params={"page": page, "sort": "position", "asc": "0", "node": mark, "base": "futures"},
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
    df = pd.DataFrame(uniq, columns=["证券代码", "证券简称"])
    df["类型"] = "期"
    df["市场"] = ""
    return df


def stock_info_cn(tick=None) -> pd.DataFrame:
    """沪深京个股、基金、国内指数、港股及港股指数（列: code,name,name_en,type,market）"""
    tick = tick or (lambda: None)
    tasks = [
        (_em_stock_df, ("m:1+t:2", "沪"), {}),                         # 沪主板 A
        (_em_stock_df, ("m:1+t:3", "沪"), {}),                         # 沪 B
        (_em_stock_df, ("m:1+t:23", "沪"), {}),                        # 科创板
        (_em_stock_df, ("m:0+t:6", "深"), {}),                         # 深主板 A
        (_em_stock_df, ("m:0+t:80", "深"), {}),                        # 创业板
        (_em_stock_df, ("m:0+t:7", "深"), {}),                         # 深 B
        (_em_stock_df, ("m:0+t:81+s:2048", "京"), {"market": "bj"}),   # 北交所
        (_fund_etf_em, (), {}),
        (_fund_lof_em, (), {}),
        (_fund_close_em, (), {}),
        (_index_cn_em, (), {}),
        (_stock_hk_name_code, (), {}),
        (_stock_hk_index_name_code, (), {}),
    ]
    frames = _run_frame_tasks(tasks, tick)

    for df in frames:
        if "英文名称" not in df.columns:
            df["英文名称"] = ""
    df = pd.concat(frames, ignore_index=True)
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "name_en", "类型": "type", "市场": "market"}, inplace=True)
    return df


def stock_info_global(tick=None) -> pd.DataFrame:
    """美股个股 + 全球主要指数（列: code,name,name_en,type,market）"""
    tick = tick or (lambda: None)
    tasks = [
        (_stock_us_name_code, (), {}),
        (_stock_global_index_name_code, (), {}),
        (_stock_us_index_name_code, (), {}),
    ]
    frames = _run_frame_tasks(tasks, tick)
    for df in frames:
        if "英文名称" not in df.columns:
            df["英文名称"] = ""
    df = pd.concat(frames, ignore_index=True)
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "name_en", "类型": "type", "市场": "market"}, inplace=True)
    return df


def stock_info_futures(tick=None) -> pd.DataFrame:
    """上期所期货（列: code,name,name_en,type,market）"""
    if tick:
        tick()
    df = _stock_shfe_futures()
    if "英文名称" not in df.columns:
        df["英文名称"] = ""
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "name_en", "类型": "type", "市场": "market"}, inplace=True)
    return df


def fetch_codes_groups(progress_cb=None) -> dict[str, dict]:
    """拉取三组代码，返回 {文件名: {"last_update": "YYYY-MM-DD", "codes": {...}}}。"""
    total_steps = 17   # cn 13 + global 3 + futures 1
    done = 0

    def _tick():
        nonlocal done
        done += 1
        if callable(progress_cb):
            progress_cb(round(done * 100 / total_steps))

    now = datetime.now().strftime("%Y-%m-%d")
    return {
        CN_FILE: {"last_update": now, "codes": _df_to_codes(stock_info_cn(_tick))},
        GLOBAL_FILE: {"last_update": now, "codes": _df_to_codes(stock_info_global(_tick))},
        FUTURES_FILE: {"last_update": now, "codes": _df_to_codes(stock_info_futures(_tick))},
    }


def write_codes_groups(target_dir: str, progress_cb=None) -> dict[str, dict] | None:
    """拉取三组代码并写三个 json 到 target_dir；失败返回 None。"""
    try:
        groups = fetch_codes_groups(progress_cb)
    except Exception as exc:
        traceback.print_exc()
        print(f"::error::拉取代码列表失败: {exc}")
        return None
    os.makedirs(target_dir, exist_ok=True)
    for fname, data in groups.items():
        path = os.path.join(target_dir, fname)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    return groups


def main() -> int:
    resources_dir = os.path.join(ROOT, "resources")
    groups = write_codes_groups(resources_dir)
    if groups is None:
        print("::error::拉取代码列表失败")
        return 1
    for fname, data in groups.items():
        print(f"更新 {fname}: {len(data.get('codes', {}))} 条")
    print("完成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
