import json
import os
from datetime import datetime
import pandas as pd
import requests

import akshare as ak
from pypinyin import Style, pinyin

# 全市场代码列表文件名（对应三个独立 JSON）
CN_FILE = "stock_codes_list.json"          # 沪深京个股、基金、国内指数、港股及港股指数
GLOBAL_FILE = "stock_codes_global.json"    # 美股个股、全球主要指数
FUTURES_FILE = "stock_codes_futures.json"  # 上期所期货

_EM_CLIST_URL = "https://push2delay.eastmoney.com/api/qt/clist/get"


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


def _df_to_codes(df: pd.DataFrame) -> dict:
    """把 [code,name,engname,type,market] 的 df 转成 codes 字典（含拼音/缩写/英文名）。"""
    codes = {}
    for _, row in df.iterrows():
        code = str(row.get("code", "") or "").strip()
        if not code:
            continue
        name = _to_halfwidth(str(row.get("name", "") or "")).strip()
        engname = _to_halfwidth(str(row.get("engname", "") or "")).strip()
        market = str(row.get("market", "") or "").strip()
        mtype = str(row.get("type", "") or "").strip()
        py_full, py_abbr = _name_pinyin(name)
        entry = {
            "code": code,
            "type": mtype,
            "market": market,
            "name": name,
            "py": py_full,
            "abbr": py_abbr,
        }
        if engname:
            entry["engname"] = engname
        codes[market + code] = entry
    return codes


def _em_clist_all(fs: str, fid: str = "f12", fields: str = "f12,f14") -> list[dict]:
    """东财 clist 分页拉全市场原始 dict 列表；单页失败重试 3 次，断连返回已收集部分。"""
    rows = []
    page = 1
    while True:
        diff = None
        for _ in range(3):
            try:
                r = requests.get(
                    _EM_CLIST_URL,
                    params={
                        "pn": page, "pz": 100, "po": 1, "np": 1, "fltt": 2, "invt": 2,
                        "fid": fid, "fs": fs, "fields": fields,
                    },
                    timeout=10,
                )
                diff = ((r.json() or {}).get("data") or {}).get("diff") or []
                break
            except Exception:
                diff = None
        if not diff:
            break
        rows.extend(diff)
        if len(diff) < 100:
            break
        page += 1
    return rows


def _stock_hk_name_code() -> pd.DataFrame:
    """港股全部股票（新浪 getHKStockData 分页，单页重试，无进度条；含英文名）"""
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
    """美股全部股票（东财延迟主机 clist，按总市值排序使热门股在前，附带英文名）。
    英文名仅取数据源提供的真实名称：无中文名（f14 为英文）时作为 engname，不人工映射。"""
    rows = []
    for d in _em_clist_all("m:105,m:106,m:107", fid="f20"):
        code = str(d.get("f12") or "").strip().lower()
        name = str(d.get("f14") or "").strip()
        if not code:
            continue
        eng = name if _is_ascii(name) else ""   # 有中文名则没有英文名可存
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
    df["市场"] = "g"
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


def _stock_shfe_futures() -> pd.DataFrame:
    """上期所全部期货合约（含上期能源；新浪 getHQFuturesData 按品种遍历，与主数据源同源）。
    主力连续 AU0 -> au0，具体合约 AU2612 -> au2612。"""
    rows = []
    try:
        marks = []
        for _, row in ak.futures_symbol_mark().iterrows():
            if str(row.get("exchange", "") or "") == "上海期货交易所":
                mark = str(row.get("mark", "") or "").strip()
                if mark:
                    marks.append(mark)
        for mark in marks:
            for page in range(1, 4):
                try:
                    r = requests.get(
                        "https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/Market_Center.getHQFuturesData",
                        params={"page": page, "sort": "position", "asc": "0", "node": mark, "base": "futures"},
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
    # 去重
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
    """沪深京个股、基金、国内指数、港股及港股指数（列: code,name,engname,type,market）"""
    frames = []
    tick = tick or (lambda: None)

    stock_sha = ak.stock_info_sh_name_code(symbol="主板A股")[["证券代码", "证券简称"]]
    stock_sha["类型"] = "沪"
    stock_sha["市场"] = "sh"
    frames.append(stock_sha)
    tick()

    stock_shb = ak.stock_info_sh_name_code(symbol="主板B股")[["证券代码", "证券简称"]]
    stock_shb["类型"] = "沪"
    stock_shb["市场"] = "sh"
    frames.append(stock_shb)
    tick()

    stock_kcb = ak.stock_info_sh_name_code(symbol="科创板")[["证券代码", "证券简称"]]
    stock_kcb["类型"] = "科"
    stock_kcb["市场"] = "sh"
    frames.append(stock_kcb)
    tick()

    stock_sza = ak.stock_info_sz_name_code(symbol="A股列表")
    stock_sza["A股代码"] = stock_sza["A股代码"].astype(str).str.zfill(6)
    stock_sza["类型"] = stock_sza["板块"].map({"主板": "深", "创业板": "创"})
    stock_sza = stock_sza[["A股代码", "A股简称", "类型"]]
    stock_sza["市场"] = "sz"
    stock_sza.rename(columns={"A股代码":"证券代码","A股简称":"证券简称"},inplace=True)
    frames.append(stock_sza)
    tick()

    stock_szb = ak.stock_info_sz_name_code(symbol="B股列表")[["B股代码", "B股简称"]]
    stock_szb.rename(columns={"B股代码":"证券代码","B股简称":"证券简称"},inplace=True)
    stock_szb["类型"] = "深"
    stock_szb["市场"] = "sz"
    frames.append(stock_szb)
    tick()

    stock_bj = _stock_info_bj_name_code()[["证券代码", "证券简称"]]
    stock_bj["类型"] = "京"
    stock_bj["市场"] = "bj"
    frames.append(stock_bj)
    tick()

    stock_etf = ak.fund_etf_category_sina(symbol="ETF基金")[["代码", "名称"]]
    stock_etf["市场"] = stock_etf["代码"].str[:2]
    stock_etf["代码"] = stock_etf["代码"].str[2:]
    stock_etf.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_etf["类型"] = "基"
    frames.append(stock_etf)
    tick()

    stock_lof = ak.fund_etf_category_sina(symbol="LOF基金")[["代码", "名称"]]
    stock_lof["市场"] = stock_lof["代码"].str[:2]
    stock_lof["代码"] = stock_lof["代码"].str[2:]
    stock_lof.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_lof["类型"] = "基"
    frames.append(stock_lof)
    tick()

    stock_closefund = ak.fund_etf_category_sina(symbol="封闭式基金")[["代码", "名称"]]
    stock_closefund["市场"] = stock_closefund["代码"].str[:2]
    stock_closefund["代码"] = stock_closefund["代码"].str[2:]
    stock_closefund.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_closefund["类型"] = "基"
    frames.append(stock_closefund)
    tick()

    index_stock = ak.index_stock_info()[["index_code","display_name"]]
    index_stock['市场'] = index_stock['index_code'].astype(str).str.zfill(6).str[0].map({'0': 'sh', '3': 'sz'})
    index_stock.rename(columns={"index_code":"证券代码","display_name":"证券简称"},inplace=True)
    index_stock["类型"] = "指"
    frames.append(index_stock)
    tick()

    # 港股 + 港股指数
    frames.append(_stock_hk_name_code())
    tick()
    frames.append(_stock_hk_index_name_code())
    tick()

    for df in frames:
        if "英文名称" not in df.columns:
            df["英文名称"] = ""
    df = pd.concat(frames, ignore_index=True)
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "engname", "类型": "type", "市场": "market"}, inplace=True)
    return df


def stock_info_global(tick=None) -> pd.DataFrame:
    """美股个股 + 全球主要指数（列: code,name,engname,type,market）"""
    frames = []
    tick = tick or (lambda: None)
    frames.append(_stock_us_name_code())
    tick()
    frames.append(_stock_global_index_name_code())
    tick()
    frames.append(_stock_us_index_name_code())
    tick()
    for df in frames:
        if "英文名称" not in df.columns:
            df["英文名称"] = ""
    df = pd.concat(frames, ignore_index=True)
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "engname", "类型": "type", "市场": "market"}, inplace=True)
    return df


def stock_info_futures(tick=None) -> pd.DataFrame:
    """上期所期货（列: code,name,engname,type,market）"""
    if tick:
        tick()
    df = _stock_shfe_futures()
    if "英文名称" not in df.columns:
        df["英文名称"] = ""
    df.rename(columns={"证券代码": "code", "证券简称": "name",
                       "英文名称": "engname", "类型": "type", "市场": "market"}, inplace=True)
    return df


def fetch_codes_groups(progress_cb=None) -> dict[str, dict]:
    """拉取三组代码，返回 {文件名: {"last_update": "YYYY-MM-DD", "codes": {...}}}。"""
    total_steps = 16   # cn 12 + global 3 + futures 1
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
    except Exception:
        return None
    os.makedirs(target_dir, exist_ok=True)
    for fname, data in groups.items():
        path = os.path.join(target_dir, fname)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    return groups


def _stock_info_bj_name_code() -> pd.DataFrame:
    """
    Akshare修改版, 去除进度条
    北京证券交易所-股票列表
    https://www.bse.cn/nq/listedcompany.html
    :return: 股票列表
    :rtype: pandas.DataFrame
    """
    url = "https://www.bse.cn/nqxxController/nqxxCnzq.do"
    payload = {
        "page": "0",
        "typejb": "T",
        "xxfcbj[]": "2",
        "xxzqdm": "",
        "sortfield": "xxzqdm",
        "sorttype": "asc",
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/110.0.0.0 Safari/537.36"
    }
    r = requests.post(url, data=payload, headers=headers)
    data_text = r.text
    data_json = json.loads(data_text[data_text.find("[") : -1])
    total_page = data_json[0]["totalPages"]
    big_df = pd.DataFrame()
    for page in range(total_page):
        payload.update({"page": page})
        r = requests.post(url, data=payload, headers=headers)
        data_text = r.text
        data_json = json.loads(data_text[data_text.find("[") : -1])
        temp_df = data_json[0]["content"]
        temp_df = pd.DataFrame(temp_df)
        big_df = pd.concat([big_df, temp_df], ignore_index=True)
    big_df.columns = [
        "上市日期",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "流通股本",
        "-",
        "-",
        "-",
        "-",
        "-",
        "所属行业",
        "-",
        "-",
        "-",
        "-",
        "报告日期",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "地区",
        "-",
        "-",
        "-",
        "-",
        "-",
        "券商",
        "总股本",
        "-",
        "证券代码",
        "-",
        "证券简称",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
    ]
    big_df = big_df[
        [
            "证券代码",
            "证券简称",
            "总股本",
            "流通股本",
            "上市日期",
            "所属行业",
            "地区",
            "报告日期",
        ]
    ]
    big_df["报告日期"] = pd.to_datetime(big_df["报告日期"], errors="coerce").dt.date
    big_df["上市日期"] = pd.to_datetime(big_df["上市日期"], errors="coerce").dt.date
    return big_df