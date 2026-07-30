import json
from datetime import datetime
import pandas as pd
import requests

from resource import load_file, save_file
import akshare as ak
from pypinyin import Style, pinyin


def _name_pinyin(name: str) -> tuple[str, str]:
    text = str(name or "").strip()
    if not text:
        return "", ""
    text = text.replace(" ", "")
    py_full = "".join(x[0] for x in pinyin(text, style=Style.NORMAL, strict=False))
    py_abbr = "".join(x[0] for x in pinyin(text, style=Style.FIRST_LETTER, strict=False))
    return py_full.lower(), py_abbr.lower()


def refresh_index_from_akshare() -> list[dict]:

    df = stock_info_all()
    codes = {}
    for _, row in df.iterrows():
        code = str(row.get("code", "")).strip()
        name = str(row.get("name", "")).strip().replace("Ａ","A").replace("Ｂ","B")
        market = str(row.get("market", "")).strip()
        mtype = str(row.get("type", "")).strip()
        py_full, py_abbr = _name_pinyin(name)
        codes[code] = {
            "type": mtype,
            "market": market,
            "name": name,
            "py": py_full,
            "abbr": py_abbr,
        }
    
    return {"last_update": datetime.now().strftime("%Y-%m-%d"), "codes": codes}


def find_suggestions(codes: dict, text: str, limit: int = 20) -> list[dict]:
    q = str(text or "").strip().lower()
    if not q:
        return []

    scored = []
    for code, item in codes.items():
        name = str(item.get("name", ""))
        py = str(item.get("py", ""))
        abbr = str(item.get("abbr", ""))
        score = 0
        if code.startswith(q):
            score = 100
        elif q in code:
            score = 85
        elif name.startswith(q):
            score = 70
        elif q in name:
            score = 60
        elif py.startswith(q):
            score = 50
        elif q in py:
            score = 40
        elif abbr.startswith(q):
            score = 30
        elif q in abbr:
            score = 20

        if score > 0:
            scored.append((score, item))

    scored.sort(key=lambda pair: (-pair[0], pair[1].get("code", "")))
    return [item for _, item in scored[:limit]]


def stock_info_all() -> pd.DataFrame:
    stock_sha = ak.stock_info_sh_name_code(symbol="主板A股")[["证券代码", "证券简称"]]
    stock_sha["类型"] = "沪"
    stock_sha["市场"] = "sh"
    stock_shb = ak.stock_info_sh_name_code(symbol="主板B股")[["证券代码", "证券简称"]]
    stock_shb["类型"] = "沪"
    stock_shb["市场"] = "sh"
    stock_kcb = ak.stock_info_sh_name_code(symbol="科创板")[["证券代码", "证券简称"]]
    stock_kcb["类型"] = "科"
    stock_kcb["市场"] = "sh"

    stock_sza = ak.stock_info_sz_name_code(symbol="A股列表")
    stock_sza["A股代码"] = stock_sza["A股代码"].astype(str).str.zfill(6)
    stock_sza["类型"] = stock_sza["板块"].map({"主板": "深", "创业板": "创"})
    stock_sza = stock_sza[["A股代码", "A股简称", "类型"]]
    stock_sza["市场"] = "sz"
    stock_sza.rename(columns={"A股代码":"证券代码","A股简称":"证券简称"},inplace=True)
    stock_szb = ak.stock_info_sz_name_code(symbol="B股列表")[["B股代码", "B股简称"]]
    stock_szb.rename(columns={"B股代码":"证券代码","B股简称":"证券简称"},inplace=True)
    stock_szb["类型"] = "深"
    stock_szb["市场"] = "sz"

    stock_bj = _stock_info_bj_name_code()[["证券代码", "证券简称"]]
    stock_bj["类型"] = "京"
    stock_bj["市场"] = "bj"

    stock_etf = ak.fund_etf_category_sina(symbol="ETF基金")[["代码", "名称"]]
    stock_etf["市场"] = stock_etf["代码"].str[:2]
    stock_etf["代码"] = stock_etf["代码"].str[2:]
    stock_etf.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_etf["类型"] = "基"

    stock_lof = ak.fund_etf_category_sina(symbol="LOF基金")[["代码", "名称"]]
    stock_lof["市场"] = stock_lof["代码"].str[:2]
    stock_lof["代码"] = stock_lof["代码"].str[2:]
    stock_lof.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_lof["类型"] = "基"

    stock_closefund = ak.fund_etf_category_sina(symbol="封闭式基金")[["代码", "名称"]]
    stock_closefund["市场"] = stock_closefund["代码"].str[:2]
    stock_closefund["代码"] = stock_closefund["代码"].str[2:]
    stock_closefund.rename(columns={"代码":"证券代码","名称":"证券简称"},inplace=True)
    stock_closefund["类型"] = "基"

    index_stock = ak.index_stock_info()[["index_code","display_name"]]
    index_stock['市场'] = index_stock['index_code'].astype(str).str.zfill(6).str[0].map({'0': 'sh', '3': 'sz'})
    index_stock.rename(columns={"index_code":"证券代码","display_name":"证券简称"},inplace=True)
    index_stock["类型"] = "指"
    # global_index = ak.index_global_name_table()
    # global_index.rename(columns={"代码":"证券代码","指数名称":"证券简称"},inplace=True)
    # global_index["类型"] = "指"

    df = pd.DataFrame()
    df = pd.concat(objs=[df, stock_sha], ignore_index=True)
    df = pd.concat(objs=[df, stock_shb], ignore_index=True)
    df = pd.concat(objs=[df, stock_kcb], ignore_index=True)
    df = pd.concat(objs=[df, stock_sza], ignore_index=True)
    df = pd.concat(objs=[df, stock_szb], ignore_index=True)
    df = pd.concat(objs=[df, stock_bj], ignore_index=True)
    df = pd.concat(objs=[df, stock_etf], ignore_index=True)
    df = pd.concat(objs=[df, stock_lof], ignore_index=True)
    df = pd.concat(objs=[df, stock_closefund], ignore_index=True)
    df = pd.concat(objs=[df, index_stock], ignore_index=True)
    df.columns = ["code", "name", "type", "market"]

    return df


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