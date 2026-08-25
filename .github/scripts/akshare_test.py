import requests
import pandas as pd

def stock_info_sh_name_code() -> pd.DataFrame:
    """
    上海证券交易所-股票列表
    https://www.sse.com.cn/assortment/stock/list/share/
    :param symbol: choice of {"主板A股": "1", "主板B股": "2", "科创板": "8"}
    :type symbol: str
    :return: 指定 indicator 的数据
    :rtype: pandas.DataFrame
    """
    # indicator_map = {"主板A股": "1", "主板B股": "2", "科创板": "8"}
    url = "https://query.sse.com.cn/sseQuery/commonQuery.do"
    headers = {
        "Host": "query.sse.com.cn",
        "Pragma": "no-cache",
        "Referer": "https://www.sse.com.cn/assortment/stock/list/share/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/81.0.4044.138 Safari/537.36",
    }
    params = {
        "STOCK_TYPE": "1",
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
    }
    r = requests.get(url, params=params, headers=headers)
    print(r.raw)
    data_json = r.json()
    temp_df = pd.DataFrame(data_json["result"])
    col_stock_code = "A_STOCK_CODE"
    temp_df.rename(
        columns={
            col_stock_code: "证券代码",
            "SEC_NAME_CN": "证券简称",
            "SEC_NAME_FULL": "证券全称",
            "COMPANY_ABBR": "公司简称",
            "FULL_NAME": "公司全称",
            "LIST_DATE": "上市日期",
        },
        inplace=True,
    )
    temp_df = temp_df[
        [
            "证券代码",
            "证券简称",
            "证券全称",
            "公司简称",
            "公司全称",
            "上市日期",
        ]
    ]
    temp_df["上市日期"] = pd.to_datetime(temp_df["上市日期"], errors="coerce").dt.date
    return temp_df

sh = stock_info_sh_name_code("主板A股")
print(sh.head(5))
