import requests
from typing import Tuple

def request_sina(req_codes: list[str]) -> Tuple[list, dict]:
    data = {}
    if len(req_codes) == 0:
        return [], {}
    label = ",".join([str(c).strip() for c in req_codes if str(c).strip()])
    url = "https://hq.sinajs.cn/list=" + label
    headers = {"Referer": "https://finance.sina.com.cn", "User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, timeout=1)
    response.encoding = "gbk"
    for line in response.text.split("\n"):
        if not line or '"' not in line:
            continue
        heads = line.split('="')[0].split("_")
        parts = line.split('="')[1].split(",")
        if len(parts) < 3:
            continue
        code = heads[2]
        data[code] = {
            "name" : parts[0],
            "opening_price" : float(parts[1] or 0),
            "prev_close" : float(parts[2] or 0),
            "current_price" : float(parts[3] or 0),
            "high_price" : float(parts[4] or 0),
            "low_price" : float(parts[5] or 0),
            # "first_pur" : float(parts[6] or 0),
            # "first_sell" : float(parts[7] or 0),
            "deals_vol" : int(parts[8] or 0),
            "deals_amt" : float(parts[9] or 0),
            "purchaser_vol" : [int(x or 0) for x in parts[10:20:2]],
            "purchaser_price" : [float(x or 0) for x in parts[11:20:2]],
            "seller_vol" : [int(x or 0) for x in parts[20:30:2]],
            "seller_price" : [float(x or 0) for x in parts[21:30:2]],
            "date": parts[31],
            "time": parts[32],
        }
        
    return [code in data for code in req_codes], data

