import requests
from typing import Tuple

def _format_volume(value: float) -> str:
    if value < 1e4:
        return f"{value}"
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    return f"{value / 1e8:.2f}亿"


def _format_amount(value: float) -> str:
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    if value < 1e12:
        return f"{value / 1e8:.2f}亿"
    return f"{value / 1e12:.2f}万亿"

def almost_eq(a, b, dec):
    try:
        return round(float(a), dec) == round(float(b), dec)
    except Exception:
        return False


def request_sina(req_codes: list[str]) -> Tuple[list, dict]:
    data = {}
    if len(req_codes) == 0:
        return [], {}
    label = ",".join([str(c).strip() for c in req_codes if str(c).strip()])
    url = "https://hq.sinajs.cn/list=" + label
    headers = {"Referer": "https://finance.sina.com.cn", "User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, timeout=3)
    response.encoding = "gbk"
    for index, line in enumerate(response.text.split("\n")):
        if not line or '"' not in line:
            continue
        heads = line.split('="')[0].split("_")
        parts = line.split('="')[1].split(",")
        if len(parts) < 3:
            continue
        code = heads[2]
        data[code] = parts
    return [(code in data) for code in req_codes], data


def fetch_stock_rows(data: dict, name_length: int, b1s1_display: str):
    price_data = []
    sign_data = []
    for code, parts in data.items():
        name = parts[0]
        opening_price = float(parts[1] or 0)
        prev_close = float(parts[2] or 0)
        current_price = float(parts[3] or 0)
        high_price = float(parts[4] or 0)
        low_price = float(parts[5] or 0)
        first_pur = float(parts[6] or 0)
        first_sell = float(parts[7] or 0)
        deals_vol = float(parts[8] or 0)
        deals_amt = float(parts[9] or 0)
        purchaser = [int(x or 0) for x in parts[10:19:2]]
        seller = [int(x or 0) for x in parts[20:29:2]]
        etf = code[2] in ("1", "5")

        b1_label = ""
        s1_label = ""
        b1_color_sign = 0
        s1_color_sign = 0

        dec = 3 if etf else 2

        buy_marker = "<" if first_pur > 0 and almost_eq(current_price, first_pur, dec) else " "
        sell_marker = ">" if first_sell > 0 and almost_eq(current_price, first_sell, dec) else " "

        if first_pur == first_sell > 0:
            current_price = first_sell
            paired = seller[0]
            unpaired_sign = -seller[1] if seller[1] > 0 else purchaser[1]
            paired_cnt = int(paired / 100)
            unpaired_cnt = int(unpaired_sign / 100)
            b_price = f"{first_pur:.3f}" if etf else f"{first_pur:.2f}"
            s_price = f"{first_sell:.3f}" if etf else f"{first_sell:.2f}"
            if b1s1_display == "price":
                b1_label = b_price
                s1_label = s_price
            elif b1s1_display == "both":
                b1_label = f"{paired_cnt:d}({b_price})"
                s1_label = f"{unpaired_cnt:+d}({s_price})"
            else:
                b1_label = f"{paired_cnt:d}"
                s1_label = f"{unpaired_cnt:+d}"
            b1_color_sign = (unpaired_sign > 0) - (unpaired_sign < 0)
            s1_color_sign = b1_color_sign
        else:
            if first_pur > 0:
                cnt = f"{int(purchaser[0] / 100)}"
                b_price = f"{first_pur:.3f}" if etf else f"{first_pur:.2f}"
                if b1s1_display == "price":
                    b1_label = f"{b_price}{buy_marker}"
                elif b1s1_display == "both":
                    b1_label = f"{cnt}({b_price}){buy_marker}"
                else:
                    b1_label = f"{cnt}{buy_marker}"
            else:
                b1_label = f"-{buy_marker}"

            if first_sell > 0:
                cnt = f"{int(seller[0] / 100)}"
                s_price = f"{first_sell:.3f}" if etf else f"{first_sell:.2f}"
                if b1s1_display == "price":
                    s1_label = f"{sell_marker}{s_price}"
                elif b1s1_display == "both":
                    s1_label = f"{sell_marker}{cnt}({s_price})"
                else:
                    s1_label = f"{sell_marker}{cnt}"
            else:
                s1_label = f"{sell_marker}-"
            b1_color_sign = 1
            s1_color_sign = -1

        if current_price == 0:
            current_price = prev_close
        if opening_price == 0:
            opening_price = current_price
            high_price = current_price
            low_price = current_price

        change = current_price - prev_close if prev_close else 0.0
        change_pct = (current_price / prev_close - 1) * 100 if prev_close else 0.0
        avg = (deals_amt / deals_vol) if deals_vol > 0 else prev_close
        p_sum, s_sum = sum(purchaser), sum(seller)
        committee = (100 * (p_sum - s_sum) / (p_sum + s_sum)) if (p_sum + s_sum) > 0 else 0.0

        arrow = " "
        if high_price > low_price:
            if current_price == high_price:
                arrow = "↑"
            elif current_price == low_price:
                arrow = "↓"

        k_payload = {"k": (opening_price, current_price, high_price, low_price, prev_close)}
        precision = 3 if etf else 2
        value_prefix = code[2:] if code[:2] in ("sh","sz","bj") else code
        name_prefix = name if name_length == 0 else name[:name_length]
        price_data.append([
            value_prefix,
            name_prefix,
            f"{current_price:.{precision}f}{arrow}",
            f"{change:+.{precision}f}",
            f"{change_pct:+.2f}%",
            b1_label,
            s1_label,
            f"{committee:+.2f}%",
            _format_volume(deals_vol),
            _format_amount(deals_amt),
            f"{avg:.{precision}f}",
            k_payload,
        ])

        sign_data.append({
            "delta": (change > 0) - (change < 0),
            "commi": (committee > 0) - (committee < 0),
            "avg": (avg > prev_close) - (avg < prev_close),
            "b1": b1_color_sign,
            "s1": s1_color_sign,
        })

    return price_data, sign_data
