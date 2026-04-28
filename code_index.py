import json
import os
import re
import threading
from datetime import datetime

from config_store import config_paths

try:
    import akshare as ak
except Exception:
    ak = None

try:
    from pypinyin import Style, pinyin
except Exception:
    pinyin = None
    Style = None


def _to_prefixed_code(code: str, board: str = "") -> str:
    code = str(code or "").strip().lower()
    board = str(board or "").strip().lower()
    if len(code) == 8 and code[:2] in ("sh", "sz", "bj"):
        return code
    if len(code) == 6 and code.isdigit():
        if board == "沪a" or code[0] in ("6", "5", "9"):
            return f"sh{code}"
        if board == "深a" or code[0] in ("0", "1", "2", "3"):
            return f"sz{code}"
        if board == "京a" or code[0] in ("4", "8"):
            return f"bj{code}"
    return code


def _norm_hk_code(code: str) -> str:
    raw = str(code or "").strip().lower()
    if raw.startswith("hk"):
        raw = raw[2:]
    raw = "".join(ch for ch in raw if ch.isdigit())
    if not raw:
        return ""
    return f"hk{raw.zfill(5)}"


def _norm_us_code(code: str) -> str:
    raw = str(code or "").strip().lower()
    if raw.startswith("us"):
        raw = raw[2:]
    raw = raw.lstrip(".")
    raw = raw.replace(" ", "")
    return f"us{raw}" if raw else ""


def _ascii_words(text: str) -> list[str]:
    return [w.lower() for w in re.findall(r"[A-Za-z0-9]+", str(text or "")) if w]


def _name_pinyin(name: str) -> tuple[str, str]:
    text = str(name or "").strip()
    if not text:
        return "", ""

    py_full = ""
    py_abbr = ""
    if pinyin is not None:
        try:
            py_full = "".join(x[0] for x in pinyin(text, style=Style.NORMAL, strict=False)).lower()
            py_abbr = "".join(x[0] for x in pinyin(text, style=Style.FIRST_LETTER, strict=False)).lower()
        except Exception:
            py_full = ""
            py_abbr = ""

    words = _ascii_words(text)
    ascii_full = "".join(words)
    ascii_abbr = "".join(w[0] for w in words if w)

    if ascii_full and ascii_full not in py_full:
        py_full = f"{py_full}{ascii_full}"
    if ascii_abbr and ascii_abbr not in py_abbr:
        py_abbr = f"{py_abbr}{ascii_abbr}"
    return py_full, py_abbr


def _cache_file(app_name: str) -> str:
    cache_dir, _ = config_paths(app_name)
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, "stock_codes_cache.json")


def load_cached_index_meta(app_name: str) -> dict:
    path = _cache_file(app_name)
    try:
        with open(path, "r", encoding="utf-8") as file:
            payload = json.load(file)
            if isinstance(payload, list):
                return {"updated_on": "", "updated_at": "", "entries": payload}
            if isinstance(payload, dict):
                entries = payload.get("entries", [])
                if isinstance(entries, list):
                    return {
                        "updated_on": str(payload.get("updated_on", "")),
                        "updated_at": str(payload.get("updated_at", "")),
                        "entries": entries,
                    }
    except Exception:
        pass
    return {"updated_on": "", "updated_at": "", "entries": []}


def load_cached_index(app_name: str) -> list[dict]:
    return load_cached_index_meta(app_name).get("entries", [])


def _save_cached_index(app_name: str, entries: list[dict]):
    path = _cache_file(app_name)
    now = datetime.now()
    payload = {
        "updated_on": now.strftime("%Y-%m-%d"),
        "updated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "entries": entries,
    }
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False)
    os.replace(tmp, path)


def _row_get(row, *keys, default=""):
    for key in keys:
        if key in row and str(row[key]).strip() not in ("", "nan", "None"):
            return str(row[key]).strip()
    return default


def refresh_index_from_akshare(app_name: str) -> list[dict]:
    if ak is None:
        raise RuntimeError("akshare unavailable")

    entries: list[dict] = []

    def add_entry(code: str, code_num: str, name: str, category: str, market: str):
        key = str(code or "").strip().lower()
        if not key:
            return
        py_full, py_abbr = _name_pinyin(name)
        symbol_words = _ascii_words(code_num)
        if symbol_words:
            merged = "".join(symbol_words)
            if merged not in py_full:
                py_full = f"{py_full}{merged}"
            if merged and merged[0] not in py_abbr:
                py_abbr = f"{py_abbr}{''.join(w[0] for w in symbol_words if w)}"
        entries.append({
            "code": key,
            "code_num": str(code_num or "").strip(),
            "name": str(name or "").strip(),
            "py": py_full,
            "abbr": py_abbr,
            "category": category,
            "market": market,
        })

    # 沪深京 A 股 + 场内基金
    df_a = ak.stock_info_a_code_name()
    for _, row in df_a.iterrows():
        code_raw = _row_get(row, "code", "代码")
        name = _row_get(row, "name", "名称")
        code = _to_prefixed_code(code_raw)
        if not code:
            continue
        category = "沪A" if code.startswith("sh") else ("深A" if code.startswith("sz") else "京A")
        add_entry(code, code_raw, name, category, "cn")

    df_etf = ak.fund_etf_spot_em()
    for _, row in df_etf.iterrows():
        code_raw = _row_get(row, "代码", "code")
        name = _row_get(row, "名称", "name")
        code = _to_prefixed_code(code_raw)
        if not code:
            continue
        category = "沪基" if code.startswith("sh") else "深基"
        add_entry(code, code_raw, name, category, "fund")

    # 港股
    df_hk = ak.stock_hk_spot_em()
    for _, row in df_hk.iterrows():
        code_raw = _row_get(row, "代码", "symbol")
        name = _row_get(row, "名称", "name")
        code = _norm_hk_code(code_raw)
        if code:
            add_entry(code, code_raw, name, "港股", "hk")

    # 美股
    df_us = ak.stock_us_spot_em()
    for _, row in df_us.iterrows():
        code_raw = _row_get(row, "代码", "symbol")
        name = _row_get(row, "名称", "name")
        code = _norm_us_code(code_raw)
        if code:
            add_entry(code, code_raw, name, "美股", "us")

    merged = {}
    for item in entries:
        merged[item["code"]] = item

    out = sorted(merged.values(), key=lambda item: item["code"])
    _save_cached_index(app_name, out)
    return out


def is_index_updated_today(app_name: str) -> bool:
    today = datetime.now().strftime("%Y-%m-%d")
    return load_cached_index_meta(app_name).get("updated_on") == today


def refresh_once_per_day(app_name: str) -> tuple[list[dict], bool]:
    entries = load_cached_index(app_name)
    if entries and is_index_updated_today(app_name):
        return entries, False
    try:
        entries = refresh_index_from_akshare(app_name)
        return entries, True
    except Exception:
        return entries, False


def refresh_in_background_if_needed(app_name: str, on_ready=None):
    entries = load_cached_index(app_name)
    if callable(on_ready):
        try:
            on_ready(entries)
        except Exception:
            pass

    if is_index_updated_today(app_name):
        return

    def _job():
        latest = entries
        try:
            latest = refresh_index_from_akshare(app_name)
        except Exception:
            pass
        if callable(on_ready):
            try:
                on_ready(latest)
            except Exception:
                pass

    threading.Thread(target=_job, daemon=True).start()


def find_suggestions(entries: list[dict], text: str, limit: int = 20) -> list[dict]:
    q = str(text or "").strip().lower()
    if not q:
        return entries[:limit]

    scored = []
    for item in entries:
        code = str(item.get("code", "")).lower()
        code_num = str(item.get("code_num", "")).lower()
        name = str(item.get("name", ""))
        name_l = name.lower()
        py = str(item.get("py", "")).lower()
        abbr = str(item.get("abbr", "")).lower()
        cat = str(item.get("category", "")).lower()
        score = 0
        if code.startswith(q) or code_num.startswith(q):
            score = 100
        elif q in code or q in code_num:
            score = 85
        elif name_l.startswith(q):
            score = 70
        elif q in name_l:
            score = 60
        elif py.startswith(q):
            score = 50
        elif q in py:
            score = 40
        elif abbr.startswith(q):
            score = 30
        elif q in abbr:
            score = 20
        elif q in cat:
            score = 10

        if score > 0:
            scored.append((score, item))

    scored.sort(key=lambda pair: (-pair[0], pair[1].get("code", "")))
    return [item for _, item in scored[:limit]]
