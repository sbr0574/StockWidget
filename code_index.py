import json
import os
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


def _to_prefixed_code(code: str) -> str:
    code = str(code or "").strip().lower()
    if len(code) == 8 and code[:2] in ("sh", "sz", "bj"):
        return code
    if len(code) == 6 and code.isdigit():
        if code[0] in ("6", "5", "9"):
            return f"sh{code}"
        if code[0] in ("0", "1", "2", "3"):
            return f"sz{code}"
        if code[0] in ("4", "8"):
            return f"bj{code}"
    return code


def _name_pinyin(name: str) -> tuple[str, str]:
    text = str(name or "").strip()
    if not text:
        return "", ""
    if pinyin is None:
        return "", ""
    py_full = "".join(x[0] for x in pinyin(text, style=Style.NORMAL, strict=False))
    py_abbr = "".join(x[0] for x in pinyin(text, style=Style.FIRST_LETTER, strict=False))
    return py_full.lower(), py_abbr.lower()


def _cache_file(app_name: str) -> str:
    cache_dir, _ = config_paths(app_name)
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, "stock_codes_cache.json")


def load_cached_index(app_name: str) -> list[dict]:
    path = _cache_file(app_name)
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_cached_index(app_name: str, entries: list[dict]):
    path = _cache_file(app_name)
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as file:
        json.dump(entries, file, ensure_ascii=False)
    os.replace(tmp, path)


def refresh_index_from_akshare(app_name: str) -> list[dict]:
    if ak is None:
        raise RuntimeError("akshare unavailable")

    df = ak.stock_info_a_code_name()
    entries = []
    for _, row in df.iterrows():
        code_raw = str(row.get("code", "")).strip()
        name = str(row.get("name", "")).strip()
        code = _to_prefixed_code(code_raw)
        py_full, py_abbr = _name_pinyin(name)
        entries.append({
            "code": code,
            "code_num": code_raw,
            "name": name,
            "py": py_full,
            "abbr": py_abbr,
        })

    entries.sort(key=lambda item: item["code"])
    _save_cached_index(app_name, entries)
    return entries


def refresh_once_per_day(app_name: str, cfg: dict) -> tuple[list[dict], dict]:
    today = datetime.now().strftime("%Y-%m-%d")
    last_day = str(cfg.get("code_index_updated", ""))
    entries = load_cached_index(app_name)
    if last_day == today and entries:
        return entries, cfg

    try:
        entries = refresh_index_from_akshare(app_name)
        cfg["code_index_updated"] = today
    except Exception:
        if not entries:
            entries = []
    return entries, cfg


def find_suggestions(entries: list[dict], text: str, limit: int = 20) -> list[dict]:
    q = str(text or "").strip().lower()
    if not q:
        return entries[:limit]

    scored = []
    for item in entries:
        code = item.get("code", "")
        code_num = item.get("code_num", "")
        name = str(item.get("name", ""))
        py = str(item.get("py", ""))
        abbr = str(item.get("abbr", ""))
        score = 0
        if code.startswith(q) or code_num.startswith(q):
            score = 100
        elif q in code or q in code_num:
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
