import json
import os
import re

APP_NAME = "StockWidget"
MARKET_PREFIXES = {"sh", "sz", "bj"}


def config_paths() -> str:
    return os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), APP_NAME)


def load_file(file_name: str, except_ret: dict | None = None) -> dict:
    fallback = {} if except_ret is None else except_ret
    path = os.path.join(config_paths(), file_name)
    if not os.path.exists(path):
        return fallback
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return fallback


def save_file(data: dict, file_name: str) -> None:
    os.makedirs(config_paths(), exist_ok=True)
    config_file = os.path.join(config_paths(), file_name)
    tmp_file = config_file + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    os.replace(tmp_file, config_file)


def code_without_market(code: str) -> str:
    value = str(code or "").strip().lower()
    return value[2:] if len(value) == 8 and value[:2] in MARKET_PREFIXES else value


def normalize_stock_entry(item: dict) -> dict:
    market = str(item.get("market", "") or "").strip().lower()
    code = str(item.get("code", "") or "").strip().zfill(6)
    key = str(item.get("key", "") or "").strip().lower()
    if not key and market and code:
        key = f"{market}{code}"
    return {
        "key": key,
        "market": market,
        "code": code,
        "name": str(item.get("name", "") or "").strip(),
        "type": str(item.get("type", "") or "").strip(),
        "py": str(item.get("py", "") or "").strip().lower(),
        "abbr": str(item.get("abbr", "") or "").strip().lower(),
    }


def _query_variants(text: str) -> set[str]:
    q = str(text or "").strip().lower().replace(" ", "")
    variants = {q}
    if len(q) == 8 and q[:2] in MARKET_PREFIXES:
        variants.add(q[2:])
    digits = re.sub(r"\D", "", q)
    if digits:
        variants.add(digits.zfill(6) if len(digits) <= 6 else digits)
    return {v for v in variants if v}


def find_suggestions(codes: dict, text: str, limit: int = 20) -> list[dict]:
    queries = _query_variants(text)
    if not queries or not isinstance(codes, dict):
        return []

    scored = []
    for raw_key, raw_info in codes.items():
        info = normalize_stock_entry({"key": raw_key, **(raw_info or {})})
        fields = {
            "key": info["key"],
            "code": info["code"],
            "name": info["name"].lower(),
            "py": info["py"],
            "abbr": info["abbr"],
        }
        best = 0
        for q in queries:
            if fields["key"] == q or fields["code"] == q or fields["name"] == q or fields["py"] == q or fields["abbr"] == q:
                best = max(best, 120)
            elif fields["key"].startswith(q):
                best = max(best, 110)
            elif fields["code"].startswith(q):
                best = max(best, 105)
            elif q in fields["key"] or q in fields["code"]:
                best = max(best, 95)
            elif fields["name"].startswith(q):
                best = max(best, 90)
            elif q in fields["name"]:
                best = max(best, 80)
            elif fields["abbr"].startswith(q):
                best = max(best, 75)
            elif q in fields["abbr"]:
                best = max(best, 65)
            elif fields["py"].startswith(q):
                best = max(best, 60)
            elif q in fields["py"]:
                best = max(best, 50)
        if best:
            scored.append((best, info))

    scored.sort(key=lambda pair: (-pair[0], pair[1]["code"], pair[1]["key"]))
    return [item for _, item in scored[:limit]]
