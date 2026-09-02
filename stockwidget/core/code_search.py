# -*- coding: utf-8 -*-
"""自选股代码搜索与规范化（纯业务逻辑，无 Qt/网络依赖）。

根据数字代码、拼音、首字母、中文名/英文名在代码列表中匹配建议。
"""

from dataclasses import dataclass
import re
import unicodedata


SEARCH_CATEGORIES = ("stock", "fund", "index", "futures")
_SPECIAL_CATEGORIES = {
    "基": "fund",
    "指": "index",
    "期": "futures",
}
_WHITESPACE_RE = re.compile(r"\s+")


def _normalize_text(value) -> str:
    """统一全半角、大小写和首尾空白，供搜索字段比较。"""
    return unicodedata.normalize("NFKC", str(value or "")).strip().casefold()


def _compact_text(value) -> str:
    """返回移除全部空白的搜索文本，兼容代码表名称中的排版空格。"""
    return _WHITESPACE_RE.sub("", _normalize_text(value))


def normalize_stock_entry(item: dict) -> dict:
    """把代码列表中的原始条目规范化为统一的搜索字段 dict。"""
    market = _normalize_text(item.get("market", ""))
    code = unicodedata.normalize("NFKC", str(item.get("code", "") or "")).strip()
    if market in {"sh", "sz", "bj"} and code.isdigit():
        code = code.zfill(6)
    if market == "hk" and code.isdigit():
        code = code.zfill(5)
    key = _normalize_text(item.get("key", ""))
    if not key and market and code:
        key = f"{market}{code}"
    name_en = _normalize_text(item.get("name_en", "") or item.get("engname", ""))
    return {
        "key": key,
        "market": market,
        "code": code,
        "name": str(item.get("name", "") or "").strip(),
        "type": str(item.get("type", "") or "").strip(),
        "py": _normalize_text(item.get("py", "")),
        "abbr": _normalize_text(item.get("abbr", "")),
        "name_en": name_en,
    }


def _entry_category(type_: str) -> str:
    """将现有 type 标签归入搜索下拉的四个互斥类别。"""
    return _SPECIAL_CATEGORIES.get(str(type_ or "").strip(), "stock")


@dataclass(frozen=True, slots=True)
class SearchRecord:
    """一条预规范化的搜索记录，避免每次按键重复整理代码表。"""

    entry: dict
    category: str
    # key, code, name, compact name, pinyin, abbreviation,
    # English name, compact English name
    fields: tuple[str, str, str, str, str, str, str, str]


def build_search_index(codes: dict) -> tuple[SearchRecord, ...]:
    """将代码表预处理成可重复查询的只读搜索索引。"""
    if not isinstance(codes, dict):
        return ()

    records = []
    for raw_key, raw_info in codes.items():
        raw_entry = raw_info if isinstance(raw_info, dict) else {}
        info = normalize_stock_entry({"key": raw_key, **raw_entry})
        name = _normalize_text(info["name"])
        name_en = _normalize_text(info["name_en"])
        records.append(SearchRecord(
            entry=info,
            category=_entry_category(info["type"]),
            fields=(
                _normalize_text(info["key"]),
                _normalize_text(info["code"]),
                name,
                _WHITESPACE_RE.sub("", name),
                _normalize_text(info["py"]),
                _normalize_text(info["abbr"]),
                name_en,
                _WHITESPACE_RE.sub("", name_en),
            ),
        ))
    return tuple(records)


def _query_variants(text: str) -> tuple[str, ...]:
    """把一个关键词扩展成候选查询（包括数字代码补零）。"""
    query = _compact_text(text)
    variants = [query]
    if query.isdigit():
        if len(query) <= 5:
            variants.append(query.zfill(5))
        if len(query) <= 6:
            variants.append(query.zfill(6))
    return tuple(dict.fromkeys(variant for variant in variants if variant))


def _query_tokens(text: str) -> tuple[tuple[str, ...], ...]:
    normalized = _normalize_text(text)
    if not normalized:
        return ()
    return tuple(_query_variants(token) for token in _WHITESPACE_RE.split(normalized))


def _match_score(fields: tuple[str, ...], query: str) -> int:
    """计算单个关键词对一条记录的最佳相关度。"""
    key, code, name, compact_name, py, abbr, name_en, compact_name_en = fields
    if query in (key, code, name, compact_name, py, abbr, name_en, compact_name_en):
        return 120
    if key.startswith(query):
        return 110
    if code.startswith(query):
        return 105
    if query in key or query in code:
        return 95
    if (name.startswith(query) or compact_name.startswith(query)
            or name_en.startswith(query) or compact_name_en.startswith(query)):
        return 90
    if (query in name or query in compact_name
            or query in name_en or query in compact_name_en):
        return 80
    if abbr.startswith(query):
        return 75
    if query in abbr:
        return 65
    if py.startswith(query):
        return 60
    if query in py:
        return 50
    return 0


def search_suggestions(index, text: str, limit: int = 20,
                       *, category: str | None = None) -> list[dict]:
    """在预构建索引中搜索；多个关键词全部命中才返回结果。"""
    tokens = _query_tokens(text)
    if not tokens or limit <= 0:
        return []
    if category is not None and category not in SEARCH_CATEGORIES:
        return []

    scored = []
    for record in index or ():
        if category is not None and record.category != category:
            continue

        total = 0
        for variants in tokens:
            token_score = 0
            for query in variants:
                token_score = max(token_score, _match_score(record.fields, query))
            if not token_score:
                break
            total += token_score
        else:
            scored.append((total, record.entry))

    scored.sort(key=lambda pair: (-pair[0], pair[1]["code"], pair[1]["key"]))
    return [item for _, item in scored[:limit]]


def find_suggestions(codes: dict, text: str, limit: int = 20,
                     *, category: str | None = None) -> list[dict]:
    """在代码列表中按相关度返回匹配建议（精确 > 前缀 > 包含）。"""
    return search_suggestions(
        build_search_index(codes),
        text,
        limit,
        category=category,
    )
