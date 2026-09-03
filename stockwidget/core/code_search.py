# -*- coding: utf-8 -*-
"""自选股代码搜索与规范化（纯业务逻辑，无 Qt/网络依赖）。

根据数字代码、拼音、首字母、中文名/英文名在代码列表中匹配建议。
"""

from dataclasses import dataclass
import re
import unicodedata


_DIRECT_REGIONS = frozenset(("sh", "sz", "bj", "hk", "us"))
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
    """将现有 type 标签归入四个互斥类别。"""
    return _SPECIAL_CATEGORIES.get(str(type_ or "").strip(), "stock")


def _entry_region(market: str) -> str:
    """将市场标识归入添加面板的六个地区。"""
    value = _normalize_text(market)
    return value if value in _DIRECT_REGIONS else "other"


@dataclass(frozen=True, slots=True)
class SearchRecord:
    """一条预规范化的搜索记录，避免每次按键重复整理代码表。"""

    entry: dict
    category: str
    region: str
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
            region=_entry_region(info["market"]),
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


@dataclass(frozen=True, slots=True)
class SearchPage:
    """添加面板的一页搜索结果。"""

    items: tuple[dict, ...]
    total: int
    page: int
    page_size: int
    page_count: int


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


def _record_score(record: SearchRecord, tokens) -> int | None:
    total = 0
    for variants in tokens:
        token_score = 0
        for query in variants:
            token_score = max(token_score, _match_score(record.fields, query))
        if not token_score:
            return None
        total += token_score
    return total


def query_search_index(index, text: str = "", *, categories=None, regions=None,
                       page: int = 1, page_size: int = 10) -> SearchPage:
    """按类别和地区查询索引，允许空关键词并返回分页信息。

    ``None`` 表示不过滤该维度，空集合表示没有结果。同一维度内取并集，
    类别与地区之间取交集。
    """
    try:
        requested_page = max(1, int(page))
    except (TypeError, ValueError):
        requested_page = 1
    try:
        normalized_page_size = max(1, int(page_size))
    except (TypeError, ValueError):
        normalized_page_size = 10

    category_filter = None if categories is None else frozenset(categories)
    region_filter = None if regions is None else frozenset(regions)
    tokens = _query_tokens(text)
    matched = []

    for position, record in enumerate(index or ()):
        if category_filter is not None and record.category not in category_filter:
            continue
        if region_filter is not None and record.region not in region_filter:
            continue
        score = _record_score(record, tokens) if tokens else 0
        if score is not None:
            matched.append((score, position, record.entry))

    if tokens:
        matched.sort(key=lambda item: (-item[0], item[2]["code"], item[2]["key"]))
    # 空查询保留代码表顺序，跨页稳定且无需额外排序。

    total = len(matched)
    page_count = (total + normalized_page_size - 1) // normalized_page_size
    current_page = min(requested_page, page_count) if page_count else 0
    start = (current_page - 1) * normalized_page_size if current_page else 0
    items = tuple(item[2] for item in matched[start:start + normalized_page_size])
    return SearchPage(
        items=items,
        total=total,
        page=current_page,
        page_size=normalized_page_size,
        page_count=page_count,
    )


def search_suggestions(index, text: str, limit: int = 20) -> list[dict]:
    """在预构建索引中搜索；多个关键词全部命中才返回结果。"""
    tokens = _query_tokens(text)
    if not tokens or limit <= 0:
        return []

    result = query_search_index(
        index,
        text,
        page=1,
        page_size=limit,
    )
    return list(result.items)
