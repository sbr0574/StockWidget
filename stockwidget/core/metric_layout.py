"""显示指标的稳定标识、顺序规范化和旧配置兼容。"""

from dataclasses import dataclass
from typing import Mapping, Sequence


@dataclass(frozen=True)
class MetricSpec:
    metric_id: str
    label: str
    headers: tuple[str, ...]
    legacy_attr: str
    default_visible: bool = False


METRIC_SPECS = (
    MetricSpec("price", "现价", ("现价",), "price_visible", True),
    MetricSpec("change", "涨跌", ("涨跌",), "change_visible"),
    MetricSpec("change_pct", "涨幅", ("涨幅",), "change_pct_visible", True),
    MetricSpec("profit", "浮盈", ("浮盈",), "profit_visible"),
    MetricSpec("b1s1", "买一/卖一", ("买一", "卖一"), "b1s1_visible"),
    MetricSpec("commi", "委比", ("委比",), "commi_visible"),
    MetricSpec("volume", "成交量", ("成交量",), "vol_visible"),
    MetricSpec("amount", "成交额", ("成交额",), "amount_visible"),
    MetricSpec("average", "均价", ("均价",), "avg_visible"),
    MetricSpec("kline", "日K线", ("K线",), "kline_visible"),
)

METRIC_BY_ID = {spec.metric_id: spec for spec in METRIC_SPECS}
METRIC_IDS = tuple(METRIC_BY_ID)
HEADER_TO_METRIC_ID = {
    header: spec.metric_id
    for spec in METRIC_SPECS
    for header in spec.headers
}
DEFAULT_VISIBLE_METRICS = tuple(
    spec.metric_id for spec in METRIC_SPECS if spec.default_visible
)


def normalize_visible_metrics(value: Sequence[str] | None) -> list[str]:
    """过滤未知项和重复项，同时保留合法指标的输入顺序。"""
    if not isinstance(value, (list, tuple)):
        return []

    normalized = []
    seen = set()
    for raw_metric_id in value:
        metric_id = str(raw_metric_id or "").strip()
        if metric_id in METRIC_BY_ID and metric_id not in seen:
            normalized.append(metric_id)
            seen.add(metric_id)
    return normalized


def visible_metrics_from_config(cfg: Mapping | None) -> list[str]:
    """读取新有序列表；字段缺失或类型错误时从旧可见布尔值迁移。"""
    cfg = cfg if isinstance(cfg, Mapping) else {}
    raw_metrics = cfg.get("visible_metrics")
    if isinstance(raw_metrics, (list, tuple)):
        return normalize_visible_metrics(raw_metrics)

    return [
        spec.metric_id
        for spec in METRIC_SPECS
        if bool(cfg.get(spec.legacy_attr, spec.default_visible))
    ]


def expand_metric_headers(metric_ids: Sequence[str] | None) -> list[str]:
    """把逻辑指标顺序展开成浮窗的物理表头顺序。"""
    headers = []
    for metric_id in normalize_visible_metrics(metric_ids):
        headers.extend(METRIC_BY_ID[metric_id].headers)
    return headers


def metric_id_for_header(header: str) -> str | None:
    return HEADER_TO_METRIC_ID.get(str(header))


def legacy_visibility(metric_ids: Sequence[str] | None) -> dict[str, bool]:
    """生成旧布尔字段，供过渡期配置和现有调用方继续使用。"""
    visible = set(normalize_visible_metrics(metric_ids))
    return {
        spec.legacy_attr: spec.metric_id in visible
        for spec in METRIC_SPECS
    }
