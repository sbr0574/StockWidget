import unittest

from stockwidget.core.metric_layout import (
    DEFAULT_VISIBLE_METRICS,
    METRIC_IDS,
    expand_metric_headers,
    legacy_visibility,
    normalize_visible_metrics,
    visible_metrics_from_config,
)


class MetricLayoutTests(unittest.TestCase):
    def test_defaults_match_existing_visibility_and_exclude_name(self):
        self.assertEqual(DEFAULT_VISIBLE_METRICS, ("price", "change_pct"))
        self.assertEqual(
            visible_metrics_from_config({}),
            ["price", "change_pct"],
        )
        self.assertNotIn("name", METRIC_IDS)

    def test_legacy_flags_migrate_in_canonical_order(self):
        config = {
            "price_visible": False,
            "change_visible": True,
            "change_pct_visible": False,
            "b1s1_visible": True,
            "kline_visible": True,
        }
        self.assertEqual(
            visible_metrics_from_config(config),
            ["change", "b1s1", "kline"],
        )

    def test_explicit_order_preserves_empty_and_filters_invalid_values(self):
        self.assertEqual(visible_metrics_from_config({"visible_metrics": []}), [])
        self.assertEqual(
            visible_metrics_from_config(
                {"visible_metrics": ["kline", "price", "kline", "bad"]}
            ),
            ["kline", "price"],
        )

    def test_normalize_and_expand_combined_level_one_metric(self):
        normalized = normalize_visible_metrics(
            ["amount", "b1s1", "change_pct"]
        )
        self.assertEqual(normalized, ["amount", "b1s1", "change_pct"])
        self.assertEqual(
            expand_metric_headers(normalized),
            ["成交额", "买一", "卖一", "涨幅"],
        )

    def test_legacy_visibility_is_kept_in_sync(self):
        flags = legacy_visibility(["volume", "price"])
        self.assertTrue(flags["vol_visible"])
        self.assertTrue(flags["price_visible"])
        self.assertFalse(flags["amount_visible"])


if __name__ == "__main__":
    unittest.main()
