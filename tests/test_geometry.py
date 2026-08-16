# -*- coding: utf-8 -*-
"""多显示器窗口位置恢复逻辑的单元测试。"""

import unittest

from stockwidget.core.geometry import clamp_point, resolve_restore_position, screen_containing

# 主屏（右），副屏（左），副屏坐标为负
PRIMARY = (0, 0, 1920, 1080)
SECONDARY = (-1920, 0, 1920, 1080)
RECTS = [PRIMARY, SECONDARY]
W, H = 200, 100  # 窗口尺寸


class TestScreenContaining(unittest.TestCase):
    def test_primary(self):
        self.assertEqual(screen_containing(100, 100, RECTS), PRIMARY)

    def test_secondary_negative_coords(self):
        self.assertEqual(screen_containing(-100, 100, RECTS), SECONDARY)

    def test_none_when_outside(self):
        self.assertIsNone(screen_containing(99999, 99999, RECTS))


class TestClampPoint(unittest.TestCase):
    def test_inside_unchanged(self):
        self.assertEqual(clamp_point(100, 100, PRIMARY, W, H), (100, 100))

    def test_outside_right_bottom(self):
        self.assertEqual(clamp_point(5000, 5000, PRIMARY, W, H), (1720, 980))

    def test_outside_left_top(self):
        self.assertEqual(clamp_point(-5000, -5000, PRIMARY, W, H), (0, 0))


class TestResolveRestorePosition(unittest.TestCase):
    def test_restore_on_secondary(self):
        # 保存在副屏 -> 原位恢复
        self.assertEqual(
            resolve_restore_position((-1500, 200), RECTS, PRIMARY, W, H),
            (-1500, 200),
        )

    def test_restore_on_primary(self):
        self.assertEqual(
            resolve_restore_position((500, 300), RECTS, PRIMARY, W, H),
            (500, 300),
        )

    def test_secondary_disconnected_falls_back_to_primary(self):
        # 副屏断开：保存位置不再落在任何屏 -> 夹取回主屏内（可见）
        self.assertEqual(
            resolve_restore_position((-1500, 200), [PRIMARY], PRIMARY, W, H),
            (0, 200),
        )

    def test_clamp_into_secondary_when_near_edge(self):
        # 保存在副屏内但靠近右边缘，窗口超出副屏边界 -> 夹取回副屏内
        self.assertEqual(
            resolve_restore_position((-150, 200), RECTS, PRIMARY, W, H),
            (-200, 200),
        )

    def test_none_saved_uses_primary_default(self):
        self.assertEqual(
            resolve_restore_position(None, RECTS, PRIMARY, W, H),
            (1920 - W - 40, 1080 - H - 80),
        )


if __name__ == "__main__":
    unittest.main()
