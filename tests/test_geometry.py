# -*- coding: utf-8 -*-
"""多显示器窗口位置恢复测试。"""

import unittest

from stockwidget.core.geometry import best_screen, clamp_point, resolve_restore_position, screen_containing


PRIMARY = (0, 0, 1920, 1080)
SECONDARY = (-1920, 0, 1920, 1080)
RECTS = [PRIMARY, SECONDARY]
WIDTH = 200
HEIGHT = 100


class ScreenContainingTests(unittest.TestCase):
    def test_primary(self):
        self.assertEqual(screen_containing(100, 100, RECTS), PRIMARY)

    def test_secondary_negative_coords(self):
        self.assertEqual(screen_containing(-100, 100, RECTS), SECONDARY)

    def test_none_when_outside(self):
        self.assertIsNone(screen_containing(99999, 99999, RECTS))


class ClampPointTests(unittest.TestCase):
    def test_inside_unchanged(self):
        self.assertEqual(clamp_point(100, 100, PRIMARY, WIDTH, HEIGHT), (100, 100))

    def test_outside_right_bottom(self):
        self.assertEqual(clamp_point(5000, 5000, PRIMARY, WIDTH, HEIGHT), (1720, 980))

    def test_outside_left_top(self):
        self.assertEqual(clamp_point(-5000, -5000, PRIMARY, WIDTH, HEIGHT), (0, 0))


class BestScreenTests(unittest.TestCase):
    def test_largest_overlap_allows_crossing_monitor_seams(self):
        self.assertEqual(best_screen(-80, 200, 200, 100, RECTS), PRIMARY)
        self.assertEqual(best_screen(-180, 200, 200, 100, RECTS), SECONDARY)

    def test_outside_uses_nearest_monitor_including_negative_coordinates(self):
        self.assertEqual(best_screen(-5000, 200, 200, 100, RECTS), SECONDARY)
        self.assertEqual(best_screen(3000, 200, 200, 100, RECTS), PRIMARY)

    def test_staggered_monitors_and_empty_screen_list(self):
        upper = (0, -1080, 1920, 1080)
        self.assertEqual(best_screen(200, -100, 200, 150, [PRIMARY, upper]), upper)
        self.assertIsNone(best_screen(0, 0, 200, 100, []))


class ResolveRestorePositionTests(unittest.TestCase):
    def test_restore_on_secondary(self):
        self.assertEqual(
            resolve_restore_position((-1500, 200), RECTS, PRIMARY, WIDTH, HEIGHT),
            (-1500, 200),
        )

    def test_restore_on_primary(self):
        self.assertEqual(
            resolve_restore_position((500, 300), RECTS, PRIMARY, WIDTH, HEIGHT),
            (500, 300),
        )

    def test_secondary_disconnected_falls_back_to_primary(self):
        self.assertEqual(
            resolve_restore_position((-1500, 200), [PRIMARY], PRIMARY, WIDTH, HEIGHT),
            (0, 200),
        )

    def test_clamp_into_secondary_when_near_edge(self):
        self.assertEqual(
            resolve_restore_position((-150, 200), RECTS, PRIMARY, WIDTH, HEIGHT),
            (-200, 200),
        )

    def test_none_saved_uses_primary_default(self):
        self.assertEqual(
            resolve_restore_position(None, RECTS, PRIMARY, WIDTH, HEIGHT),
            (1680, 900),
        )


if __name__ == "__main__":
    unittest.main()
