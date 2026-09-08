# -*- coding: utf-8 -*-
"""统一颜色和方向颜色角色的测试。"""

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from stockwidget.ui.table_model import (
    COLOR_ROLE_DOWN,
    COLOR_ROLE_NEUTRAL,
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
    KLineDelegate,
    SimpleTableModel,
    direction_color_role,
)


class ColorSystemTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])

    def test_direction_role_mapping(self):
        self.assertEqual(direction_color_role(1), COLOR_ROLE_UP)
        self.assertEqual(direction_color_role(-1), COLOR_ROLE_DOWN)
        self.assertEqual(direction_color_role(0), COLOR_ROLE_NEUTRAL)

    def test_table_uses_separate_roles_then_unifies_to_text_color(self):
        model = SimpleTableModel()
        model.set_rows_headers(
            [["普通", "+1", "-1", "0"]],
            ["普通", "上涨", "下跌", "中性"],
            [[COLOR_ROLE_TEXT, COLOR_ROLE_UP, COLOR_ROLE_DOWN, COLOR_ROLE_NEUTRAL]],
        )
        colors = (
            QColor("#112233"),
            QColor("#aa0000"),
            QColor("#00aa00"),
            QColor("#777777"),
        )
        model.set_colors(False, *colors)

        actual = [
            model.data(model.index(0, column), Qt.ItemDataRole.ForegroundRole).name()
            for column in range(4)
        ]
        self.assertEqual(actual, [color.name() for color in colors])
        self.assertEqual(
            model.headerData(0, Qt.Orientation.Horizontal, Qt.ItemDataRole.ForegroundRole).name(),
            colors[0].name(),
        )

        model.set_colors(True, *colors)
        unified = [
            model.data(model.index(0, column), Qt.ItemDataRole.ForegroundRole).name()
            for column in range(4)
        ]
        self.assertEqual(unified, [colors[0].name()] * 4)

    def test_kline_uses_direction_colors_and_unifies_when_enabled(self):
        delegate = KLineDelegate()
        colors = (
            QColor("#112233"),
            QColor("#aa0000"),
            QColor("#00aa00"),
            QColor("#777777"),
        )
        delegate.set_colors(False, *colors)

        self.assertEqual(delegate.candle_color(1, 2).name(), colors[1].name())
        self.assertEqual(delegate.candle_color(2, 1).name(), colors[2].name())
        self.assertEqual(delegate.candle_color(1, 1).name(), colors[3].name())
        self.assertEqual(delegate.reference_color().name(), colors[3].name())

        delegate.set_colors(True, *colors)
        self.assertEqual(delegate.candle_color(1, 2).name(), colors[0].name())
        self.assertEqual(delegate.reference_color().name(), colors[0].name())


if __name__ == "__main__":
    unittest.main()
