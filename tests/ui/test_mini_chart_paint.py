"""Rendered miniature charts preserve candle output and comparable price positions."""

from PySide6.QtCore import QRect, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter
from PySide6.QtWidgets import QFrame, QStyleOptionViewItem
from shiboken6 import delete

from stockwidget.core.quote_presentation import MiniChartCell
from stockwidget.ui.controls.mini_chart import (
    AXIS_COLOR,
    MiniChartDelegate,
    amplitude_x,
    mini_chart_width,
    paint_amplitude_axis,
)
from stockwidget.ui.controls.quote_view import KLineDelegate, QuoteTableView, SimpleTableModel
from tests.support import QtTestCase


class MiniChartPaintTests(QtTestCase):
    def setUp(self):
        self.delegate = MiniChartDelegate()
        self.delegate.set_colors(False, QColor("#ffffff"), QColor("#f00000"),
                                 QColor("#00d000"), QColor("#808080"))

    def render(self, cell, width=100, height=30):
        image = QImage(width, height, QImage.Format_ARGB32_Premultiplied)
        image.fill(Qt.black)
        painter = QPainter(image)
        self.delegate.paint_cell(painter, QRect(0, 0, width, height), cell)
        painter.end()
        return image

    @staticmethod
    def pixels(image, color):
        return [(x, y) for y in range(image.height()) for x in range(image.width())
                if image.pixelColor(x, y) == color]

    def test_daily_candle_output_is_identical_to_the_existing_renderer(self):
        for candle in ((10, 11, 12, 9, 10), (11, 9, 12, 8, 10), (10, 10, 10, 10, 10)):
            with self.subTest(candle=candle):
                model = SimpleTableModel([[{"k": candle}]], ["迷你图"])
                old = KLineDelegate()
                old.set_colors(False, self.delegate.text_color, self.delegate.up_color,
                               self.delegate.down_color, self.delegate.neutral_color)
                image = QImage(30, 30, QImage.Format_ARGB32_Premultiplied)
                image.fill(Qt.black)
                painter = QPainter(image)
                option = QStyleOptionViewItem()
                option.rect = QRect(0, 0, 30, 30)
                old.paint(painter, option, model.index(0, 0))
                painter.end()
                self.assertEqual(image, self.render(MiniChartCell("kline", candle), 30))

    def test_trend_draws_the_price_line_and_a_gray_reference_in_both_color_modes(self):
        cell = MiniChartCell("trend", (100, 102, 102, 100, 100), ((0, 101), (1, 102)))
        for unicolor in (False, True):
            with self.subTest(unicolor=unicolor):
                self.delegate.unicolor = unicolor
                image = self.render(cell)
                color = self.delegate.text_color if unicolor else self.delegate.up_color
                self.assertTrue(self.pixels(image, color))
                gray_reference = [image.pixelColor(x, y) for x in range(10, image.width())
                                  for y in range(image.height() - 4, image.height())]
                self.assertTrue(any(0 < c.red() == c.green() == c.blue() < 200 for c in gray_reference))

    def test_whole_trend_uses_current_quote_direction_even_when_last_sample_differs(self):
        strokes = []

        class Recorder(QPainter):
            def drawLine(self, *args):
                if self.pen().style() == Qt.SolidLine:
                    strokes.append(self.pen().color())
                return super().drawLine(*args)

        for current, expected in ((102, self.delegate.up_color), (98, self.delegate.down_color),
                                  (100, self.delegate.neutral_color)):
            for unicolor in (False, True):
                with self.subTest(current=current, unicolor=unicolor):
                    self.delegate.unicolor = unicolor
                    strokes.clear()
                    image = QImage(100, 30, QImage.Format_ARGB32_Premultiplied)
                    image.fill(Qt.black)
                    painter = Recorder(image)
                    self.delegate.paint_cell(painter, QRect(0, 0, 100, 30),
                                             MiniChartCell("trend", (100, current, 102, 98, 100),
                                                           ((0, 98), (.5, 102), (1, 99))))
                    painter.end()
                    color = self.delegate.text_color if unicolor else expected
                    self.assertTrue(self.pixels(image, color))
                    self.assertEqual(strokes, [color, color])

    def test_trend_without_previous_close_draws_neutral_prices_without_a_false_axis(self):
        calls = []

        class Recorder(QPainter):
            def drawLine(self, *args):
                calls.append((self.pen().color(), self.pen().style()))
                return super().drawLine(*args)

        for reference in (0, -1, float("nan")):
            with self.subTest(reference=reference):
                calls.clear()
                image = QImage(100, 30, QImage.Format_ARGB32_Premultiplied)
                image.fill(Qt.black)
                painter = Recorder(image)
                cell = MiniChartCell("trend", (100, 102, 102, 98, reference), ((0, 98), (1, 102)))
                self.delegate.paint_cell(painter, QRect(0, 0, 100, 30), cell)
                painter.end()
                self.assertTrue(calls)
                self.assertTrue(all(color == self.delegate.neutral_color and style == Qt.SolidLine
                                    for color, style in calls))
                self.assertTrue(self.pixels(image, self.delegate.neutral_color))
                self.assertFalse(self.pixels(image, self.delegate.up_color))
                self.assertFalse(self.pixels(image, self.delegate.down_color))

    def test_trend_has_no_point_markers_even_for_a_single_sample(self):
        ellipses = []

        class Recorder(QPainter):
            def drawEllipse(self, *args):
                ellipses.append(args)
                return super().drawEllipse(*args)

        for points in (((0, 100),), ((0, 100), (0.001, 99)), ((0, 100), (1, 99))):
            image = QImage(100, 30, QImage.Format_ARGB32_Premultiplied)
            image.fill(Qt.black)
            painter = Recorder(image)
            self.delegate.paint_cell(painter, QRect(0, 0, 100, 30),
                                     MiniChartCell("trend", (100, 99, 100, 99, 100), points))
            painter.end()
        self.assertEqual(ellipses, [])

    def test_font_scaling_preserves_amplitude_dot_ratio_and_fits_tight_rows_at_each_dpi(self):
        strokes, radii = [], []

        class Recorder(QPainter):
            def drawLine(self, *args):
                if self.pen().style() == Qt.SolidLine:
                    strokes.append(self.pen().widthF())
                return super().drawLine(*args)

            def drawEllipse(self, center, rx, ry):
                radii.append((rx, ry))
                return super().drawEllipse(center, rx, ry)

        self.delegate.unicolor = True
        for ratio in (1, 1.5, 2):
            for height in (8, 30, 60):
                previous = {"amplitude": 0, "trend": 0}
                for size in (8, 12, 24):
                    for kind in ("amplitude", "trend"):
                        with self.subTest(ratio=ratio, height=height, size=size, kind=kind):
                            strokes.clear()
                            radii.clear()
                            image = QImage(round(200 * ratio), round(height * ratio), QImage.Format_ARGB32_Premultiplied)
                            image.setDevicePixelRatio(ratio)
                            image.fill(Qt.black)
                            painter = Recorder(image)
                            painter.setFont(QFont("Arial", size))
                            cell = MiniChartCell(kind, (100, 100, 104, 96, 100),
                                                 ((0, 96), (1, 104)), (-5, 5))
                            self.delegate.paint_cell(painter, QRect(0, 0, 200, height), cell)
                            painter.end()
                            stroke = strokes[-1]
                            self.assertGreaterEqual(stroke, previous[kind])
                            if height == 60:
                                self.assertGreater(stroke, previous[kind])
                            previous[kind] = stroke
                            if kind == "amplitude":
                                self.assertAlmostEqual(radii[0][0] / stroke, 1.5)
                                self.assertEqual(radii[0][0], radii[0][1])
                                self.assertLessEqual(radii[0][0] * 2, height - 4)
                            else:
                                self.assertFalse(radii)
                            self.assertTrue(any(image.pixelColor(x, y).value() > 200
                                                for x in range(image.width()) for y in range(image.height())))

    def test_amplitude_uses_column_limits_and_the_current_dot_matches_its_gradient(self):
        limits = (-10, 10)
        for low, current, high in ((-8, -4, 3), (-2, 2, 7), (1, 1, 1)):
            with self.subTest(low=low, current=current, high=high):
                candle = (100, 100 + current, 100 + high, 100 + low, 100)
                cell = MiniChartCell("amplitude", candle, limits=limits)
                image = self.render(cell)
                x = amplitude_x(QRect(0, 0, 100, 30), current, limits)
                # The solid dot covers more than the two-pixel range stroke.
                for y in (14, 15):
                    self.assertEqual(image.pixelColor(round(x), y), self.delegate.amplitude_color(current))
                for value in (low, high):
                    endpoint = amplitude_x(QRect(0, 0, 100, 30), value, limits)
                    self.assertNotEqual(image.pixelColor(round(endpoint), 15), QColor(Qt.black))
        # Color responds to absolute change, so equal changes in two columns agree.
        self.assertGreater(self.delegate.amplitude_color(3).red(), self.delegate.amplitude_color(1).red())
        first_step = self.delegate.amplitude_color(1).red() - self.delegate.amplitude_color(0).red()
        later_step = self.delegate.amplitude_color(4).red() - self.delegate.amplitude_color(3).red()
        self.assertGreater(first_step, later_step)

    def test_range_without_zero_uses_the_full_width_and_has_no_reference_axis(self):
        image = QImage(100, 90, QImage.Format_ARGB32_Premultiplied)
        image.fill(Qt.black)
        painter = QPainter(image)
        paint_amplitude_axis(painter, QRectF(0, 0, 100, 90), (1, 5))
        painter.end()
        self.assertFalse(self.pixels(image, AXIS_COLOR))
        self.assertEqual(amplitude_x(QRect(0, 0, 100, 30), 1, (1, 5)), 4)
        self.assertEqual(amplitude_x(QRect(0, 0, 100, 30), 5, (1, 5)), 96)

    def test_chart_column_width_tracks_ten_characters_when_font_size_changes(self):
        for size in (8, 12, 18):
            font = QFont("Arial", size)
            for kind in ("trend", "amplitude"):
                self.assertEqual(mini_chart_width(font, kind), QFontMetrics(font).horizontalAdvance("0000000000"))

    def test_table_zero_axis_continues_between_rows_without_recoloring_current_dot(self):
        table = QuoteTableView()
        self.addCleanup(delete, table)
        table.setStyleSheet("QTableView { background-color: #000000; border: none; }")
        table.setFrameShape(QFrame.NoFrame)
        table.horizontalHeader().hide()
        table.verticalHeader().hide()
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setShowGrid(False)
        cells = [MiniChartCell("amplitude", (100, 100, 103, 97, 100), limits=(-5, 5)) for _ in range(3)]
        model = SimpleTableModel([[cell] for cell in cells], ["迷你图"], parent=table)
        table.setModel(model)
        table.setItemDelegateForColumn(0, self.delegate)
        table.setColumnWidth(0, 100)
        for row in range(3):
            table.setRowHeight(row, 30)
        table.resize(100, 90)
        table.show()
        self.app.processEvents()
        image = table.viewport().grab().toImage()
        scale = image.devicePixelRatio()
        first = QRectF(table.visualRect(model.index(0, 0)))
        x = round(amplitude_x(first, 0, (-5, 5)) * scale)
        for border in (30, 60):
            ink = [image.pixelColor(px, round(y * scale)) for px in (x - 1, x, x + 1)
                   for y in range(border - 2, border + 3)]
            self.assertTrue(any(color.red() >= 64 for color in ink))
        for row in range(3):
            rect = QRectF(table.visualRect(model.index(row, 0)))
            center = round(rect.center().y() * scale)
            self.assertEqual(image.pixelColor(x, center), self.delegate.amplitude_color(0))
