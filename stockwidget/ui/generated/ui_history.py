# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'history.ui'
##
## Created by: Qt User Interface Compiler version 6.11.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QDialog, QGridLayout, QHBoxLayout,
    QLabel, QPushButton, QSizePolicy, QSpacerItem,
    QVBoxLayout, QWidget)

from stockwidget.ui.controls.history_chart import HistoryChart
from stockwidget.ui.controls.history_details import (HistoryQuoteHeader, OrderBookWidget)
from stockwidget.ui.controls.settings_widgets import FilledCheckBox

class Ui_HistoryDialog(object):
    def setupUi(self, HistoryDialog):
        if not HistoryDialog.objectName():
            HistoryDialog.setObjectName(u"HistoryDialog")
        HistoryDialog.resize(480, 320)
        self.chart_layout = QVBoxLayout(HistoryDialog)
        self.chart_layout.setSpacing(4)
        self.chart_layout.setObjectName(u"chart_layout")
        self.chart_layout.setContentsMargins(8, 8, 8, 8)
        self.title_row = QWidget(HistoryDialog)
        self.title_row.setObjectName(u"title_row")
        self.title_layout = QHBoxLayout(self.title_row)
        self.title_layout.setSpacing(4)
        self.title_layout.setObjectName(u"title_layout")
        self.title_layout.setContentsMargins(0, 0, 0, 0)
        self.title = QLabel(self.title_row)
        self.title.setObjectName(u"title")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.title.sizePolicy().hasHeightForWidth())
        self.title.setSizePolicy(sizePolicy)

        self.title_layout.addWidget(self.title)

        self.pop_out_button = QPushButton(self.title_row)
        self.pop_out_button.setObjectName(u"pop_out_button")
        self.pop_out_button.setAutoDefault(False)
        self.pop_out_button.setMinimumSize(QSize(26, 26))
        self.pop_out_button.setMaximumSize(QSize(26, 26))

        self.title_layout.addWidget(self.pop_out_button)

        self.close_button = QPushButton(self.title_row)
        self.close_button.setObjectName(u"close_button")
        self.close_button.setAutoDefault(False)
        self.close_button.setMinimumSize(QSize(26, 26))
        self.close_button.setMaximumSize(QSize(26, 26))
        self.close_button.setProperty(u"destructive", True)

        self.title_layout.addWidget(self.close_button)


        self.chart_layout.addWidget(self.title_row)

        self.quote_header = HistoryQuoteHeader(HistoryDialog)
        self.quote_header.setObjectName(u"quote_header")
        self.quote_header_layout = QHBoxLayout(self.quote_header)
        self.quote_header_layout.setSpacing(8)
        self.quote_header_layout.setObjectName(u"quote_header_layout")
        self.quote_header_layout.setContentsMargins(0, 0, 0, 6)
        self.quote_primary = QWidget(self.quote_header)
        self.quote_primary.setObjectName(u"quote_primary")
        self.quote_primary.setMinimumSize(QSize(140, 0))
        self.quote_primary_layout = QVBoxLayout(self.quote_primary)
        self.quote_primary_layout.setSpacing(3)
        self.quote_primary_layout.setObjectName(u"quote_primary_layout")
        self.quote_primary_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_name = QLabel(self.quote_primary)
        self.quote_name.setObjectName(u"quote_name")
        sizePolicy.setHeightForWidth(self.quote_name.sizePolicy().hasHeightForWidth())
        self.quote_name.setSizePolicy(sizePolicy)
        font = QFont()
        font.setPointSize(11)
        font.setBold(True)
        self.quote_name.setFont(font)

        self.quote_primary_layout.addWidget(self.quote_name)

        self.quote_code = QLabel(self.quote_primary)
        self.quote_code.setObjectName(u"quote_code")
        self.quote_code.setProperty(u"settingDescription", True)

        self.quote_primary_layout.addWidget(self.quote_code)

        self.quote_price_layout = QHBoxLayout()
        self.quote_price_layout.setSpacing(10)
        self.quote_price_layout.setObjectName(u"quote_price_layout")
        self.quote_current = QLabel(self.quote_primary)
        self.quote_current.setObjectName(u"quote_current")
        font1 = QFont()
        font1.setPointSize(18)
        font1.setBold(True)
        self.quote_current.setFont(font1)

        self.quote_price_layout.addWidget(self.quote_current)

        self.quote_changes_layout = QVBoxLayout()
        self.quote_changes_layout.setSpacing(3)
        self.quote_changes_layout.setObjectName(u"quote_changes_layout")
        self.quote_change = QLabel(self.quote_primary)
        self.quote_change.setObjectName(u"quote_change")
        font2 = QFont()
        font2.setBold(True)
        self.quote_change.setFont(font2)

        self.quote_changes_layout.addWidget(self.quote_change)

        self.quote_change_pct = QLabel(self.quote_primary)
        self.quote_change_pct.setObjectName(u"quote_change_pct")
        self.quote_change_pct.setFont(font2)

        self.quote_changes_layout.addWidget(self.quote_change_pct)


        self.quote_price_layout.addLayout(self.quote_changes_layout)

        self.quote_price_space = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.quote_price_layout.addItem(self.quote_price_space)


        self.quote_primary_layout.addLayout(self.quote_price_layout)


        self.quote_header_layout.addWidget(self.quote_primary)

        self.quote_navigation = QWidget(self.quote_header)
        self.quote_navigation.setObjectName(u"quote_navigation")
        self.quote_navigation_layout = QVBoxLayout(self.quote_navigation)
        self.quote_navigation_layout.setSpacing(5)
        self.quote_navigation_layout.setObjectName(u"quote_navigation_layout")
        self.quote_navigation_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_previous = QPushButton(self.quote_navigation)
        self.btn_previous.setObjectName(u"btn_previous")
        self.btn_previous.setAutoDefault(False)
        self.btn_previous.setMinimumSize(QSize(26, 26))
        self.btn_previous.setMaximumSize(QSize(26, 26))

        self.quote_navigation_layout.addWidget(self.btn_previous)

        self.btn_next = QPushButton(self.quote_navigation)
        self.btn_next.setObjectName(u"btn_next")
        self.btn_next.setAutoDefault(False)
        self.btn_next.setMinimumSize(QSize(26, 26))
        self.btn_next.setMaximumSize(QSize(26, 26))

        self.quote_navigation_layout.addWidget(self.btn_next)


        self.quote_header_layout.addWidget(self.quote_navigation)

        self.quote_column_0 = QWidget(self.quote_header)
        self.quote_column_0.setObjectName(u"quote_column_0")
        self.quote_column_0_layout = QGridLayout(self.quote_column_0)
        self.quote_column_0_layout.setSpacing(4)
        self.quote_column_0_layout.setObjectName(u"quote_column_0_layout")
        self.quote_column_0_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_high_caption = QLabel(self.quote_column_0)
        self.quote_high_caption.setObjectName(u"quote_high_caption")
        self.quote_high_caption.setProperty(u"settingDescription", True)

        self.quote_column_0_layout.addWidget(self.quote_high_caption, 0, 0, 1, 1)

        self.quote_high = QLabel(self.quote_column_0)
        self.quote_high.setObjectName(u"quote_high")
        self.quote_high.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1 = QSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.quote_high.sizePolicy().hasHeightForWidth())
        self.quote_high.setSizePolicy(sizePolicy1)
        self.quote_high.setFont(font2)
        self.quote_high.setMinimumSize(QSize(40, 0))

        self.quote_column_0_layout.addWidget(self.quote_high, 0, 1, 1, 1)

        self.quote_low_caption = QLabel(self.quote_column_0)
        self.quote_low_caption.setObjectName(u"quote_low_caption")
        self.quote_low_caption.setProperty(u"settingDescription", True)

        self.quote_column_0_layout.addWidget(self.quote_low_caption, 1, 0, 1, 1)

        self.quote_low = QLabel(self.quote_column_0)
        self.quote_low.setObjectName(u"quote_low")
        self.quote_low.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_low.sizePolicy().hasHeightForWidth())
        self.quote_low.setSizePolicy(sizePolicy1)
        self.quote_low.setFont(font2)
        self.quote_low.setMinimumSize(QSize(40, 0))

        self.quote_column_0_layout.addWidget(self.quote_low, 1, 1, 1, 1)

        self.quote_open_caption = QLabel(self.quote_column_0)
        self.quote_open_caption.setObjectName(u"quote_open_caption")
        self.quote_open_caption.setProperty(u"settingDescription", True)

        self.quote_column_0_layout.addWidget(self.quote_open_caption, 2, 0, 1, 1)

        self.quote_open = QLabel(self.quote_column_0)
        self.quote_open.setObjectName(u"quote_open")
        self.quote_open.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_open.sizePolicy().hasHeightForWidth())
        self.quote_open.setSizePolicy(sizePolicy1)
        self.quote_open.setFont(font2)
        self.quote_open.setMinimumSize(QSize(40, 0))

        self.quote_column_0_layout.addWidget(self.quote_open, 2, 1, 1, 1)

        self.quote_prev_close_caption = QLabel(self.quote_column_0)
        self.quote_prev_close_caption.setObjectName(u"quote_prev_close_caption")
        self.quote_prev_close_caption.setProperty(u"settingDescription", True)

        self.quote_column_0_layout.addWidget(self.quote_prev_close_caption, 3, 0, 1, 1)

        self.quote_prev_close = QLabel(self.quote_column_0)
        self.quote_prev_close.setObjectName(u"quote_prev_close")
        self.quote_prev_close.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_prev_close.sizePolicy().hasHeightForWidth())
        self.quote_prev_close.setSizePolicy(sizePolicy1)
        self.quote_prev_close.setFont(font2)
        self.quote_prev_close.setMinimumSize(QSize(40, 0))

        self.quote_column_0_layout.addWidget(self.quote_prev_close, 3, 1, 1, 1)

        self.quote_column_0_layout.setColumnStretch(1, 1)

        self.quote_header_layout.addWidget(self.quote_column_0)

        self.quote_column_1 = QWidget(self.quote_header)
        self.quote_column_1.setObjectName(u"quote_column_1")
        self.quote_column_1_layout = QGridLayout(self.quote_column_1)
        self.quote_column_1_layout.setSpacing(4)
        self.quote_column_1_layout.setObjectName(u"quote_column_1_layout")
        self.quote_column_1_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_inside_caption = QLabel(self.quote_column_1)
        self.quote_inside_caption.setObjectName(u"quote_inside_caption")
        self.quote_inside_caption.setProperty(u"settingDescription", True)

        self.quote_column_1_layout.addWidget(self.quote_inside_caption, 0, 0, 1, 1)

        self.quote_inside = QLabel(self.quote_column_1)
        self.quote_inside.setObjectName(u"quote_inside")
        self.quote_inside.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_inside.sizePolicy().hasHeightForWidth())
        self.quote_inside.setSizePolicy(sizePolicy1)
        self.quote_inside.setFont(font2)
        self.quote_inside.setMinimumSize(QSize(40, 0))

        self.quote_column_1_layout.addWidget(self.quote_inside, 0, 1, 1, 1)

        self.quote_outside_caption = QLabel(self.quote_column_1)
        self.quote_outside_caption.setObjectName(u"quote_outside_caption")
        self.quote_outside_caption.setProperty(u"settingDescription", True)

        self.quote_column_1_layout.addWidget(self.quote_outside_caption, 1, 0, 1, 1)

        self.quote_outside = QLabel(self.quote_column_1)
        self.quote_outside.setObjectName(u"quote_outside")
        self.quote_outside.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_outside.sizePolicy().hasHeightForWidth())
        self.quote_outside.setSizePolicy(sizePolicy1)
        self.quote_outside.setFont(font2)
        self.quote_outside.setMinimumSize(QSize(40, 0))

        self.quote_column_1_layout.addWidget(self.quote_outside, 1, 1, 1, 1)

        self.quote_volume_caption = QLabel(self.quote_column_1)
        self.quote_volume_caption.setObjectName(u"quote_volume_caption")
        self.quote_volume_caption.setProperty(u"settingDescription", True)

        self.quote_column_1_layout.addWidget(self.quote_volume_caption, 2, 0, 1, 1)

        self.quote_volume = QLabel(self.quote_column_1)
        self.quote_volume.setObjectName(u"quote_volume")
        self.quote_volume.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_volume.sizePolicy().hasHeightForWidth())
        self.quote_volume.setSizePolicy(sizePolicy1)
        self.quote_volume.setFont(font2)
        self.quote_volume.setMinimumSize(QSize(40, 0))

        self.quote_column_1_layout.addWidget(self.quote_volume, 2, 1, 1, 1)

        self.quote_amount_caption = QLabel(self.quote_column_1)
        self.quote_amount_caption.setObjectName(u"quote_amount_caption")
        self.quote_amount_caption.setProperty(u"settingDescription", True)

        self.quote_column_1_layout.addWidget(self.quote_amount_caption, 3, 0, 1, 1)

        self.quote_amount = QLabel(self.quote_column_1)
        self.quote_amount.setObjectName(u"quote_amount")
        self.quote_amount.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_amount.sizePolicy().hasHeightForWidth())
        self.quote_amount.setSizePolicy(sizePolicy1)
        self.quote_amount.setFont(font2)
        self.quote_amount.setMinimumSize(QSize(40, 0))

        self.quote_column_1_layout.addWidget(self.quote_amount, 3, 1, 1, 1)

        self.quote_column_1_layout.setColumnStretch(1, 1)

        self.quote_header_layout.addWidget(self.quote_column_1)

        self.quote_column_2 = QWidget(self.quote_header)
        self.quote_column_2.setObjectName(u"quote_column_2")
        self.quote_column_2_layout = QGridLayout(self.quote_column_2)
        self.quote_column_2_layout.setSpacing(4)
        self.quote_column_2_layout.setObjectName(u"quote_column_2_layout")
        self.quote_column_2_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_average_caption = QLabel(self.quote_column_2)
        self.quote_average_caption.setObjectName(u"quote_average_caption")
        self.quote_average_caption.setProperty(u"settingDescription", True)

        self.quote_column_2_layout.addWidget(self.quote_average_caption, 0, 0, 1, 1)

        self.quote_average = QLabel(self.quote_column_2)
        self.quote_average.setObjectName(u"quote_average")
        self.quote_average.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_average.sizePolicy().hasHeightForWidth())
        self.quote_average.setSizePolicy(sizePolicy1)
        self.quote_average.setFont(font2)
        self.quote_average.setMinimumSize(QSize(40, 0))

        self.quote_column_2_layout.addWidget(self.quote_average, 0, 1, 1, 1)

        self.quote_turnover_caption = QLabel(self.quote_column_2)
        self.quote_turnover_caption.setObjectName(u"quote_turnover_caption")
        self.quote_turnover_caption.setProperty(u"settingDescription", True)

        self.quote_column_2_layout.addWidget(self.quote_turnover_caption, 1, 0, 1, 1)

        self.quote_turnover = QLabel(self.quote_column_2)
        self.quote_turnover.setObjectName(u"quote_turnover")
        self.quote_turnover.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_turnover.sizePolicy().hasHeightForWidth())
        self.quote_turnover.setSizePolicy(sizePolicy1)
        self.quote_turnover.setFont(font2)
        self.quote_turnover.setMinimumSize(QSize(40, 0))

        self.quote_column_2_layout.addWidget(self.quote_turnover, 1, 1, 1, 1)

        self.quote_amplitude_caption = QLabel(self.quote_column_2)
        self.quote_amplitude_caption.setObjectName(u"quote_amplitude_caption")
        self.quote_amplitude_caption.setProperty(u"settingDescription", True)

        self.quote_column_2_layout.addWidget(self.quote_amplitude_caption, 2, 0, 1, 1)

        self.quote_amplitude = QLabel(self.quote_column_2)
        self.quote_amplitude.setObjectName(u"quote_amplitude")
        self.quote_amplitude.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_amplitude.sizePolicy().hasHeightForWidth())
        self.quote_amplitude.setSizePolicy(sizePolicy1)
        self.quote_amplitude.setFont(font2)
        self.quote_amplitude.setMinimumSize(QSize(40, 0))

        self.quote_column_2_layout.addWidget(self.quote_amplitude, 2, 1, 1, 1)

        self.quote_volume_ratio_caption = QLabel(self.quote_column_2)
        self.quote_volume_ratio_caption.setObjectName(u"quote_volume_ratio_caption")
        self.quote_volume_ratio_caption.setProperty(u"settingDescription", True)

        self.quote_column_2_layout.addWidget(self.quote_volume_ratio_caption, 3, 0, 1, 1)

        self.quote_volume_ratio = QLabel(self.quote_column_2)
        self.quote_volume_ratio.setObjectName(u"quote_volume_ratio")
        self.quote_volume_ratio.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_volume_ratio.sizePolicy().hasHeightForWidth())
        self.quote_volume_ratio.setSizePolicy(sizePolicy1)
        self.quote_volume_ratio.setFont(font2)
        self.quote_volume_ratio.setMinimumSize(QSize(40, 0))

        self.quote_column_2_layout.addWidget(self.quote_volume_ratio, 3, 1, 1, 1)

        self.quote_column_2_layout.setColumnStretch(1, 1)

        self.quote_header_layout.addWidget(self.quote_column_2)

        self.quote_column_3 = QWidget(self.quote_header)
        self.quote_column_3.setObjectName(u"quote_column_3")
        self.quote_column_3_layout = QGridLayout(self.quote_column_3)
        self.quote_column_3_layout.setSpacing(4)
        self.quote_column_3_layout.setObjectName(u"quote_column_3_layout")
        self.quote_column_3_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_pe_ttm_caption = QLabel(self.quote_column_3)
        self.quote_pe_ttm_caption.setObjectName(u"quote_pe_ttm_caption")
        self.quote_pe_ttm_caption.setProperty(u"settingDescription", True)

        self.quote_column_3_layout.addWidget(self.quote_pe_ttm_caption, 2, 0, 1, 1)

        self.quote_pe_ttm = QLabel(self.quote_column_3)
        self.quote_pe_ttm.setObjectName(u"quote_pe_ttm")
        self.quote_pe_ttm.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_pe_ttm.sizePolicy().hasHeightForWidth())
        self.quote_pe_ttm.setSizePolicy(sizePolicy1)
        self.quote_pe_ttm.setFont(font2)
        self.quote_pe_ttm.setMinimumSize(QSize(40, 0))

        self.quote_column_3_layout.addWidget(self.quote_pe_ttm, 2, 1, 1, 1)

        self.quote_pe_static_caption = QLabel(self.quote_column_3)
        self.quote_pe_static_caption.setObjectName(u"quote_pe_static_caption")
        self.quote_pe_static_caption.setProperty(u"settingDescription", True)

        self.quote_column_3_layout.addWidget(self.quote_pe_static_caption, 0, 0, 1, 1)

        self.quote_pe_static = QLabel(self.quote_column_3)
        self.quote_pe_static.setObjectName(u"quote_pe_static")
        self.quote_pe_static.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_pe_static.sizePolicy().hasHeightForWidth())
        self.quote_pe_static.setSizePolicy(sizePolicy1)
        self.quote_pe_static.setFont(font2)
        self.quote_pe_static.setMinimumSize(QSize(40, 0))

        self.quote_column_3_layout.addWidget(self.quote_pe_static, 0, 1, 1, 1)

        self.quote_pb_caption = QLabel(self.quote_column_3)
        self.quote_pb_caption.setObjectName(u"quote_pb_caption")
        self.quote_pb_caption.setProperty(u"settingDescription", True)

        self.quote_column_3_layout.addWidget(self.quote_pb_caption, 3, 0, 1, 1)

        self.quote_pb = QLabel(self.quote_column_3)
        self.quote_pb.setObjectName(u"quote_pb")
        self.quote_pb.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_pb.sizePolicy().hasHeightForWidth())
        self.quote_pb.setSizePolicy(sizePolicy1)
        self.quote_pb.setFont(font2)
        self.quote_pb.setMinimumSize(QSize(40, 0))

        self.quote_column_3_layout.addWidget(self.quote_pb, 3, 1, 1, 1)

        self.quote_pe_dynamic_caption = QLabel(self.quote_column_3)
        self.quote_pe_dynamic_caption.setObjectName(u"quote_pe_dynamic_caption")
        self.quote_pe_dynamic_caption.setProperty(u"settingDescription", True)

        self.quote_column_3_layout.addWidget(self.quote_pe_dynamic_caption, 1, 0, 1, 1)

        self.quote_pe_dynamic = QLabel(self.quote_column_3)
        self.quote_pe_dynamic.setObjectName(u"quote_pe_dynamic")
        self.quote_pe_dynamic.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_pe_dynamic.sizePolicy().hasHeightForWidth())
        self.quote_pe_dynamic.setSizePolicy(sizePolicy1)
        self.quote_pe_dynamic.setFont(font2)
        self.quote_pe_dynamic.setMinimumSize(QSize(40, 0))

        self.quote_column_3_layout.addWidget(self.quote_pe_dynamic, 1, 1, 1, 1)

        self.quote_column_3_layout.setColumnStretch(1, 1)

        self.quote_header_layout.addWidget(self.quote_column_3)

        self.quote_column_4 = QWidget(self.quote_header)
        self.quote_column_4.setObjectName(u"quote_column_4")
        self.quote_column_4_layout = QGridLayout(self.quote_column_4)
        self.quote_column_4_layout.setSpacing(4)
        self.quote_column_4_layout.setObjectName(u"quote_column_4_layout")
        self.quote_column_4_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_market_cap_caption = QLabel(self.quote_column_4)
        self.quote_market_cap_caption.setObjectName(u"quote_market_cap_caption")
        self.quote_market_cap_caption.setProperty(u"settingDescription", True)

        self.quote_column_4_layout.addWidget(self.quote_market_cap_caption, 0, 0, 1, 1)

        self.quote_market_cap = QLabel(self.quote_column_4)
        self.quote_market_cap.setObjectName(u"quote_market_cap")
        self.quote_market_cap.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_market_cap.sizePolicy().hasHeightForWidth())
        self.quote_market_cap.setSizePolicy(sizePolicy1)
        self.quote_market_cap.setFont(font2)
        self.quote_market_cap.setMinimumSize(QSize(40, 0))

        self.quote_column_4_layout.addWidget(self.quote_market_cap, 0, 1, 1, 1)

        self.quote_total_shares_caption = QLabel(self.quote_column_4)
        self.quote_total_shares_caption.setObjectName(u"quote_total_shares_caption")
        self.quote_total_shares_caption.setProperty(u"settingDescription", True)

        self.quote_column_4_layout.addWidget(self.quote_total_shares_caption, 1, 0, 1, 1)

        self.quote_total_shares = QLabel(self.quote_column_4)
        self.quote_total_shares.setObjectName(u"quote_total_shares")
        self.quote_total_shares.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_total_shares.sizePolicy().hasHeightForWidth())
        self.quote_total_shares.setSizePolicy(sizePolicy1)
        self.quote_total_shares.setFont(font2)
        self.quote_total_shares.setMinimumSize(QSize(40, 0))

        self.quote_column_4_layout.addWidget(self.quote_total_shares, 1, 1, 1, 1)

        self.quote_float_market_cap_caption = QLabel(self.quote_column_4)
        self.quote_float_market_cap_caption.setObjectName(u"quote_float_market_cap_caption")
        self.quote_float_market_cap_caption.setProperty(u"settingDescription", True)

        self.quote_column_4_layout.addWidget(self.quote_float_market_cap_caption, 2, 0, 1, 1)

        self.quote_float_market_cap = QLabel(self.quote_column_4)
        self.quote_float_market_cap.setObjectName(u"quote_float_market_cap")
        self.quote_float_market_cap.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_float_market_cap.sizePolicy().hasHeightForWidth())
        self.quote_float_market_cap.setSizePolicy(sizePolicy1)
        self.quote_float_market_cap.setFont(font2)
        self.quote_float_market_cap.setMinimumSize(QSize(40, 0))

        self.quote_column_4_layout.addWidget(self.quote_float_market_cap, 2, 1, 1, 1)

        self.quote_float_shares_caption = QLabel(self.quote_column_4)
        self.quote_float_shares_caption.setObjectName(u"quote_float_shares_caption")
        self.quote_float_shares_caption.setProperty(u"settingDescription", True)

        self.quote_column_4_layout.addWidget(self.quote_float_shares_caption, 3, 0, 1, 1)

        self.quote_float_shares = QLabel(self.quote_column_4)
        self.quote_float_shares.setObjectName(u"quote_float_shares")
        self.quote_float_shares.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        sizePolicy1.setHeightForWidth(self.quote_float_shares.sizePolicy().hasHeightForWidth())
        self.quote_float_shares.setSizePolicy(sizePolicy1)
        self.quote_float_shares.setFont(font2)
        self.quote_float_shares.setMinimumSize(QSize(40, 0))

        self.quote_column_4_layout.addWidget(self.quote_float_shares, 3, 1, 1, 1)

        self.quote_column_4_layout.setColumnStretch(1, 1)

        self.quote_header_layout.addWidget(self.quote_column_4)

        self.quote_header_layout.setStretch(2, 1)
        self.quote_header_layout.setStretch(3, 1)
        self.quote_header_layout.setStretch(4, 1)
        self.quote_header_layout.setStretch(5, 1)
        self.quote_header_layout.setStretch(6, 1)

        self.chart_layout.addWidget(self.quote_header)

        self.toolbar = QWidget(HistoryDialog)
        self.toolbar.setObjectName(u"toolbar")
        self.toolbar_layout = QHBoxLayout(self.toolbar)
        self.toolbar_layout.setSpacing(2)
        self.toolbar_layout.setObjectName(u"toolbar_layout")
        self.toolbar_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_intraday = QPushButton(self.toolbar)
        self.btn_intraday.setObjectName(u"btn_intraday")
        self.btn_intraday.setCheckable(True)
        self.btn_intraday.setChecked(True)
        self.btn_intraday.setAutoDefault(False)
        self.btn_intraday.setMinimumSize(QSize(28, 24))
        self.btn_intraday.setMaximumSize(QSize(16777215, 24))
        self.btn_intraday.setProperty(u"chartTab", True)

        self.toolbar_layout.addWidget(self.btn_intraday)

        self.btn_five_day = QPushButton(self.toolbar)
        self.btn_five_day.setObjectName(u"btn_five_day")
        self.btn_five_day.setCheckable(True)
        self.btn_five_day.setAutoDefault(False)
        self.btn_five_day.setMinimumSize(QSize(28, 24))
        self.btn_five_day.setMaximumSize(QSize(16777215, 24))
        self.btn_five_day.setProperty(u"chartTab", True)

        self.toolbar_layout.addWidget(self.btn_five_day)

        self.btn_daily = QPushButton(self.toolbar)
        self.btn_daily.setObjectName(u"btn_daily")
        self.btn_daily.setCheckable(True)
        self.btn_daily.setAutoDefault(False)
        self.btn_daily.setMinimumSize(QSize(28, 24))
        self.btn_daily.setMaximumSize(QSize(16777215, 24))
        self.btn_daily.setProperty(u"chartTab", True)

        self.toolbar_layout.addWidget(self.btn_daily)

        self.toolbar_space = QSpacerItem(8, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.toolbar_layout.addItem(self.toolbar_space)

        self.cb_volume = FilledCheckBox(self.toolbar)
        self.cb_volume.setObjectName(u"cb_volume")

        self.toolbar_layout.addWidget(self.cb_volume)

        self.cb_average = FilledCheckBox(self.toolbar)
        self.cb_average.setObjectName(u"cb_average")

        self.toolbar_layout.addWidget(self.cb_average)

        self.cb_order_book = FilledCheckBox(self.toolbar)
        self.cb_order_book.setObjectName(u"cb_order_book")

        self.toolbar_layout.addWidget(self.cb_order_book)


        self.chart_layout.addWidget(self.toolbar)

        self.ma_row = QWidget(HistoryDialog)
        self.ma_row.setObjectName(u"ma_row")
        self.ma_layout = QHBoxLayout(self.ma_row)
        self.ma_layout.setSpacing(2)
        self.ma_layout.setObjectName(u"ma_layout")
        self.ma_layout.setContentsMargins(0, 0, 0, 0)
        self.cb_ma5 = FilledCheckBox(self.ma_row)
        self.cb_ma5.setObjectName(u"cb_ma5")

        self.ma_layout.addWidget(self.cb_ma5)

        self.cb_ma10 = FilledCheckBox(self.ma_row)
        self.cb_ma10.setObjectName(u"cb_ma10")

        self.ma_layout.addWidget(self.cb_ma10)

        self.cb_ma20 = FilledCheckBox(self.ma_row)
        self.cb_ma20.setObjectName(u"cb_ma20")

        self.ma_layout.addWidget(self.cb_ma20)

        self.cb_ma30 = FilledCheckBox(self.ma_row)
        self.cb_ma30.setObjectName(u"cb_ma30")

        self.ma_layout.addWidget(self.cb_ma30)

        self.cb_ma60 = FilledCheckBox(self.ma_row)
        self.cb_ma60.setObjectName(u"cb_ma60")

        self.ma_layout.addWidget(self.cb_ma60)

        self.ma_space = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.ma_layout.addItem(self.ma_space)


        self.chart_layout.addWidget(self.ma_row)

        self.content_layout = QHBoxLayout()
        self.content_layout.setSpacing(6)
        self.content_layout.setObjectName(u"content_layout")
        self.chart = HistoryChart(HistoryDialog)
        self.chart.setObjectName(u"chart")
        self.chart.setMinimumSize(QSize(200, 100))

        self.content_layout.addWidget(self.chart)

        self.order_book = OrderBookWidget(HistoryDialog)
        self.order_book.setObjectName(u"order_book")

        self.content_layout.addWidget(self.order_book)

        self.content_layout.setStretch(0, 1)

        self.chart_layout.addLayout(self.content_layout)

        self.chart_readout = QLabel(HistoryDialog)
        self.chart_readout.setObjectName(u"chart_readout")
        self.chart_readout.setWordWrap(False)
        self.chart_readout.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        sizePolicy1.setHeightForWidth(self.chart_readout.sizePolicy().hasHeightForWidth())
        self.chart_readout.setSizePolicy(sizePolicy1)

        self.chart_layout.addWidget(self.chart_readout)

        self.status = QLabel(HistoryDialog)
        self.status.setObjectName(u"status")
        sizePolicy2 = QSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        sizePolicy2.setHorizontalStretch(0)
        sizePolicy2.setVerticalStretch(0)
        sizePolicy2.setHeightForWidth(self.status.sizePolicy().hasHeightForWidth())
        self.status.setSizePolicy(sizePolicy2)
        self.status.setProperty(u"settingDescription", True)
        self.status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.chart_layout.addWidget(self.status)

        self.chart_layout.setStretch(4, 1)

        self.retranslateUi(HistoryDialog)

        QMetaObject.connectSlotsByName(HistoryDialog)
    # setupUi

    def retranslateUi(self, HistoryDialog):
        self.pop_out_button.setText(QCoreApplication.translate("HistoryDialog", u"\u2197", None))
#if QT_CONFIG(accessibility)
        self.pop_out_button.setAccessibleName(QCoreApplication.translate("HistoryDialog", u"\u6253\u5f00\u72ec\u7acb\u7a97\u53e3", None))
#endif // QT_CONFIG(accessibility)
        self.close_button.setText(QCoreApplication.translate("HistoryDialog", u"\u00d7", None))
#if QT_CONFIG(accessibility)
        self.close_button.setAccessibleName(QCoreApplication.translate("HistoryDialog", u"\u5173\u95ed\u56fe\u8868", None))
#endif // QT_CONFIG(accessibility)
        self.quote_code.setText("")
        self.quote_current.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_change.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_change_pct.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.btn_previous.setText("")
#if QT_CONFIG(accessibility)
        self.btn_previous.setAccessibleName(QCoreApplication.translate("HistoryDialog", u"\u4e0a\u4e00\u53ea\u81ea\u9009", None))
#endif // QT_CONFIG(accessibility)
        self.btn_next.setText("")
#if QT_CONFIG(accessibility)
        self.btn_next.setAccessibleName(QCoreApplication.translate("HistoryDialog", u"\u4e0b\u4e00\u53ea\u81ea\u9009", None))
#endif // QT_CONFIG(accessibility)
        self.quote_high_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u9ad8", None))
        self.quote_high.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_low_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u4f4e", None))
        self.quote_low.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_open_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5f00", None))
        self.quote_open.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_prev_close_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6628\u6536", None))
        self.quote_prev_close.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_inside_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5185\u76d8", None))
        self.quote_inside.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_outside_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5916\u76d8", None))
        self.quote_outside.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_volume_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6210\u4ea4\u91cf", None))
        self.quote_volume.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_amount_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6210\u4ea4\u989d", None))
        self.quote_amount.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_average_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5747\u4ef7", None))
        self.quote_average.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_turnover_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6362\u624b\u7387", None))
        self.quote_turnover.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_amplitude_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u632f\u5e45", None))
        self.quote_amplitude.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_volume_ratio_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u91cf\u6bd4", None))
        self.quote_volume_ratio.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_pe_ttm_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5e02\u76c8\uff08TTM\uff09", None))
        self.quote_pe_ttm.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_pe_static_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5e02\u76c8\uff08\u9759\uff09", None))
        self.quote_pe_static.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_pb_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5e02\u51c0\u7387", None))
        self.quote_pb.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_pe_dynamic_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5e02\u76c8\uff08\u52a8\uff09", None))
        self.quote_pe_dynamic.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_market_cap_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u603b\u5e02\u503c", None))
        self.quote_market_cap.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_total_shares_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u603b\u80a1\u672c", None))
        self.quote_total_shares.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_float_market_cap_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6d41\u901a\u5e02\u503c", None))
        self.quote_float_market_cap.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_float_shares_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6d41\u901a\u80a1", None))
        self.quote_float_shares.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.btn_intraday.setText(QCoreApplication.translate("HistoryDialog", u"\u5206\u65f6", None))
        self.btn_five_day.setText(QCoreApplication.translate("HistoryDialog", u"5\u65e5", None))
        self.btn_daily.setText(QCoreApplication.translate("HistoryDialog", u"\u65e5K", None))
        self.cb_volume.setText(QCoreApplication.translate("HistoryDialog", u"\u6210\u4ea4\u91cf", None))
        self.cb_average.setText(QCoreApplication.translate("HistoryDialog", u"\u5747\u7ebf", None))
        self.cb_order_book.setText(QCoreApplication.translate("HistoryDialog", u"\u4e94\u6863", None))
        self.cb_ma5.setText(QCoreApplication.translate("HistoryDialog", u"MA5", None))
        self.cb_ma10.setText(QCoreApplication.translate("HistoryDialog", u"MA10", None))
        self.cb_ma20.setText(QCoreApplication.translate("HistoryDialog", u"MA20", None))
        self.cb_ma30.setText(QCoreApplication.translate("HistoryDialog", u"MA30", None))
        self.cb_ma60.setText(QCoreApplication.translate("HistoryDialog", u"MA60", None))
        self.chart_readout.setText(QCoreApplication.translate("HistoryDialog", u"\u5747 \u2014  \u4ef7\u683c \u2014 (\u2014)  \u91cf \u2014", None))
        pass
    # retranslateUi
