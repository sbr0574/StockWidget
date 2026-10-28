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
        self.quote_header_layout.setSpacing(20)
        self.quote_header_layout.setObjectName(u"quote_header_layout")
        self.quote_header_layout.setContentsMargins(0, 0, 0, 6)
        self.quote_primary = QWidget(self.quote_header)
        self.quote_primary.setObjectName(u"quote_primary")
        self.quote_primary.setMinimumSize(QSize(200, 0))
        self.quote_primary_layout = QVBoxLayout(self.quote_primary)
        self.quote_primary_layout.setSpacing(6)
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

        self.quote_session = QWidget(self.quote_header)
        self.quote_session.setObjectName(u"quote_session")
        self.quote_session_layout = QGridLayout(self.quote_session)
        self.quote_session_layout.setObjectName(u"quote_session_layout")
        self.quote_session_layout.setHorizontalSpacing(12)
        self.quote_session_layout.setVerticalSpacing(4)
        self.quote_session_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_high_caption = QLabel(self.quote_session)
        self.quote_high_caption.setObjectName(u"quote_high_caption")
        self.quote_high_caption.setProperty(u"quoteCaption", True)
        self.quote_high_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_session_layout.addWidget(self.quote_high_caption, 0, 0, 1, 1)

        self.quote_high = QLabel(self.quote_session)
        self.quote_high.setObjectName(u"quote_high")
        self.quote_high.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_high.setFont(font2)

        self.quote_session_layout.addWidget(self.quote_high, 0, 1, 1, 1)

        self.quote_low_caption = QLabel(self.quote_session)
        self.quote_low_caption.setObjectName(u"quote_low_caption")
        self.quote_low_caption.setProperty(u"quoteCaption", True)
        self.quote_low_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_session_layout.addWidget(self.quote_low_caption, 1, 0, 1, 1)

        self.quote_low = QLabel(self.quote_session)
        self.quote_low.setObjectName(u"quote_low")
        self.quote_low.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_low.setFont(font2)

        self.quote_session_layout.addWidget(self.quote_low, 1, 1, 1, 1)

        self.quote_open_caption = QLabel(self.quote_session)
        self.quote_open_caption.setObjectName(u"quote_open_caption")
        self.quote_open_caption.setProperty(u"quoteCaption", True)
        self.quote_open_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_session_layout.addWidget(self.quote_open_caption, 2, 0, 1, 1)

        self.quote_open = QLabel(self.quote_session)
        self.quote_open.setObjectName(u"quote_open")
        self.quote_open.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_open.setFont(font2)

        self.quote_session_layout.addWidget(self.quote_open, 2, 1, 1, 1)


        self.quote_header_layout.addWidget(self.quote_session)

        self.quote_header_space = QSpacerItem(8, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.quote_header_layout.addItem(self.quote_header_space)

        self.quote_activity = QWidget(self.quote_header)
        self.quote_activity.setObjectName(u"quote_activity")
        self.quote_activity_layout = QGridLayout(self.quote_activity)
        self.quote_activity_layout.setObjectName(u"quote_activity_layout")
        self.quote_activity_layout.setHorizontalSpacing(12)
        self.quote_activity_layout.setVerticalSpacing(4)
        self.quote_activity_layout.setContentsMargins(0, 0, 0, 0)
        self.quote_amount_caption = QLabel(self.quote_activity)
        self.quote_amount_caption.setObjectName(u"quote_amount_caption")
        self.quote_amount_caption.setProperty(u"quoteCaption", True)
        self.quote_amount_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_activity_layout.addWidget(self.quote_amount_caption, 0, 0, 1, 1)

        self.quote_amount = QLabel(self.quote_activity)
        self.quote_amount.setObjectName(u"quote_amount")
        self.quote_amount.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_amount.setFont(font2)

        self.quote_activity_layout.addWidget(self.quote_amount, 0, 1, 1, 1)

        self.quote_volume_caption = QLabel(self.quote_activity)
        self.quote_volume_caption.setObjectName(u"quote_volume_caption")
        self.quote_volume_caption.setProperty(u"quoteCaption", True)
        self.quote_volume_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_activity_layout.addWidget(self.quote_volume_caption, 1, 0, 1, 1)

        self.quote_volume = QLabel(self.quote_activity)
        self.quote_volume.setObjectName(u"quote_volume")
        self.quote_volume.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_volume.setFont(font2)

        self.quote_activity_layout.addWidget(self.quote_volume, 1, 1, 1, 1)

        self.quote_average_caption = QLabel(self.quote_activity)
        self.quote_average_caption.setObjectName(u"quote_average_caption")
        self.quote_average_caption.setProperty(u"quoteCaption", True)
        self.quote_average_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_activity_layout.addWidget(self.quote_average_caption, 2, 0, 1, 1)

        self.quote_average = QLabel(self.quote_activity)
        self.quote_average.setObjectName(u"quote_average")
        self.quote_average.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_average.setFont(font2)

        self.quote_activity_layout.addWidget(self.quote_average, 2, 1, 1, 1)

        self.quote_volume_ratio_caption = QLabel(self.quote_activity)
        self.quote_volume_ratio_caption.setObjectName(u"quote_volume_ratio_caption")
        self.quote_volume_ratio_caption.setProperty(u"quoteCaption", True)
        self.quote_volume_ratio_caption.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)

        self.quote_activity_layout.addWidget(self.quote_volume_ratio_caption, 3, 0, 1, 1)

        self.quote_volume_ratio = QLabel(self.quote_activity)
        self.quote_volume_ratio.setObjectName(u"quote_volume_ratio")
        self.quote_volume_ratio.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.quote_volume_ratio.setFont(font2)

        self.quote_activity_layout.addWidget(self.quote_volume_ratio, 3, 1, 1, 1)


        self.quote_header_layout.addWidget(self.quote_activity)

        self.quote_header_layout.setStretch(2, 1)

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

        self.status = QLabel(HistoryDialog)
        self.status.setObjectName(u"status")
        sizePolicy1 = QSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.status.sizePolicy().hasHeightForWidth())
        self.status.setSizePolicy(sizePolicy1)
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
        self.quote_current.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_change.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_change_pct.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_high_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u9ad8", None))
        self.quote_high.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_low_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u4f4e", None))
        self.quote_low.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_open_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5f00", None))
        self.quote_open.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_amount_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6210\u4ea4\u989d", None))
        self.quote_amount.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_volume_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u6210\u4ea4\u91cf", None))
        self.quote_volume.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_average_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u5747\u4ef7", None))
        self.quote_average.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
        self.quote_volume_ratio_caption.setText(QCoreApplication.translate("HistoryDialog", u"\u91cf\u6bd4", None))
        self.quote_volume_ratio.setText(QCoreApplication.translate("HistoryDialog", u"\u2014", None))
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
        pass
    # retranslateUi
