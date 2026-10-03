# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'metric_pool.ui'
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
from PySide6.QtWidgets import (QApplication, QLabel, QListView, QListWidgetItem,
    QSizePolicy, QVBoxLayout, QWidget)

from stockwidget.ui.controls.metrics import (MetricCatalogWidget, MetricListWidget)

class Ui_MetricPool(object):
    def setupUi(self, metric_pool):
        if not metric_pool.objectName():
            metric_pool.setObjectName(u"metric_pool")
        metric_pool.resize(240, 220)
        self.metricPoolLayout = QVBoxLayout(metric_pool)
        self.metricPoolLayout.setSpacing(5)
        self.metricPoolLayout.setObjectName(u"metricPoolLayout")
        self.metricPoolLayout.setContentsMargins(0, 0, 0, 0)
        self.available_label = QLabel(metric_pool)
        self.available_label.setObjectName(u"available_label")
        self.available_label.setMinimumSize(QSize(0, 26))
        self.available_label.setMaximumSize(QSize(16777215, 26))
        font = QFont()
        font.setBold(True)
        self.available_label.setFont(font)
        self.available_label.setProperty(u"settingTitle", True)

        self.metricPoolLayout.addWidget(self.available_label)

        self.available_description = QLabel(metric_pool)
        self.available_description.setObjectName(u"available_description")
        self.available_description.setWordWrap(True)
        self.available_description.setProperty(u"settingDescription", True)
        self.available_description.setMinimumSize(QSize(310, 0))
        self.available_description.setMaximumSize(QSize(310, 16777215))
        self.available_description.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignTop)

        self.metricPoolLayout.addWidget(self.available_description)

        self.metric_available_pool = MetricCatalogWidget(metric_pool)
        self.metric_available_pool.setObjectName(u"metric_available_pool")
        self.metric_available_pool.setFlow(QListView.Flow.LeftToRight)
        self.metric_available_pool.setWrapping(True)
        self.metric_available_pool.setSpacing(2)
        self.metric_available_pool.setMinimumSize(QSize(0, 56))
        self.metric_available_pool.setMaximumSize(QSize(16777215, 56))

        self.metricPoolLayout.addWidget(self.metric_available_pool)

        self.displayed_label = QLabel(metric_pool)
        self.displayed_label.setObjectName(u"displayed_label")
        self.displayed_label.setMinimumSize(QSize(0, 26))
        self.displayed_label.setMaximumSize(QSize(16777215, 26))
        self.displayed_label.setFont(font)
        self.displayed_label.setProperty(u"settingTitle", True)

        self.metricPoolLayout.addWidget(self.displayed_label)

        self.metric_displayed_pool = MetricListWidget(metric_pool)
        self.metric_displayed_pool.setObjectName(u"metric_displayed_pool")
        self.metric_displayed_pool.setMinimumSize(QSize(0, 56))
        self.metric_displayed_pool.setMaximumSize(QSize(16777215, 56))
        self.metric_displayed_pool.setFlow(QListView.Flow.LeftToRight)
        self.metric_displayed_pool.setWrapping(True)
        self.metric_displayed_pool.setSpacing(2)

        self.metricPoolLayout.addWidget(self.metric_displayed_pool)


        self.retranslateUi(metric_pool)

        QMetaObject.connectSlotsByName(metric_pool)
    # setupUi

    def retranslateUi(self, metric_pool):
        self.available_label.setText(QCoreApplication.translate("MetricPool", u"\u53ef\u7528\u6307\u6807", None))
        self.available_description.setText(QCoreApplication.translate("MetricPool", u"\u62d6\u52a8\u6216\u53cc\u51fb\u53ef\u7528\u6307\u6807\u6dfb\u52a0\uff1b\u62d6\u52a8\u5df2\u663e\u793a\u6307\u6807\u6392\u5e8f\uff0c\u53cc\u51fb\u6216\u62d6\u56de\u76ee\u5f55\u79fb\u9664\u3002", None))
        self.metric_available_pool.setProperty(u"emptyText", QCoreApplication.translate("MetricPool", u"\u5df2\u5168\u90e8\u663e\u793a", None))
        self.displayed_label.setText(QCoreApplication.translate("MetricPool", u"\u5df2\u663e\u793a\u6307\u6807", None))
        self.metric_displayed_pool.setProperty(u"emptyText", QCoreApplication.translate("MetricPool", u"\u62d6\u5165\u8981\u663e\u793a\u7684\u6307\u6807", None))
        pass
    # retranslateUi

