# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'add_code_panel.ui'
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
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFrame, QHBoxLayout,
    QLabel, QPushButton, QSizePolicy, QSpacerItem,
    QVBoxLayout, QWidget)

from stockwidget.ui.add_code_panel import (CodeSearchInput, FilledCheckBox, FilterCheckRow, SearchResultList,
    SelectAllCheckBox)

class Ui_AddCodePanel(object):
    def setupUi(self, add_code_panel):
        if not add_code_panel.objectName():
            add_code_panel.setObjectName(u"add_code_panel")
        add_code_panel.setMinimumSize(QSize(430, 0))
        add_code_panel.setMaximumSize(QSize(430, 16777215))
        add_code_panel.resize(430, 380)
        self.addCodeLayout = QVBoxLayout(add_code_panel)
        self.addCodeLayout.setSpacing(7)
        self.addCodeLayout.setObjectName(u"addCodeLayout")
        self.addCodeLayout.setContentsMargins(10, 10, 10, 10)
        self.add_code_search_input = CodeSearchInput(add_code_panel)
        self.add_code_search_input.setObjectName(u"add_code_search_input")

        self.addCodeLayout.addWidget(self.add_code_search_input)

        self.category_filters = FilterCheckRow(add_code_panel)
        self.category_filters.setObjectName(u"category_filters")
        self.categoryFilterLayout = QHBoxLayout(self.category_filters)
        self.categoryFilterLayout.setSpacing(8)
        self.categoryFilterLayout.setObjectName(u"categoryFilterLayout")
        self.categoryFilterLayout.setContentsMargins(0, 0, 0, 0)
        self.category_filter_title = QLabel(self.category_filters)
        self.category_filter_title.setObjectName(u"category_filter_title")

        self.categoryFilterLayout.addWidget(self.category_filter_title)

        self.category_filter_all = SelectAllCheckBox(self.category_filters)
        self.category_filter_all.setObjectName(u"category_filter_all")
        self.category_filter_all.setTristate(True)
        self.category_filter_all.setChecked(True)

        self.categoryFilterLayout.addWidget(self.category_filter_all)

        self.category_filter_stock = FilledCheckBox(self.category_filters)
        self.category_filter_stock.setObjectName(u"category_filter_stock")
        self.category_filter_stock.setChecked(True)

        self.categoryFilterLayout.addWidget(self.category_filter_stock)

        self.category_filter_fund = FilledCheckBox(self.category_filters)
        self.category_filter_fund.setObjectName(u"category_filter_fund")
        self.category_filter_fund.setChecked(True)

        self.categoryFilterLayout.addWidget(self.category_filter_fund)

        self.category_filter_index = FilledCheckBox(self.category_filters)
        self.category_filter_index.setObjectName(u"category_filter_index")
        self.category_filter_index.setChecked(True)

        self.categoryFilterLayout.addWidget(self.category_filter_index)

        self.category_filter_futures = FilledCheckBox(self.category_filters)
        self.category_filter_futures.setObjectName(u"category_filter_futures")
        self.category_filter_futures.setChecked(True)

        self.categoryFilterLayout.addWidget(self.category_filter_futures)

        self.categoryFilterSpacer = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.categoryFilterLayout.addItem(self.categoryFilterSpacer)


        self.addCodeLayout.addWidget(self.category_filters)

        self.region_filters = FilterCheckRow(add_code_panel)
        self.region_filters.setObjectName(u"region_filters")
        self.regionFilterLayout = QHBoxLayout(self.region_filters)
        self.regionFilterLayout.setSpacing(8)
        self.regionFilterLayout.setObjectName(u"regionFilterLayout")
        self.regionFilterLayout.setContentsMargins(0, 0, 0, 0)
        self.region_filter_title = QLabel(self.region_filters)
        self.region_filter_title.setObjectName(u"region_filter_title")

        self.regionFilterLayout.addWidget(self.region_filter_title)

        self.region_filter_all = SelectAllCheckBox(self.region_filters)
        self.region_filter_all.setObjectName(u"region_filter_all")
        self.region_filter_all.setTristate(True)
        self.region_filter_all.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_all)

        self.region_filter_sh = FilledCheckBox(self.region_filters)
        self.region_filter_sh.setObjectName(u"region_filter_sh")
        self.region_filter_sh.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_sh)

        self.region_filter_sz = FilledCheckBox(self.region_filters)
        self.region_filter_sz.setObjectName(u"region_filter_sz")
        self.region_filter_sz.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_sz)

        self.region_filter_bj = FilledCheckBox(self.region_filters)
        self.region_filter_bj.setObjectName(u"region_filter_bj")
        self.region_filter_bj.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_bj)

        self.region_filter_hk = FilledCheckBox(self.region_filters)
        self.region_filter_hk.setObjectName(u"region_filter_hk")
        self.region_filter_hk.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_hk)

        self.region_filter_us = FilledCheckBox(self.region_filters)
        self.region_filter_us.setObjectName(u"region_filter_us")
        self.region_filter_us.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_us)

        self.region_filter_other = FilledCheckBox(self.region_filters)
        self.region_filter_other.setObjectName(u"region_filter_other")
        self.region_filter_other.setChecked(True)

        self.regionFilterLayout.addWidget(self.region_filter_other)

        self.regionFilterSpacer = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.regionFilterLayout.addItem(self.regionFilterSpacer)


        self.addCodeLayout.addWidget(self.region_filters)

        self.add_code_results = SearchResultList(add_code_panel)
        self.add_code_results.setObjectName(u"add_code_results")
        self.add_code_results.setUniformItemSizes(True)
        self.add_code_results.setSpacing(0)
        self.add_code_results.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.add_code_results.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.add_code_results.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.add_code_results.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.add_code_results.setMinimumSize(QSize(0, 242))
        self.add_code_results.setMaximumSize(QSize(16777215, 242))

        self.addCodeLayout.addWidget(self.add_code_results)

        self.addCodePagesLayout = QHBoxLayout()
        self.addCodePagesLayout.setObjectName(u"addCodePagesLayout")
        self.addCodePagesLayout.setContentsMargins(0, 0, 0, 0)
        self.add_code_previous_page = QPushButton(add_code_panel)
        self.add_code_previous_page.setObjectName(u"add_code_previous_page")
        self.add_code_previous_page.setAutoDefault(False)

        self.addCodePagesLayout.addWidget(self.add_code_previous_page)

        self.add_code_page_label = QLabel(add_code_panel)
        self.add_code_page_label.setObjectName(u"add_code_page_label")
        self.add_code_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.addCodePagesLayout.addWidget(self.add_code_page_label)

        self.add_code_next_page = QPushButton(add_code_panel)
        self.add_code_next_page.setObjectName(u"add_code_next_page")
        self.add_code_next_page.setAutoDefault(False)

        self.addCodePagesLayout.addWidget(self.add_code_next_page)

        self.addCodePagesLayout.setStretch(1, 1)

        self.addCodeLayout.addLayout(self.addCodePagesLayout)


        self.retranslateUi(add_code_panel)

        QMetaObject.connectSlotsByName(add_code_panel)
    # setupUi

    def retranslateUi(self, add_code_panel):
        self.add_code_search_input.setPlaceholderText(QCoreApplication.translate("AddCodePanel", u"\u641c\u7d22\u4ee3\u7801\u3001\u540d\u79f0\u3001\u62fc\u97f3\u6216\u7f29\u5199\uff0c\u7a7a\u683c\u533a\u5206\u5173\u952e\u8bcd", None))
        self.category_filter_title.setText(QCoreApplication.translate("AddCodePanel", u"\u7c7b\u522b", None))
        self.category_filter_all.setText(QCoreApplication.translate("AddCodePanel", u"\u5168\u9009", None))
        self.category_filter_stock.setText(QCoreApplication.translate("AddCodePanel", u"\u80a1\u7968", None))
        self.category_filter_fund.setText(QCoreApplication.translate("AddCodePanel", u"\u57fa\u91d1", None))
        self.category_filter_index.setText(QCoreApplication.translate("AddCodePanel", u"\u6307\u6570", None))
        self.category_filter_futures.setText(QCoreApplication.translate("AddCodePanel", u"\u671f\u8d27", None))
        self.region_filter_title.setText(QCoreApplication.translate("AddCodePanel", u"\u5730\u533a", None))
        self.region_filter_all.setText(QCoreApplication.translate("AddCodePanel", u"\u5168\u9009", None))
        self.region_filter_sh.setText(QCoreApplication.translate("AddCodePanel", u"\u6caa", None))
        self.region_filter_sz.setText(QCoreApplication.translate("AddCodePanel", u"\u6df1", None))
        self.region_filter_bj.setText(QCoreApplication.translate("AddCodePanel", u"\u4eac", None))
        self.region_filter_hk.setText(QCoreApplication.translate("AddCodePanel", u"\u6e2f", None))
        self.region_filter_us.setText(QCoreApplication.translate("AddCodePanel", u"\u7f8e", None))
        self.region_filter_other.setText(QCoreApplication.translate("AddCodePanel", u"\u5176\u4ed6", None))
        self.add_code_previous_page.setText(QCoreApplication.translate("AddCodePanel", u"\u4e0a\u4e00\u9875", None))
        self.add_code_page_label.setText(QCoreApplication.translate("AddCodePanel", u"0 / 0\uff08\u5171 0 \u6761\uff09", None))
        self.add_code_next_page.setText(QCoreApplication.translate("AddCodePanel", u"\u4e0b\u4e00\u9875", None))
        pass
    # retranslateUi

