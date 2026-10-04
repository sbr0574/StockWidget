# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'floating_widget.ui'
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
    QHeaderView, QLabel, QSizePolicy, QVBoxLayout,
    QWidget)

from stockwidget.ui.controls.quote_view import (PagerWidget, QuoteTableView)

class Ui_FloatingWidget(object):
    def setupUi(self, floating_widget):
        if not floating_widget.objectName():
            floating_widget.setObjectName(u"floating_widget")
        floating_widget.resize(240, 90)
        self.panel = QWidget(floating_widget)
        self.panel.setObjectName(u"panel")
        self.panel.setGeometry(QRect(0, 0, 240, 90))
        self.vbox = QVBoxLayout(self.panel)
        self.vbox.setSpacing(0)
        self.vbox.setObjectName(u"vbox")
        self.vbox.setContentsMargins(10, 6, 10, 6)
        self.hide_notice = QLabel(self.panel)
        self.hide_notice.setObjectName(u"hide_notice")

        self.vbox.addWidget(self.hide_notice)

        self.message_label = QLabel(self.panel)
        self.message_label.setObjectName(u"message_label")

        self.vbox.addWidget(self.message_label)

        self.data_layout = QHBoxLayout()
        self.data_layout.setSpacing(3)
        self.data_layout.setObjectName(u"data_layout")
        self.data_layout.setContentsMargins(0, 0, 0, 0)
        self.pager = PagerWidget(self.panel)
        self.pager.setObjectName(u"pager")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pager.sizePolicy().hasHeightForWidth())
        self.pager.setSizePolicy(sizePolicy)

        self.data_layout.addWidget(self.pager)

        self.table = QuoteTableView(self.panel)
        self.table.setObjectName(u"table")
        self.table.setShowGrid(False)
        self.table.setFrameShape(QFrame.Shape.NoFrame)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setTextElideMode(Qt.TextElideMode.ElideNone)

        self.data_layout.addWidget(self.table)

        self.split_separator = QFrame(self.panel)
        self.split_separator.setObjectName(u"split_separator")
        self.split_separator.setFrameShape(QFrame.Shape.VLine)
        self.split_separator.setFrameShadow(QFrame.Shadow.Plain)
        self.split_separator.setLineWidth(1)

        self.data_layout.addWidget(self.split_separator)

        self.right_table = QuoteTableView(self.panel)
        self.right_table.setObjectName(u"right_table")
        self.right_table.setShowGrid(False)
        self.right_table.setFrameShape(QFrame.Shape.NoFrame)
        self.right_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.right_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.right_table.setTextElideMode(Qt.TextElideMode.ElideNone)

        self.data_layout.addWidget(self.right_table)


        self.vbox.addLayout(self.data_layout)


        self.retranslateUi(floating_widget)

        QMetaObject.connectSlotsByName(floating_widget)
    # setupUi

    def retranslateUi(self, floating_widget):
        self.hide_notice.setText("")
        self.hide_notice.setStyleSheet(QCoreApplication.translate("FloatingWidget", u"padding: 2px 4px;", None))
        self.message_label.setText(QCoreApplication.translate("FloatingWidget", u"\u52a0\u8f7d\u4e2d\u2026", None))
        self.message_label.setStyleSheet(QCoreApplication.translate("FloatingWidget", u"padding: 2px 4px;", None))
        pass
    # retranslateUi

