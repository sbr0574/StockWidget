# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'unit_settings.ui'
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
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel,
    QRadioButton, QSizePolicy, QSpacerItem, QVBoxLayout,
    QWidget)

class Ui_UnitSettingsPanel(object):
    def setupUi(self, unit_settings_panel):
        if not unit_settings_panel.objectName():
            unit_settings_panel.setObjectName(u"unit_settings_panel")
        unit_settings_panel.resize(210, 80)
        self.unitSettingsLayout = QVBoxLayout(unit_settings_panel)
        self.unitSettingsLayout.setSpacing(7)
        self.unitSettingsLayout.setObjectName(u"unitSettingsLayout")
        self.unitSettingsLayout.setContentsMargins(10, 10, 10, 10)
        self.title_label = QLabel(unit_settings_panel)
        self.title_label.setObjectName(u"title_label")

        self.unitSettingsLayout.addWidget(self.title_label)

        self.unitOptionsLayout = QHBoxLayout()
        self.unitOptionsLayout.setSpacing(10)
        self.unitOptionsLayout.setObjectName(u"unitOptionsLayout")
        self.unitOptionsLayout.setContentsMargins(0, 0, 0, 0)
        self.unit_option_cn = QRadioButton(unit_settings_panel)
        self.unit_option_cn.setObjectName(u"unit_option_cn")

        self.unitOptionsLayout.addWidget(self.unit_option_cn)

        self.unit_option_en = QRadioButton(unit_settings_panel)
        self.unit_option_en.setObjectName(u"unit_option_en")

        self.unitOptionsLayout.addWidget(self.unit_option_en)

        self.unit_option_auto = QRadioButton(unit_settings_panel)
        self.unit_option_auto.setObjectName(u"unit_option_auto")

        self.unitOptionsLayout.addWidget(self.unit_option_auto)

        self.unitOptionsSpacer = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.unitOptionsLayout.addItem(self.unitOptionsSpacer)


        self.unitSettingsLayout.addLayout(self.unitOptionsLayout)


        self.retranslateUi(unit_settings_panel)

        QMetaObject.connectSlotsByName(unit_settings_panel)
    # setupUi

    def retranslateUi(self, unit_settings_panel):
        self.title_label.setText(QCoreApplication.translate("UnitSettingsPanel", u"\u6570\u503c\u5355\u4f4d\uff1a", None))
        self.unit_option_cn.setText(QCoreApplication.translate("UnitSettingsPanel", u"\u4e2d\u6587", None))
        self.unit_option_en.setText(QCoreApplication.translate("UnitSettingsPanel", u"\u82f1\u6587", None))
        self.unit_option_auto.setText(QCoreApplication.translate("UnitSettingsPanel", u"\u81ea\u52a8", None))
#if QT_CONFIG(tooltip)
        self.unit_option_auto.setToolTip(QCoreApplication.translate("UnitSettingsPanel", u"\u7f8e\u80a1\u3001\u56fd\u9645\u6307\u6570\u4f7f\u7528\u82f1\u6587\u5355\u4f4d\uff0c\u56fd\u5185\u3001\u6e2f\u80a1\u7b49\u4f7f\u7528\u4e2d\u6587\u5355\u4f4d", None))
#endif // QT_CONFIG(tooltip)
        pass
    # retranslateUi

