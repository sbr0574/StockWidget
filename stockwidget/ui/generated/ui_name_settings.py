# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'name_settings.ui'
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
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QFrame,
    QHBoxLayout, QLabel, QSizePolicy, QSpacerItem,
    QVBoxLayout, QWidget)

class Ui_NameSettingsPanel(object):
    def setupUi(self, name_settings_panel):
        if not name_settings_panel.objectName():
            name_settings_panel.setObjectName(u"name_settings_panel")
        name_settings_panel.resize(220, 80)
        self.nameSettingsLayout = QVBoxLayout(name_settings_panel)
        self.nameSettingsLayout.setSpacing(7)
        self.nameSettingsLayout.setObjectName(u"nameSettingsLayout")
        self.nameSettingsLayout.setContentsMargins(10, 10, 10, 10)
        self.nameLengthLayout = QHBoxLayout()
        self.nameLengthLayout.setSpacing(6)
        self.nameLengthLayout.setObjectName(u"nameLengthLayout")
        self.nameLengthLayout.setContentsMargins(0, 0, 0, 0)
        self.length_label = QLabel(name_settings_panel)
        self.length_label.setObjectName(u"length_label")

        self.nameLengthLayout.addWidget(self.length_label)

        self.name_settings_namelen = QComboBox(name_settings_panel)
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.addItem("")
        self.name_settings_namelen.setObjectName(u"name_settings_namelen")

        self.nameLengthLayout.addWidget(self.name_settings_namelen)


        self.nameSettingsLayout.addLayout(self.nameLengthLayout)

        self.nameFlagsLayout = QHBoxLayout()
        self.nameFlagsLayout.setSpacing(10)
        self.nameFlagsLayout.setObjectName(u"nameFlagsLayout")
        self.nameFlagsLayout.setContentsMargins(0, 0, 0, 0)
        self.name_settings_code = QCheckBox(name_settings_panel)
        self.name_settings_code.setObjectName(u"name_settings_code")

        self.nameFlagsLayout.addWidget(self.name_settings_code)

        self.name_settings_type = QCheckBox(name_settings_panel)
        self.name_settings_type.setObjectName(u"name_settings_type")

        self.nameFlagsLayout.addWidget(self.name_settings_type)

        self.nameFlagsSpacer = QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.nameFlagsLayout.addItem(self.nameFlagsSpacer)


        self.nameSettingsLayout.addLayout(self.nameFlagsLayout)


        self.retranslateUi(name_settings_panel)

        QMetaObject.connectSlotsByName(name_settings_panel)
    # setupUi

    def retranslateUi(self, name_settings_panel):
        self.length_label.setText(QCoreApplication.translate("NameSettingsPanel", u"\u663e\u793a\u5b57\u6570\uff1a", None))
        self.name_settings_namelen.setItemText(0, QCoreApplication.translate("NameSettingsPanel", u"\u4e0d\u663e\u793a", None))
        self.name_settings_namelen.setItemText(1, QCoreApplication.translate("NameSettingsPanel", u"\u5168\u90e8\u663e\u793a", None))
        self.name_settings_namelen.setItemText(2, QCoreApplication.translate("NameSettingsPanel", u"1\u4e2a\u5b57", None))
        self.name_settings_namelen.setItemText(3, QCoreApplication.translate("NameSettingsPanel", u"2\u4e2a\u5b57", None))
        self.name_settings_namelen.setItemText(4, QCoreApplication.translate("NameSettingsPanel", u"3\u4e2a\u5b57", None))
        self.name_settings_namelen.setItemText(5, QCoreApplication.translate("NameSettingsPanel", u"4\u4e2a\u5b57", None))

        self.name_settings_code.setText(QCoreApplication.translate("NameSettingsPanel", u"\u663e\u793a\u4ee3\u7801", None))
        self.name_settings_type.setText(QCoreApplication.translate("NameSettingsPanel", u"\u663e\u793a\u7c7b\u578b", None))
        pass
    # retranslateUi

