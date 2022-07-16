from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from db_window import Ui_DBWindow
from datetime import date
from back_end import backend


class Ui_MainWindow(object):

    ############################################################################################################
    def openDBWindow(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_DBWindow()
        self.ui.setupUi(self.window)
        self.window.show()

    def browse_xml(self):
        browse = QFileDialog.getOpenFileName(caption="Open file", filter="XML files (*.xml)")
        self.xml_path.setText(browse[0])

    def browse_xlsx(self):
        browse = QFileDialog.getExistingDirectory(None, 'Select Folder')
        self.xlsx_path.setText(browse)

    def undisable_button(self):
        if len(self.xml_path.text())>0 and len(self.xlsx_path.text())>0:
            self.export_button.setDisabled(False)


    def export(self):
            xml_path = self.xml_path.text()
            xlsx_path = self.xlsx_path.text()
            if self.xlsx_name.text() != '':
                xlsx_path = f'{xlsx_path}/{self.xlsx_name.text()}.xlsx'
            else:
                xlsx_path = f'{xlsx_path}/Invoices_{date.today()}.xlsx'

            backend(xlsx_path, xml_path)

            msg_box = QMessageBox()
            msg_box.setWindowTitle("Բարեհաջող ավարտ")
            msg_box.setText("Դուք բարեհաջող արտահանեցիք excel ֆայլը։")
            button = msg_box.exec()

    ############################################################################################################
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(600, 340)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(600, 340))
        MainWindow.setMaximumSize(QtCore.QSize(600, 340))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(
            "C:\\Users\\hayk.asatryan\\Desktop\\Haik\\pythonprojects\\eInvoiceToExcel\\ui\\../Untergunter-Leaf-Mimes-Text-xml.ico"),
                       QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("QMainWindow{\n"
                                 "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:0 rgba(166, 206, 57, 255), stop:1 rgba(255, 255, 255, 255));\n"
                                 "}\n"
                                 "font: 8pt \"Helvetica\";")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 574, 271))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.xml_path = QtWidgets.QLineEdit(self.verticalLayoutWidget)

        self.xml_path.textChanged.connect(lambda: self.undisable_button()) # signal to function

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(15)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.xml_path.sizePolicy().hasHeightForWidth())
        self.xml_path.setSizePolicy(sizePolicy)
        self.xml_path.setBaseSize(QtCore.QSize(0, 0))
        self.xml_path.setObjectName("xml_path")
        self.horizontalLayout_4.addWidget(self.xml_path)
        self.xml_browse_button = QtWidgets.QPushButton(self.verticalLayoutWidget)

        self.xml_browse_button.clicked.connect(lambda: self.browse_xml())  # signal to browse

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.xml_browse_button.sizePolicy().hasHeightForWidth())
        self.xml_browse_button.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.xml_browse_button.setFont(font)
        self.xml_browse_button.setObjectName("xml_browse_button")
        self.horizontalLayout_4.addWidget(self.xml_browse_button)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.xlsx_path = QtWidgets.QLineEdit(self.verticalLayoutWidget)

        self.xlsx_path.textChanged.connect(lambda: self.undisable_button())  # signal to function

        self.xlsx_path.setEnabled(True)
        self.xlsx_path.setText("")
        self.xlsx_path.setObjectName("xlsx_path")
        self.horizontalLayout_5.addWidget(self.xlsx_path)
        self.xlsx_browse_button = QtWidgets.QPushButton(self.verticalLayoutWidget)

        self.xlsx_browse_button.clicked.connect(lambda: self.browse_xlsx())  # signal to browse

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.xlsx_browse_button.sizePolicy().hasHeightForWidth())
        self.xlsx_browse_button.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.xlsx_browse_button.setFont(font)
        self.xlsx_browse_button.setObjectName("xlsx_browse_button")
        self.horizontalLayout_5.addWidget(self.xlsx_browse_button)
        self.verticalLayout.addLayout(self.horizontalLayout_5)
        spacerItem = QtWidgets.QSpacerItem(15, 16, QtWidgets.QSizePolicy.Policy.Minimum,
                                           QtWidgets.QSizePolicy.Policy.Preferred)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.xlsx_name_label = QtWidgets.QLabel(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.xlsx_name_label.sizePolicy().hasHeightForWidth())
        self.xlsx_name_label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.xlsx_name_label.setFont(font)
        self.xlsx_name_label.setObjectName("xlsx_name_label")
        self.horizontalLayout_6.addWidget(self.xlsx_name_label)
        self.xlsx_name = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.xlsx_name.setObjectName("xlsx_name")
        self.horizontalLayout_6.addWidget(self.xlsx_name)
        self.verticalLayout.addLayout(self.horizontalLayout_6)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Policy.Minimum,
                                            QtWidgets.QSizePolicy.Policy.Minimum)
        self.verticalLayout.addItem(spacerItem1)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.export_button = QtWidgets.QPushButton(self.verticalLayoutWidget)

        self.export_button.setDisabled(True)
        self.export_button.clicked.connect(lambda: self.export())  # signal to browse

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.export_button.sizePolicy().hasHeightForWidth())
        self.export_button.setSizePolicy(sizePolicy)
        self.export_button.setMinimumSize(QtCore.QSize(189, 27))
        self.export_button.setBaseSize(QtCore.QSize(0, 0))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.export_button.setFont(font)
        self.export_button.setObjectName("export_button")
        self.horizontalLayout_7.addWidget(self.export_button)
        self.verticalLayout.addLayout(self.horizontalLayout_7)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 600, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.db_path_action = QtGui.QAction(MainWindow)

        self.db_path_action.triggered.connect(lambda: self.openDBWindow())  # signal to next window

        self.db_path_action.setObjectName("db_path_action")
        self.menu.addAction(self.db_path_action)
        self.menubar.addAction(self.menu.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "eInvoice XML to Excel"))
        self.label.setText(_translate("MainWindow", "Խնդրում եմ ներմուծել eInvoice-ից արտահանված xml ֆայլը"))
        self.xml_browse_button.setText(_translate("MainWindow", "Փնտրել"))
        self.label_2.setText(_translate("MainWindow", "Խնդրում եմ նշել թե որտեղ արտահանել excel ֆայլը"))
        self.xlsx_browse_button.setText(_translate("MainWindow", "Փնտրել"))
        self.xlsx_name_label.setText(_translate("MainWindow", "Գրեք արտահանվող excel-ի անունը /optional/"))
        self.export_button.setText(_translate("MainWindow", "Արտահանել excel ֆայլը"))
        self.menu.setTitle(_translate("MainWindow", "Կարգավորումներ"))
        self.db_path_action.setText(_translate("MainWindow", "ՀԾ հաճախորդների ուղի"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
