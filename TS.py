from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
import urllib
import psycopg2
import socket
import ssl
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *

xlsx_location = "./PT_testing.xlsx"
keyword1 = "Webai"
keyword2 = "Keywords"
url_count = 5
keywords_match = 4

df = []
df2 = []

app = QApplication(sys.argv)
MainWindow = QMainWindow()
ui = None

class Main():
    def window():
        global ui
        ui = Ui_MainWindow()
        ui.setupUi(MainWindow)
        MainWindow.show()
        sys.exit(app.exec_())

    def load_data_from_xml():
        global df, df2
        try:
            #reading the inputs for the xlsx file
            xlsx_location = ui.lineEdit.text()
            keyword1 = ui.lineEdit_2.text()
            keyword2 = ui.lineEdit_3.text()
            #loading the URLS
            df = pd.read_excel(xlsx_location, sheet_name=keyword1, usecols='B')
            #loading the keywords
            df2 = pd.read_excel(xlsx_location, sheet_name=keyword2, usecols='B')
            print('Data loading complete')
            ui.textBrowser.clear()
            ui.updateText(MainWindow, ["Database loaded.", ""])

        except ValueError as e:
            print(e)
            if "Worksheet named" in str(e):
                ui.updateText(MainWindow, ["Excel file is badly structured. Check the sheet names.", ""])
            else:
                ui.updateText(MainWindow, ["Excel file is badly structured. Check the column structure.", ""])
            pass
        except FileNotFoundError as e:
            print(e)
            ui.updateText(MainWindow, ["File not found! Check the directory if the file exists", ""])
            pass

    def scrape():
        print('Scraping started')
        global url_count
        global df, df2

        temp_url_count = url_count
        try:
            ui.textBrowser.clear()
            yes = 0
            no = 0
            manual = 0

            for URL in df.values.tolist():
                try:
                    text = ["",""]
                    temp_url_count -= 1

                    if temp_url_count < 0:
                        break

                    #web scraper logic
                    web_url = "https://www.{}".format(URL[0])
                    text[0] = "<a href={}>{}</a>".format(web_url, web_url) #for printing to screen
                    hdr = {'User-Agent': 'Mozilla/93.0'}
                    try:
                        req = Request(web_url,headers=hdr)
                        page = urlopen(req, timeout=10)
                    except urllib.error.HTTPError as e:
                        web_url = "http://www.{}".format(URL[0])
                        text[0] = "<a href={}>{}</a>".format(web_url, web_url) #for printing to screen
                        req = Request(web_url,headers=hdr)
                        page = urlopen(req, timeout=5)
                        pass

                    html = page.read().decode("utf-8")
                    soup = BeautifulSoup(html, "html.parser")

                    num = 0

                    for keyword in df2.values.tolist():
                        if keyword[0] in soup.get_text():
                            num += 1

                    #print(" ".join(soup.get_text().split()))

                    if num >= keywords_match:
                        print('true')
                        text[1] += " - YES "
                        yes += 1
                    else:
                        print('false')
                        text[1] += " - NO "
                        no += 1

                except UnboundLocalError as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                except urllib.error.HTTPError as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                except socket.timeout as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                except UnicodeDecodeError as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                except ssl.SSLCertVerificationError as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                except urllib.error.URLError as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    pass
                ui.updateText(MainWindow, text)
            text[0] = "<br> - Total yes: {0}, no: {1}, check manually: {2}.".format(yes, no, manual)
            ui.updateText(MainWindow, text)

        except AttributeError as e:
            print(e)
            ui.updateText(MainWindow, ["Database not loaded!", ""])

    def changeNumberOfLinks():
        global url_count
        url_count = ui.spinBox.value()

    def changeNumberOfKeywords():
        global keywords_match
        keywords_match = ui.spinBox_2.value()

    def clearScreen():
        ui.textBrowser.clear()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(654, 571)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")

        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(420, 170, 221, 23))
        self.pushButton.clicked.connect(Main.load_data_from_xml)

        self.spinBox = QSpinBox(self.centralwidget)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(QRect(600, 130, 42, 22))
        self.spinBox.setValue(url_count)
        self.spinBox.valueChanged.connect(Main.changeNumberOfLinks)

        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(420, 130, 111, 20))

        self.textBrowser = QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName(u"textBrowser")
        self.textBrowser.setGeometry(QRect(10, 10, 391, 511))
        self.textBrowser.setOpenExternalLinks(True)
        self.textBrowser.setStyleSheet("font-size: 15px;")
        self.textBrowser.setText("TimberScraper started")

        self.spinBox_2 = QSpinBox(self.centralwidget)
        self.spinBox_2.setObjectName(u"spinBox_2")
        self.spinBox_2.setGeometry(QRect(600, 100, 42, 22))
        self.spinBox_2.setValue(keywords_match)
        self.spinBox_2.valueChanged.connect(Main.changeNumberOfKeywords)

        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(420, 100, 91, 16))
        self.label_2.setTextFormat(Qt.PlainText)

        self.pushButton_2 = QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(420, 200, 221, 23))
        self.pushButton_2.clicked.connect(Main.scrape)

        self.lineEdit = QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(532, 10, 111, 20))

        self.label_3 = QLabel(self.centralwidget)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(420, 10, 71, 16))

        self.lineEdit_2 = QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName(u"lineEdit_2")
        self.lineEdit_2.setGeometry(QRect(532, 40, 111, 20))

        self.lineEdit_3 = QLineEdit(self.centralwidget)
        self.lineEdit_3.setObjectName(u"lineEdit_3")
        self.lineEdit_3.setGeometry(QRect(532, 70, 111, 20))

        self.label_4 = QLabel(self.centralwidget)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(420, 40, 81, 16))

        self.label_5 = QLabel(self.centralwidget)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(420, 70, 111, 16))

        self.pushButton_3 = QPushButton(self.centralwidget)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setGeometry(QRect(420, 230, 221, 23))
        self.pushButton_3.clicked.connect(Main.clearScreen)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 654, 21))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"TimberScraper", None))
        self.pushButton.setText(QCoreApplication.translate("MainWindow", u"Load database", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"Number of links to load", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"Keywords to match", None))
        self.pushButton_2.setText(QCoreApplication.translate("MainWindow", u"Scrape", None))
        self.lineEdit.setText(QCoreApplication.translate("MainWindow", u"filename.xlsx", None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u".xlsx database", None))
        self.lineEdit_2.setText(QCoreApplication.translate("MainWindow", u"sheet1", None))
        self.lineEdit_3.setText(QCoreApplication.translate("MainWindow", u"sheet2", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"Links sheet name", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"Keywords sheet name", None))
        self.pushButton_3.setText(QCoreApplication.translate("MainWindow", u"Clear screen", None))
    # retranslateUi

    def updateText(self, MainWindow, text):
        #text += "<br>"
        finalText = ""

        if text[1] == "":
            finalText = text[0]
        else:
            #formatting logic
            spaces = 10 - len(text[1])
            if "??" in text[1]: #special case
                spaces += 1

            for x in range(spaces):
                text[1] += "_"

            finalText = "{} {}".format(text[1], text[0])

        self.textBrowser.append(finalText)


if __name__ == "__main__":
    Main.window()
