from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
import urllib
import psycopg2
import pandas as pd
import sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from os.path import exists
import openpyxl
import re

#for memory usage limitations:
from limiter import limit_memory
limit_memory(1000) # 1000 megs maximum

VERSION = "v20211113"

xlsx_location = "./PT_testing.xlsx"
results_xlsx = "results.xlsx"
keyword1 = "Webai"
keyword2 = "Keywords"
url_count = 5
keywords_match = 4
column_read = 'B'
read_from_line = 1

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
            column_read = ui.lineEdit_4.text()
            #loading the URLS
            df = pd.read_excel(xlsx_location, sheet_name=keyword1, usecols=column_read, skiprows=read_from_line-2)
            #loading the keywords
            df2 = pd.read_excel(xlsx_location, sheet_name=keyword2, usecols=column_read)
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

        file_exists = exists(results_xlsx)
        if not file_exists:
            df3 = pd.DataFrame([['', '', '']],columns=['Link', 'Result', 'Email'])
            df3.to_excel(results_xlsx, index=False)

        temp_url_count = url_count
        try:
            ui.textBrowser.clear()
            yes = 0
            manual = 0
            no = 0

            for URL in df.values.tolist():
                try:
                    isNo = False #dont print the line if it doesnt match the keywords
                    email = []

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

                    website_text = soup.get_text()
                    for keyword in df2.values.tolist():
                        if keyword[0] in website_text:
                            num += 1

                    #print(" ".join(soup.get_text().split()))

                    if num >= keywords_match:
                        print('true')
                        text[1] += " - YES "
                        yes += 1
                        #searching for email:
                        email = re.search(".*@.*(pt)$", website_text)
                    else:
                        print('false')
                        no += 1
                        isNo = True

                except Exception as e:
                    print(e)
                    text[1] += " - ?? "
                    manual += 1
                    #searching for email:
                    email = re.search(".*@.*pt", website_text)
                    pass

                #if the url passed
                if not isNo:
                    #writing to excel file
                    try:
                        workbook = openpyxl.load_workbook(results_xlsx)
                        worksheet = workbook["Sheet1"]

                        if email is None:
                            email = [None] * 2
                            email[0] = ""

                        worksheet.append([web_url, text[1], email[0]])
                        workbook.save(results_xlsx)

                    except Exception as e:
                        print(e)

                        if results_xlsx in str(e):
                            Main.clearScreen()
                            ui.updateText(MainWindow, ["Error with xlsx file. Try closing the {0} file!".format(results_xlsx), ""])
                            break
                        else:
                            pass

                    ui.updateText(MainWindow, text)
                print("Link number: {0}".format(yes+no+manual))

            text[0] = "<br> - Total yes: {0}, no: {1} check manually: {2}.".format(yes, no, manual)
            ui.updateText(MainWindow, text)

        except AttributeError as e:
            print(e)
            ui.updateText(MainWindow, ["Database not loaded!", ""])
        except Exception as e:
            print(e)
            pass

    def changeNumberOfLinks():
        global url_count
        url_count = ui.spinBox.value()

    def changeNumberOfKeywords():
        global keywords_match
        keywords_match = ui.spinBox_2.value()

    def changeFromWhichLineToRead():
        global read_from_line
        read_from_line = ui.spinBox_3.value()

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
        self.pushButton.setGeometry(QRect(420, 290, 221, 23))
        self.pushButton.clicked.connect(Main.load_data_from_xml)

        self.spinBox = QSpinBox(self.centralwidget)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(QRect(580, 130, 61, 22))
        self.spinBox.setValue(url_count)
        self.spinBox.setMinimum(1)
        self.spinBox.setMaximum(10000)
        self.spinBox.valueChanged.connect(Main.changeNumberOfLinks)

        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(410, 130, 111, 20))

        self.textBrowser = QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName(u"textBrowser")
        self.textBrowser.setGeometry(QRect(10, 10, 391, 511))
        self.textBrowser.setOpenExternalLinks(True)
        self.textBrowser.setStyleSheet("font-size: 15px;")
        self.textBrowser.setText("TimberScraper started")

        self.spinBox_2 = QSpinBox(self.centralwidget)
        self.spinBox_2.setObjectName(u"spinBox_2")
        self.spinBox_2.setGeometry(QRect(580, 160, 61, 22))
        self.spinBox_2.setValue(keywords_match)
        self.spinBox_2.setMinimum(1)
        self.spinBox_2.valueChanged.connect(Main.changeNumberOfKeywords)

        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(410, 160, 91, 16))
        self.label_2.setTextFormat(Qt.PlainText)

        self.pushButton_2 = QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(420, 320, 221, 23))
        self.pushButton_2.clicked.connect(Main.scrape)

        self.lineEdit = QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(532, 10, 111, 20))

        self.label_3 = QLabel(self.centralwidget)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(410, 10, 71, 16))

        self.lineEdit_2 = QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName(u"lineEdit_2")
        self.lineEdit_2.setGeometry(QRect(532, 40, 111, 20))

        self.lineEdit_3 = QLineEdit(self.centralwidget)
        self.lineEdit_3.setObjectName(u"lineEdit_3")
        self.lineEdit_3.setGeometry(QRect(532, 70, 111, 20))

        self.label_4 = QLabel(self.centralwidget)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(410, 40, 81, 16))

        self.label_5 = QLabel(self.centralwidget)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(410, 70, 111, 16))

        self.pushButton_3 = QPushButton(self.centralwidget)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setGeometry(QRect(420, 350, 221, 23))
        self.pushButton_3.clicked.connect(Main.clearScreen)

        self.lineEdit_4 = QLineEdit(self.centralwidget)
        self.lineEdit_4.setObjectName(u"lineEdit_4")
        self.lineEdit_4.setGeometry(QRect(620, 190, 21, 20))
        self.lineEdit_4.setText(column_read)

        self.label_6 = QLabel(self.centralwidget)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(410, 190, 151, 16))

        self.spinBox_3 = QSpinBox(self.centralwidget)
        self.spinBox_3.setObjectName(u"spinBox_3")
        self.spinBox_3.setGeometry(QRect(530, 100, 111, 22))
        self.spinBox_3.setMinimum(2)
        self.spinBox_3.setMaximum(1000000)
        self.spinBox_3.setValue(read_from_line)
        self.spinBox_3.valueChanged.connect(Main.changeFromWhichLineToRead)

        self.label_7 = QLabel(self.centralwidget)
        self.label_7.setObjectName(u"label_7")
        self.label_7.setGeometry(QRect(410, 100, 71, 16))

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
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"TimberScraper {0}".format(VERSION), None))
        self.pushButton.setText(QCoreApplication.translate("MainWindow", u"Load database", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"Number of links to load", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"Keywords to match", None))
        self.pushButton_2.setText(QCoreApplication.translate("MainWindow", u"Scrape", None))
        self.lineEdit.setText(QCoreApplication.translate("MainWindow", u"{0}".format(xlsx_location), None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u".xlsx database", None))
        self.lineEdit_2.setText(QCoreApplication.translate("MainWindow", u"{0}".format(keyword1), None))
        self.lineEdit_3.setText(QCoreApplication.translate("MainWindow", u"{0}".format(keyword2), None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"Links sheet name", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"Keywords sheet name", None))
        self.pushButton_3.setText(QCoreApplication.translate("MainWindow", u"Clear screen", None))
        self.label_6.setText(QCoreApplication.translate("MainWindow", u"Links and keywords in column", None))
        self.label_7.setText(QCoreApplication.translate("MainWindow", u"Load from line", None))
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
