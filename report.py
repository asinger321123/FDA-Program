import sys
import os
import xlrd
import win32com.client as win32
from win32com.client import Dispatch
from pandas.io.excel import ExcelWriter
import pandas
from itertools import zip_longest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from PyQt4 import QtCore, QtGui, uic
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from PyQt4.QtWebKit import QWebView
from dateutil import parser
import datetime
import time
from shutil import copyfile
import unicodecsv
import csv
import subprocess
import openpyxl
import xlwt
import sqlite3
import warnings
from webdriver_manager.chrome import ChromeDriverManager


userhome = os.path.expanduser('~')
desktop = os.path.join(userhome, 'Desktop')
downloads = os.path.join(userhome, 'Downloads')
fdaReports = os.path.join(desktop, 'FDA Reports')

qtCreatorFile = os.path.join(desktop, 'FDA Program','Main.ui') # Enter file here.
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

dt = str(datetime.datetime.now().strftime("%m_%d"))
macro_dt = str(datetime.datetime.now().strftime("%m_%d_%Y"))
month_year = str(datetime.datetime.now().strftime("%B %Y"))
year_month_day = str(datetime.datetime.now().strftime("%Y%m%d"))
monthDayYear = str(datetime.datetime.now().strftime("%m/%d/%Y"))

thisMonth = str(datetime.datetime.now().strftime("%B"))
monthDay = str(datetime.datetime.now().strftime("%B_%d"))
thisYear = str(datetime.datetime.now().strftime("%Y"))

defaultMonth = int(datetime.datetime.now().strftime("%m")) - 1

drugReportName2 = 'DrugsFDA FDA-Approved Drugs.csv'
drugReportName = 'DrugsFDA FDA-Approved Drugs.xlsx'
cleanMacro = os.path.join(fdaReports, 'Macro Report', 'Clean', 'Macro Template Drug Surveillance Changes.xlsm')
macroDestination = os.path.join(fdaReports, 'Macro Report', 'Macro Template Drug Surveillance Changes_'+macro_dt+'.xlsm')



class MyApp(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        # self.currentFileForCompare = ''

        self.monthChoice = self.findChild(QComboBox, 'month_comboBox')
        self.yearChoice = self.findChild(QComboBox, 'year_comboBox')
        self.downloadReport = self.findChild(QPushButton, 'downloadReport_pushButton')
        self.compareFilePath = self.findChild(QLineEdit, 'filePath_lineEdit')
        self.browserButton = self.findChild(QPushButton, 'browser_pushButton')
        self.compareButton = self.findChild(QPushButton, 'compareReports_pushButton')
        self.currentFileForCompare = self.findChild(QLineEdit, 'downloadedFile_lineEdit')
        self.browserButton2 = self.findChild(QPushButton, 'browser2_pushButton')

        self.downloadReport.clicked.connect(self.returnFDAReport)
        self.browserButton.clicked.connect(self.getfiles)
        self.browserButton2.clicked.connect(self.getfiles2)
        # self.compareButton.clicked.connect(self.compareReports)
        self.compareButton.clicked.connect(self.compareReports)


        self.monthChoice.setCurrentIndex(defaultMonth)
        self.yearChoice.setCurrentIndex(9)

    def setCompareFile(self, file=None):
        fullFilePath = str(file[0])
        filenames = str(file[0]).rsplit('/', 1)[-1]
        return fullFilePath 


    def setCurrentFile(self, file=None):
        fullFilePath = str(file[0])
        filenames = str(file[0]).rsplit('/', 1)[-1]
        return fullFilePath     


    def copyMacro(self):
        copyfile(cleanMacro, macroDestination)


    def compareReports(self):
        warnings.simplefilter("ignore")
        self.copyMacro()
        # subprocess.call(['C:\\Users\\asinger\\Desktop\\FDA Program\\excomp.bat', str(self.compareFilePath.text()), str(self.currentFileForCompare)])

        df = pandas.read_excel(str(self.currentFileForCompare.text()))
        df2 = pandas.read_excel(str(self.compareFilePath.text()))

        conn = sqlite3.connect("test_table.db")
        conn.text_factory = str
        cur = conn.cursor()

        df.to_sql('current', conn, if_exists='replace', index=False)
        df2.to_sql('previous', conn, if_exists='replace', index=False)

        wb1 = openpyxl.load_workbook(macroDestination, keep_vba=True)
        ws1 = wb1.active

        export = """--join previous report to current report on drug name and take only the joined drugs to get our matches
        DROP TABLE IF EXISTS matched_drugs;
        CREATE TABLE matched_drugs AS SELECT *
                                        FROM current
                                             LEFT JOIN
                                             previous ON current.[Drug Name] = previous.[Drug Name] AND 
                                                       current.Submission = previous.Submission
                                       WHERE previous.[Drug Name] IS NOT NULL;


        -- find rows that join 1 ti many or many to many
        DROP TABLE IF EXISTS multi_joins;
        CREATE TABLE multi_joins as
        SELECT count("Drug Name" || "Submission"),
               "Drug Name" || "Submission" AS merged,
               "Drug Name"
          FROM matched_drugs
         GROUP BY merged
        HAVING count(merged) > 1;


        --delete the rows when Drug Name and Submission are equal but third check on Active ingredient is not equal
        DELETE FROM matched_drugs
        WHERE "Drug Name" IN (
                SELECT "Drug Name"
                  FROM multi_joins
            )
            AND 
            "Drug Name" = "Drug Name:1" AND 
            "Submission" = "Submission:1" AND 
            "Active Ingredients" != "Active Ingredients:1";


        --create table and check if any column differ from previous week
        DROP TABLE IF EXISTS matched_drug_changes;
        CREATE TABLE matched_drug_changes AS SELECT *
                                               FROM matched_drugs
                                              WHERE "Active Ingredients" <> "Active Ingredients:1" OR 
                                                    "Company" <> "Company:1" OR 
                                                    "Submission Classification *" <> "Submission Classification *:1" OR 
                                                    "Submission Status" <> "Submission Status:1";

               
        --Join to find newly added drugs
        DROP TABLE IF EXISTS new_drugs;
        CREATE TABLE new_drugs AS SELECT *
                                    FROM current
                                         LEFT JOIN
                                         previous ON current.[Drug Name] = previous.[Drug Name] AND 
                                                   current.Submission = previous.Submission
                                   WHERE previous."Drug Name" IS NULL;


        --opposite Join to find anything deleted from the previous day/compare
        DROP TABLE IF EXISTS deleted_drugs;
        CREATE TABLE deleted_drugs AS SELECT *
                                        FROM previous
                                             LEFT JOIN
                                             current ON previous.[Drug Name] = current.[Drug Name] AND 
                                                       previous.Submission = current.Submission
                                       WHERE current."Drug Name" IS NULL;

        DROP TABLE IF EXISTS final;
        CREATE TABLE final AS SELECT "Approval Date",
                                     "Drug Name",
                                     "Submission",
                                     "Active Ingredients",
                                     "Company",
                                     "Submission Classification *",
                                     "Submission Status"
                                FROM matched_drug_changes
        UNION
        SELECT "Approval Date",
               "Drug Name",
               "Submission",
               "Active Ingredients",
               "Company",
               "Submission Classification *",
               "Submission Status"
          FROM new_drugs;"""

        cur.executescript(export)
        conn.commit()

        export = """Select * from final;"""

        multi_Joins = """Select * FROM multi_joins"""
        multi_results = cur.execute(multi_Joins)
        for row in multi_results:
            delRow = row[2]
            print(delRow, ": Joined Multiple Times. Attempted to Automatically Make correct Join but Please double Check.")

        checkCount = """SELECT sum(count) 
  FROM (
           SELECT count( * ) AS count
             FROM matched_drugs
           UNION
           SELECT count( * ) AS count
             FROM new_drugs
       );"""

        results = cur.execute(checkCount)

        for row in results:
            mergedCount = row[0]

        checkCount2 = """SELECT count(*) as count FROM current"""

        results2 = cur.execute(checkCount2)

        for row in results2:
            currentCount = row[0]

        if currentCount == mergedCount:
            print('All Compared Rows Accounted For\n')
        else:
            print('All Compared Rows Unaccounted For:\nCurrent Report: {} Rows Found\nCompared Rows: {} Rows Found'.format(currentCount, mergedCount))


        deletedRows = """SELECT count(*) as count, * FROM deleted_drugs"""

        results3 = cur.execute(deletedRows)

        for row in results3:
            deletedDrugs = row[1:8]
            print('Deleted from Previous Report: ', deletedDrugs)

        pandas.read_sql_query(export, conn).to_csv(os.path.join(fdaReports, 'FDA Compare Results.csv'), index=False, sep=',')

        with open(os.path.join(fdaReports, 'FDA Compare Results.csv'), 'r') as inFile:
            reader = csv.reader(inFile)
            next(reader)
            for row in reader:
                ws1.append(['']+row)

        wb1.save(macroDestination)

        print('Complete')

        #Open FDA Compare Results

        #write results to macroDestination FOolder



        # if os.path.exists(macroDestination):
        xl = win32.Dispatch('Excel.Application')
        xl.DisplayAlerts = False
        xl.Visible = True

        # xl=win32com.client.Dispatch("Excel.Application")
        # xl.Workbooks.Open(os.path.abspath(macroDestination))
        xl.Workbooks.Open(macroDestination)
        time.sleep(2)
        subprocess.call([os.path.join(desktop, 'FDA Program\\runMacro.exe')])
        # xl.SaveAs(macroDestination)
        # xl.Application.Quit()
        # del xl


    def getfiles(self):
        dlg = QFileDialog()
        dlg.setDirectory(fdaReports)
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setFilter("FDA Excel Reports (*.xlsx)")
        filenames = list()
        
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            # filenames = str(filenames[0]).rsplit('/', 1)[-1]
            self.fileChoice = filenames

            self.setCompareFile(self.fileChoice)
            self.compareFilePath.setText(str(self.setCompareFile(self.fileChoice)).replace('/', '\\'))
            # self.compareFilePath.setText(str(self.fileChoice))

    def getfiles2(self):
        dlg = QFileDialog()
        dlg.setDirectory(fdaReports)
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setFilter("FDA Excel Reports (*.xlsx)")
        filenames = list()
        
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            # filenames = str(filenames[0]).rsplit('/', 1)[-1]
            self.fileChoice = filenames

            self.setCurrentFile(self.fileChoice)
            self.currentFileForCompare.setText(str(self.setCurrentFile(self.fileChoice)).replace('/', '\\'))
            # self.compareFilePath.setText(str(self.fileChoice))


    def storeReports(self):

        #Initial Desktop FOlder to store Reports
        if os.path.exists(os.path.join(desktop, 'FDA Reports')):
            reportsParent = os.path.join(desktop, 'FDA Reports')
            pass
        else:
            os.mkdir(os.path.join(desktop, 'FDA Reports'))
            reportsParent = os.path.join(desktop, 'FDA Reports')

        if str(self.monthChoice.currentText()) == thisMonth and str(self.yearChoice.currentText()) == thisYear:

            os.chdir(reportsParent)
            if os.path.exists(os.path.join(reportsParent, month_year)):
                currentMonth = os.path.join(reportsParent, month_year)
                pass
            else:
                os.mkdir(os.path.join(reportsParent, month_year))
                currentMonth = os.path.join(reportsParent, month_year)

            #rename downloaded file to current month_date
            currentFile = 'FDA Approvals_'+dt+'.xlsx'
            if os.path.exists(os.path.join(downloads, currentFile)):
                pass
            else:
                os.rename(os.path.join(downloads, drugReportName), os.path.join(downloads, 'FDA Approvals_'+dt+'.xlsx'))
            

            #copy file to month year folde ron desktop
            if os.path.exists(os.path.join(currentMonth, currentFile)):
                self.currentFileForCompare.setText(str(os.path.join(currentMonth, currentFile)))
                pass
            else:
                copyfile(os.path.join(downloads, currentFile), os.path.join(currentMonth, currentFile))
                self.currentFileForCompare.setText(str(os.path.join(currentMonth, currentFile)))

            #clean up files in downloads
            if os.path.exists(os.path.join(downloads, currentFile)):
                os.remove(os.path.join(downloads, currentFile))
            if os.path.exists(os.path.join(downloads, drugReportName)):
                os.remove(os.path.join(downloads, drugReportName))

        else:
                        #USer choice date conversions
            userMonthChoice = str(self.monthChoice.currentText())
            userYearChoice = str(self.yearChoice.currentText())
            convertedDate = parser.parse('{} {}'.format(userMonthChoice, userYearChoice))
            monthYearChoice = datetime.datetime.strftime(convertedDate, '%B %Y')
            mm_YYChoice = datetime.datetime.strftime(convertedDate, '%m_%d')

            currentFile = 'FDA Approvals_'+mm_YYChoice+'.xlsx'
            os.chdir(reportsParent)
            if os.path.exists(os.path.join(reportsParent, str(self.monthChoice.currentText())+' '+str(self.yearChoice.currentText()))):
                currentMonth = os.path.join(reportsParent, str(self.monthChoice.currentText())+' '+str(self.yearChoice.currentText()))
                self.currentFileForCompare.setText(str(os.path.join(currentMonth, currentFile)))
                pass
            else:
                os.mkdir(os.path.join(reportsParent, str(self.monthChoice.currentText())+' '+str(self.yearChoice.currentText())))
                currentMonth = os.path.join(reportsParent, str(self.monthChoice.currentText())+' '+str(self.yearChoice.currentText()))
                self.currentFileForCompare.setText(str(os.path.join(currentMonth, currentFile)))

            if os.path.exists(os.path.join(downloads, currentFile)):
                pass
            else:
                os.rename(os.path.join(downloads, drugReportName), os.path.join(downloads, 'FDA Approvals_'+mm_YYChoice+'.xlsx'))

            #copy file to month year folde ron desktop
            if os.path.exists(os.path.join(currentMonth, currentFile)):
                # if os.path.exists(macroDestination):
                pass
            else:
                copyfile(os.path.join(downloads, currentFile), os.path.join(currentMonth, currentFile))

            #clean up files in downloads
            if os.path.exists(os.path.join(downloads, currentFile)):
                os.remove(os.path.join(downloads, currentFile))
            if os.path.exists(os.path.join(downloads, drugReportName)):
                os.remove(os.path.join(downloads, drugReportName))

    def returnChromeVersion(self, filename):
        parser = Dispatch("Scripting.FileSystemObject")
        version = parser.GetFileVersion(filename)
        return version


    def returnFDAReport(self):
        # month = raw_input('Enter Month: ')
        # year = raw_input('Enter Year: ')

        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get("https://www.accessdata.fda.gov/scripts/cder/daf/index.cfm?event=reportsSearch.process")

        select1 = Select(driver.find_element_by_id('reportSelectMonth'))
        select2 =Select(driver.find_element_by_id('reportSelectYear'))

        # select by visible text
        select1.select_by_visible_text(str(self.monthChoice.currentText()))
        select2.select_by_visible_text(str(self.yearChoice.currentText()))

        elem = driver.find_elements_by_xpath("""//*[@id="mp-pusher"]/div/div/div/div/div[3]/div/form/div[2]/button""")[0].click()

        #csvButton
        elem = driver.find_elements_by_xpath("""//*[@id="example_1_wrapper"]/div[1]/a[1]""")[0].click()

        #excelButton
        # elem = driver.find_elements_by_xpath("""//*[@id="example_1_wrapper"]/div[1]/a[2]""")[0].click()

        time.sleep(10)
        driver.close()

        raw = pandas.read_csv(os.path.join(downloads, drugReportName2))
        raw.to_excel(os.path.join(downloads, drugReportName), index=False)

        # with ExcelWriter(os.path.join(downloads, drugReportName)) as ew:
        #     pandas.read_csv(os.path.join(downloads, drugReportName2)).to_excel(ew, index=False)

        os.remove(os.path.join(downloads, drugReportName2))
        
        self.storeReports()

    # def get_version_via_com(filename):
    #     parser = Dispatch("Scripting.FileSystemObject")
    #     version = parser.GetFileVersion(filename)
    #     return version

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    QtGui.QApplication.setStyle(QtGui.QStyleFactory.create('plastique'))

    window.show()
    sys.exit(app.exec_())