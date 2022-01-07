import os
import shutil
from PyQt5.QtGui import QIcon, QStandardItemModel
from PyQt5.uic.uiparser import QtCore, QtWidgets
import numpy as np
import pandas as pd
import calendar
from datetime import datetime
import matplotlib.pyplot as plt
import time
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic

form_class = uic.loadUiType("panda_qt.ui")[0]

class Data():
    def __init__(self):
        self.__students = []
        self.__maxCor = 0
        d = {"이름": self.__students, "맞은 갯수": [], "점수": [], "순위": []}
        self.__df = pd.DataFrame(data=d)

    #getter
    def getStudents(self):
        return self.__students

    def getMaxCor(self):
        return self.__maxCor

    def getDf(self):
        return self.__df

    #setter
    def setStudnets(self, value):
        self.__students.append(value)
        data_to_insert = {"이름": value, "맞은 갯수": 0, "점수": 0, "순위": 0}
        self.__df = self.__df.append(data_to_insert, ignore_index=True)

    def setMaxCor(self, value):
        self.__maxCor = value

    #calc function
    def calcScore(self):
        for i in range(len(self.__df.index)):
            score = self.__df.iloc[i, 1] / self.__maxCor
            score2 = round(score, 2) * 100
            self.__df.iloc[i, 2] = score2
        
        self.__df.sort_values(by=['점수'], axis=0, ascending=False, inplace=True)

        grade = 1
        for i in range(len(self.__df.index)):
            self.__df.iloc[i, 3] = grade
            grade += 1
        del grade

    def sortStudent(self):
        self.__df.sort_values(by=['이름'], axis=0, ascending=True, inplace=True)

    #delete function
    def delStudent(self, value):
        self.__df.drop(index=self.__df.loc[self.__df.이름 == value].index, inplace=True)


data = Data()

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle('판다 성적 입력 프로그램')
        self.setWindowIcon(QIcon('panda_bear_icon_153300.svg'))
        self.show()

        #quit button
        self.quitButton.clicked.connect(QCoreApplication.instance().quit)
        
        #add, modify and set student button
        self.setStudent.clicked.connect(self.addTableItemDialog)
        self.modifyStudent.clicked.connect(self.modifyTableItemDialog)
        self.delStudent.clicked.connect(self.delTableItem)

        #dateEdit widget
        date = self.dateEdit
        date.setDate(QDate.currentDate())

        #add, del Item to List button
        addBtn = self.addItemBtn
        addBtn.clicked.connect(self.addItemToList)
        delBtn = self.delItemBtn
        delBtn.clicked.connect(self.delItemToList)

        #correct number lineEdit
        MSinput = self.maxScoreInput
        MSinput.textChanged.connect(self.changedScoreSignal)

        #sort button
        sortStuBtn = self.sortByStudentName
        sortScoBtn = self.sortByScore
        sortScoBtn.clicked.connect(self.sortScoreSignal)
        sortStuBtn.clicked.connect(self.sortStudentSignal)
        self.printBtn.clicked.connect(self.test)

        #table widget
        tableHeader = self.tableWidget.horizontalHeader()
        tableHeader.setSectionResizeMode(0, QHeaderView.Stretch)
        tableHeader.setSectionResizeMode(1, QHeaderView.Stretch)
        tableHeader.setSectionResizeMode(2, QHeaderView.Stretch)
        table = self.tableWidget
        table.cellChanged.connect(self.changedTableSignal)

        #all clear button!!!!!!
        self.allClearBtn.setStyleSheet("background-color: red")
        

    def addTableItemDialog(self):
        text, ok = QInputDialog.getText(self, '입력할 학생 정보 입력', '이름: ')

        if ok:
            table = self.tableWidget
            row_count = table.rowCount()
            table.insertRow(row_count)
            table.setVerticalHeaderItem(row_count, QTableWidgetItem(str(text)))
            data.setStudnets(str(text))
    
    def modifyTableItemDialog(self):
        table = self.tableWidget
        cur_row = table.currentRow()
        item = table.verticalHeaderItem(cur_row)

        if type(item) != QTableWidgetItem:
            QMessageBox.about(self, '경고', '선택된 학생이 없습니다.')
            return

        text, ok = QInputDialog.getText(self, '변경할 학생 정보 입력', '이름: ')

        if ok:
            data.getDf().iloc[cur_row, 0] = str(text)

            table.setVerticalHeaderItem(cur_row, QTableWidgetItem(str(text)))

    def delTableItem(self):
        table = self.tableWidget
        cur_row = table.currentRow()
        value = table.verticalHeaderItem(cur_row)
        
        if type(value) == QTableWidgetItem:
            data.delStudent(value.text())
            print(value.text())

        table.removeRow(cur_row)

    def alertDialog(self, alertText):
        QMessageBox.about(self, "경고", alertText)

    def addItemToList(self):
        text = self.titleEdit.text()
        pandaList = self.listWidget

        if text == '':
            pandaDate = self.dateEdit.date().toString(Qt.ISODate)
            pandaGrade = self.comboBox.currentText()
            itemText = f'{pandaGrade} {pandaDate} 성적표'
            pandaList.addItem(itemText)
        else:
            pandaList.addItem(text)
            
    
    def delItemToList(self):
        pandaList = self.listWidget
        cur_row = pandaList.currentRow()
        pandaList.takeItem(cur_row)

    def sortStudentSignal(self):
        data.sortStudent()
        table = self.tableWidget
        r_list = data.getDf().이름
        c_list = ['맞은 갯수', '점수', '순위']

        table.clear()
        table.setHorizontalHeaderLabels(c_list)
        table.setVerticalHeaderLabels(r_list)

        for i in range(len(c_list)):
            for j in range(len(r_list)):
                table.setItem(j, i, QTableWidgetItem(str(data.getDf().iloc[j, i+1])))

    def sortScoreSignal(self):
        if data.getMaxCor() <= 0:
            QMessageBox.about(self, '경고', '모든 학생의 점수를 입력하세요.')
            return
        
        data.calcScore()
        table = self.tableWidget
        
        r_list = data.getDf().이름
        c_list = ['맞은 갯수', '점수', '순위']

        table.clear()
        table.setHorizontalHeaderLabels(c_list)
        table.setVerticalHeaderLabels(r_list)

        for i in range(len(c_list)):
            for j in range(len(r_list)):
                table.setItem(j, i, QTableWidgetItem(str(data.getDf().iloc[j, i+1])))

    def changedScoreSignal(self):
        value = self.maxScoreInput.text()
        if value != '':
            value2 = int(value)
            data.setMaxCor(value2)
            print(data.getMaxCor())
        else:
            return
        
        
    def changedTableSignal(self):
        value = self.tableWidget.currentItem()
        cur_row = self.tableWidget.currentRow()
        cur_col = self.tableWidget.currentColumn()
        if type(value) == QTableWidgetItem:
            print(cur_row, ' ', cur_col)
            data.getDf().iloc[cur_row, cur_col + 1] = int(value.text())

    def test(self):
        print(data.getDf())

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()