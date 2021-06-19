import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, \
    QDesktopWidget, QVBoxLayout, QHBoxLayout, QLabel, \
    QGroupBox, QGridLayout, QLineEdit, QPushButton, \
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, \
    QAbstractItemView, QTableView
from PyQt5 import QtCore

class MyApp(QWidget):
    dailyselectedindex = -1
    totalselectedindex = -1
    today = ''
    dailycheck = 0
    totalcheck = 0
    total = []
    daily = []
    man = []
    searchkeyword = ''
    columnheaders = ['이름', '생년월일', '주소', '연락처']

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Babfor (MadeBy. 근로장학생)')
        self.resize(1000, 600)
        self.center()
        self.show()

        vbox = QVBoxLayout()
        self.setLayout(vbox)
        self.dailylist = QTableWidget(self)
        self.dailylist.setColumnCount(4)
        self.dailylist.verticalHeader().setVisible(False)
        self.dailylist.setHorizontalHeaderLabels(self.columnheaders)
        self.dailylist.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.dailylist.setSelectionBehavior(QTableView.SelectRows)
        self.dailylist.setSelectionMode(QAbstractItemView.SingleSelection)
        self.dailylist.cellClicked.connect(self.dailyClickMan)

        self.totallist = QTableWidget(self)
        self.totallist.setColumnCount(4)
        self.totallist.verticalHeader().setVisible(False)
        self.totallist.setHorizontalHeaderLabels(self.columnheaders)
        self.totallist.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.totallist.setSelectionBehavior(QTableView.SelectRows)
        self.totallist.setSelectionMode(QAbstractItemView.SingleSelection)
        self.totallist.cellClicked.connect(self.totalClickMan)

        self.name = QLineEdit()
        self.name.setPlaceholderText('홍길동')

        self.birthday = QLineEdit()
        self.birthday.setMaxLength(6)
        self.birthday.setPlaceholderText('970614')

        self.addr = QLineEdit()
        self.addr.setPlaceholderText('서울시 동대문구 ...')

        self.phonenum = QLineEdit()
        self.phonenum.setPlaceholderText('010-1234-5678')

        vbox.addWidget(self.createWeatherGroup())
        vbox.addWidget(self.createMainGroup())

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def createWeatherGroup(self): # 날짜
        wgroup = QGroupBox()
        hbox = QHBoxLayout()
        wgroup.setLayout(hbox)

        self.today = QtCore.QDate.currentDate().toString('yyyy-MM-dd')
        hbox.addStretch(1)
        hbox.addWidget(QLabel(self.today))
        hbox.addStretch(1)

        return wgroup

    def createMainGroup(self):
        mgroup = QGroupBox()
        hbox = QHBoxLayout()
        mgroup.setLayout(hbox)

        inputform = self.createInputForm()
        hbox.addWidget(inputform)
        list = self.createList()
        hbox.addWidget(list)
        hbox.setStretchFactor(inputform, 1)
        hbox.setStretchFactor(list, 2)

        return mgroup

    def createInputForm(self):
        igroup = QGroupBox('명단 입력')
        vbox = QVBoxLayout()
        igroup.setLayout(vbox)

        searchgroup = QGroupBox()
        hbox1 = QHBoxLayout()
        searchgroup.setLayout(hbox1)
        searchname = QLineEdit()
        searchname.setPlaceholderText('이름 검색')
        searchname.textChanged[str].connect(self.searchChange)
        hbox1.addWidget(searchname)
        searchbutton = QPushButton('검색')
        searchbutton.clicked.connect(self.search)
        hbox1.addWidget(searchbutton)
        vbox.addWidget(searchgroup)
        vbox.addStretch(1)

        form = QGroupBox()
        grid = QGridLayout()
        grid.setColumnStretch(1, 1)
        form.setLayout(grid)
        grid.addWidget(QLabel('이름 : '), 0, 0)

        grid.addWidget(self.name, 0, 1)
        grid.addWidget(QLabel('생년월일 : '), 1, 0)

        grid.addWidget(self.birthday, 1, 1)
        grid.addWidget(QLabel('주소 : '), 2, 0)

        grid.addWidget(self.addr, 3, 0, 1, 2)
        grid.addWidget(QLabel('전화번호 : '), 4, 0)

        grid.addWidget(self.phonenum, 5, 0, 1, 2)
        vbox.addWidget(form)
        vbox.addStretch(1)

        subgroup = QGroupBox()
        hbox2 = QHBoxLayout()
        subgroup.setLayout(hbox2)
        hbox2.addStretch(1)
        clear = QPushButton('초기화')
        clear.setFixedSize(100, 40)
        clear.clicked.connect(self.clearForm)
        hbox2.addWidget(clear)

        submit = QPushButton('추가')
        submit.setFixedSize(100, 40)
        submit.clicked.connect(self.dailyAdd)
        hbox2.addWidget(submit)

        hbox2.addStretch(1)

        vbox.addWidget(subgroup)
        return igroup

    def createList(self):
        # ListGroup
        lgroup = QGroupBox()
        vbox = QVBoxLayout()
        lgroup.setLayout(vbox)

        lgsubgroup = QGroupBox()
        hbox = QHBoxLayout()
        lgsubgroup.setLayout(hbox)
        hbox.addWidget(self.createDailyList())
        hbox.addWidget(self.createTotalList())

        vbox.addWidget(lgsubgroup)

        return lgroup

    def createDailyList(self):
        dlgroup = QGroupBox('오늘 명단')
        vbox = QVBoxLayout()
        dlgroup.setLayout(vbox)

        vbox.addWidget(self.dailylist)

        subgroup = QGroupBox()
        hbox = QHBoxLayout()
        subgroup.setLayout(hbox)
        hbox.addStretch(1)
        delete = QPushButton('명단 삭제')
        delete.setFixedSize(100, 24)
        delete.clicked.connect(self.realDailyDelete)
        hbox.addWidget(delete)
        submit = QPushButton('파일 열기')
        submit.setFixedSize(100, 24)
        submit.clicked.connect(self.showDailyFileDialog)
        hbox.addWidget(submit)
        sort = QPushButton('정렬')
        sort.setFixedSize(100, 24)
        sort.clicked.connect(self.dailySort)
        hbox.addWidget(sort)
        hbox.addStretch(1)
        vbox.addWidget(subgroup)

        subgroup2 = QGroupBox()
        hbox2 = QHBoxLayout()
        subgroup2.setLayout(hbox2)
        hbox2.addStretch(1)
        save = QPushButton('저장')
        save.setFixedSize(150, 40)
        save.clicked.connect(self.saveDailyList)
        hbox2.addWidget(save)
        hbox2.addStretch(1)
        vbox.addWidget(subgroup2)
        return dlgroup

    def createTotalList(self):
        tlgroup = QGroupBox('전체 명단')
        vbox = QVBoxLayout()
        tlgroup.setLayout(vbox)

        vbox.addWidget(self.totallist)

        subgroup = QGroupBox()
        hbox = QHBoxLayout()
        subgroup.setLayout(hbox)
        hbox.addStretch(1)
        delete = QPushButton('명단 삭제')
        delete.setFixedSize(100, 24)
        delete.clicked.connect(self.realTotalDelete)
        hbox.addWidget(delete)
        submit = QPushButton('파일 열기')
        submit.setFixedSize(100, 24)
        submit.clicked.connect(self.showTotalFileDialog)
        hbox.addWidget(submit)
        sort = QPushButton('정렬')
        sort.setFixedSize(100, 24)
        sort.clicked.connect(self.totalSort)
        hbox.addWidget(sort)
        hbox.addStretch(1)
        vbox.addWidget(subgroup)

        subgroup2 = QGroupBox()
        hbox2 = QHBoxLayout()
        subgroup2.setLayout(hbox2)
        hbox2.addStretch(1)
        save = QPushButton('저장')
        save.setFixedSize(150, 40)
        save.clicked.connect(self.saveTotalList)
        hbox2.addWidget(save)
        hbox2.addStretch(1)
        vbox.addWidget(subgroup2)
        return tlgroup

    def showTotalFileDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Open File', '', ' Excel File(*.xlsx *.csv)')

        if fname[0][-4:] == 'xlsx':
            data = pd.read_excel(fname[0], engine='openpyxl', dtype=str)
            data = data.values.tolist()
            data = [map(str, row) for row in data]
            for i in range(len(data)):
                data[i] = list(data[i])

            self.getTotalList(data)
        else:
            self.showDialog('fail')

    def showDailyFileDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Open File', '', 'Excel File(*.xlsx *.csv)')

        if fname[0][-4:] == 'xlsx':
            data = pd.read_excel(fname[0], engine='openpyxl', dtype=str)
            data = data.values.tolist()
            data = [map(str, row) for row in data]
            for i in range(len(data)):
                data[i] = list(data[i])
            self.getDailyList(data)
        else:
            self.showDialog('fail')

    def showDialog(self, message):
        msg = QMessageBox()
        if message == 'fail':
            msg.setWindowTitle('잘못된 파일 형식입니다.')
            msg.setText('잘못된 파일 형식입니다.')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

        elif message == 'name':
            msg.setWindowTitle('이름을 입력해주세요.')
            msg.setText('이름을 입력해주세요.')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
        elif message == 'columnproblem':
            msg.setWindowTitle('파일의 양식이 맞지 않습니다.')
            msg.setText('파일의 양식이 맞지 않습니다.')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
    def getTotalList(self, data):
        self.totallist.setRowCount(0)
        self.total = data
        for row in self.total:
            rowPosition = self.totallist.rowCount()
            self.totallist.insertRow(rowPosition)
            for i, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.totallist.setItem(rowPosition, i, item)

    def getDailyList(self, data):
        self.dailylist.setRowCount(0)
        self.daily = data
        for row in self.daily:
            rowPosition = self.dailylist.rowCount()
            self.dailylist.insertRow(rowPosition)
            for i, val in enumerate(row):

                item = QTableWidgetItem(str(val))
                self.dailylist.setItem(rowPosition, i, item)

    def search(self):
        check = 0
        for index, data in enumerate(self.daily):
            if data[0] == self.searchkeyword:
                self.dailylist.selectRow(index)
                self.man = data
                self.updateInputForm()
                self.dailycheck = 1
                check = 1
                break;
        for index, data in enumerate(self.total):
            if data[0] == self.searchkeyword:
                self.totallist.selectRow(index)
                self.man = data
                self.updateInputForm()
                self.totalcheck = 1
                check = 2
                break
        if check == 0:
            msg = QMessageBox();
            msg.setWindowTitle('명단을 찾을 수 없습니다.')
            msg.setText('명단을 찾을 수 없습니다.')
            msg.setStandardButtons(QMessageBox.Cancel)
            msg.setIcon(QMessageBox.Information)
            msg.exec_()

    def searchChange(self, text):
        self.searchkeyword = text

    def dailySort(self):
        if self.daily:
            self.daily.sort(key=lambda x: x[0])
            self.setDailyList()

    def totalSort(self):
        if self.total:
            self.total.sort(key=lambda x: x[0])
            self.setTotalList()

    def dailyAdd(self):
        if self.name.text():
            self.man = [self.name.text(), self.birthday.text(), self.addr.text(), self.phonenum.text()]
            if self.dailycheck != 1:
                self.daily.append(self.man)
                self.dailySort()

            if self.totalcheck != 1:
                self.total.append(self.man)
                self.totalSort()

            self.dailycheck = 0
            self.totalcheck = 0
            self.name.setText('')
            self.birthday.setText('')
            self.addr.setText('')
            self.phonenum.setText('')
        else:
            self.showDialog('name')

    def dailyClickMan(self, data):
        self.dailyselectedindex = data
        if len(self.daily) > data:
            self.man = self.daily[data]
            self.dailycheck = 1
            self.updateInputForm()

    def totalClickMan(self, data):
        self.totalselectedindex = data
        if len(self.total) > data:
            self.man = self.total[data]
            self.totalcheck = 1
            self.updateInputForm()

    def updateInputForm(self):
        if self.man[0]:
            self.name.setText(self.man[0])
        if self.man[1]:
            self.birthday.setText(self.man[1])
        if self.man[2]:
            self.addr.setText(self.man[2])
        if self.man[3]:
            self.phonenum.setText(self.man[3])

    def setDailyList(self):
        self.dailylist.setRowCount(0)
        for row in self.daily:
            rowPosition = self.dailylist.rowCount()
            self.dailylist.insertRow(rowPosition)
            for i, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.dailylist.setItem(rowPosition, i, item)


    def setTotalList(self):
        self.totallist.setRowCount(0)
        for row in self.total:
            rowPosition = self.totallist.rowCount()
            self.totallist.insertRow(rowPosition)
            for i, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.totallist.setItem(rowPosition, i, item)

    def deleteDailyListItem(self):
        if self.dailyselectedindex >= 0:
            del self.daily[self.dailyselectedindex]
            self.setDailyList()
            self.dailycheck = 0
            self.dailyselectedindex = -1

    def deleteTotalListItem(self):
        if self.totalselectedindex >= 0:
            del self.total[self.totalselectedindex]
            self.setTotalList()
            self.totalcheck = 0
            self.totalselectedindex = -1

    def realDailyDelete(self):
        reply = QMessageBox.question(self, '경고', '정말 제거 하시겠습니까?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.deleteDailyListItem()

    def realTotalDelete(self):
        reply = QMessageBox.question(self, '경고', '정말 제거 하시겠습니까?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.deleteTotalListItem()

    def clearForm(self):
        self.man = []
        self.name.setText('')
        self.birthday.setText('')
        self.addr.setText('')
        self.phonenum.setText('')

    def saveDailyList(self):
        try:
            df = pd.DataFrame(self.daily, columns=self.columnheaders)
            df = df.set_index('이름')
            fname = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel File(*.xlsx *.csv)')
            if fname:
                writer = pd.ExcelWriter(fname[0], engine='xlsxwriter')
                sheetname = 'Sheet1'
                df.to_excel(writer, sheet_name=sheetname)

                worksheet = writer.sheets[sheetname]  # pull worksheet object
                worksheet.set_column(0, 0, 10)
                worksheet.set_column(1, 1, 10)
                worksheet.set_column(2, 2, 50)
                worksheet.set_column(3, 3, 20)
                writer.save()
        except:
            self.showDialog('columnproblem')
    def saveTotalList(self):
        try:
            df = pd.DataFrame(self.total, columns=self.columnheaders)
            df = df.set_index('이름')
            fname = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel File(*.xlsx *.csv)')
            if fname:
                writer = pd.ExcelWriter(fname[0], engine='xlsxwriter')
                sheetname = 'Sheet1'
                df.to_excel(writer, sheet_name=sheetname)

                worksheet = writer.sheets[sheetname]  # pull worksheet object
                worksheet.set_column(0, 0, 10)
                worksheet.set_column(1, 1, 10)
                worksheet.set_column(2, 2, 50)
                worksheet.set_column(3, 3, 20)
                writer.save()
        except:
            self.showDialog('columnproblem')

    def closeEvent(self, event):

        reply = QMessageBox.question(self, 'Message',
                                           "종료하시겠습니까?", QMessageBox.Yes |
                                           QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())