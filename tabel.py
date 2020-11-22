import docx
import os
import sqlite3
import sys

from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon

from window1 import Ui_MainWindow


class TestApp(Ui_MainWindow):
    def __init__(self, dialog):
        Ui_MainWindow.__init__(self)
        self.setupUi(dialog)

        self.addButton.clicked.connect(self.addRow)
        self.editButton.clicked.connect(self.editRow)
        self.deleteButton.clicked.connect(self.deleteRow)
        self.saveButton.clicked.connect(self.save)
        self.updateButton.clicked.connect(self.update)

    def addRow(self):
        x = self.object.text()
        y = self.hour.text()
        if y == '':
            error_dialog = QtWidgets.QMessageBox()
            error_dialog.setText("Не заполнено количество часов")
            error_dialog.setWindowTitle("ОШИБКА")
            error_dialog.exec_()
        else:
            # Create a empty row at bottom of table
            numRows = self.tableWidget.rowCount()
            self.tableWidget.insertRow(numRows)
            # Add text to the row
            self.tableWidget.setItem(numRows, 0, QtWidgets.QTableWidgetItem(x))
            self.tableWidget.setItem(numRows, 1, QtWidgets.QTableWidgetItem(y))
            self.tableWidget.resizeColumnsToContents()
            self.object.setText('')
            self.hour.setText('')
            numRows = self.tableWidget.rowCount()
            ylist = []
            for row in range(numRows):
                yitem = self.tableWidget.item(row, 1)
                y = int(yitem.text())
                ylist.append(y)
            ysum = str(sum(ylist))
            self.lcd.display(ysum)

    def editRow(self):
        row = self.tableWidget.currentRow()
        obj = self.tableWidget.item(row, 0).text()
        hou = self.tableWidget.item(row, 1).text()
        self.object.setText(obj)
        self.hour.setText(hou)
        self.tableWidget.removeRow(row)
        numRows = self.tableWidget.rowCount()
        ylist = []
        for row in range(numRows):
            yitem = self.tableWidget.item(row, 1)
            y = int(yitem.text())
            ylist.append(y)
        ysum = str(sum(ylist))
        self.lcd.display(ysum)

    def deleteRow(self):
        row = self.tableWidget.currentRow()
        self.tableWidget.removeRow(row)
        numRows = self.tableWidget.rowCount()
        ylist = []
        for row in range(numRows):
            yitem = self.tableWidget.item(row, 1)
            y = int(yitem.text())
            ylist.append(y)
        ysum = str(sum(ylist))
        self.lcd.display(ysum)

    def update(self):
        cur_mounth = self.mounth.currentText()
        cur_year = self.year.currentText()
        cur_name = self.name.currentText()
        conn = sqlite3.connect('databases/' + cur_mounth + cur_year + '.db')
        c = conn.cursor()
        query = 'SELECT * FROM ' + cur_name
        result = c.execute(query)
        self.tableWidget.setRowCount(0)
        for row_number, row_data in enumerate(result):
            self.tableWidget.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                self.tableWidget.setItem(row_number, column_number, QtWidgets.QTableWidgetItem(str(data)))
                self.tableWidget.resizeColumnsToContents()
        conn.commit()
        c.close()
        conn.close()
        numRows = self.tableWidget.rowCount()
        ylist = []
        for row in range(numRows):
            yitem = self.tableWidget.item(row, 1)
            y = int(yitem.text())
            ylist.append(y)
        ysum = str(sum(ylist))
        self.lcd.display(ysum)

    def save(self):
        cur_mounth = self.mounth.currentText()
        cur_year = self.year.currentText()
        cur_name = self.name.currentText()
        file_path = 'files/' + cur_mounth + cur_year + '.docx'
        if os.path.exists(file_path):
            mydoc = docx.Document('files/' + cur_mounth + cur_year + '.docx')
        else:
            mydoc = docx.Document()
        para = mydoc.add_paragraph(cur_name + ''':
        ''')
        conn = sqlite3.connect('databases/' + cur_mounth + cur_year + '.db')
        c = conn.cursor()
        c.execute("DROP TABLE IF EXISTS " + cur_name)
        c.execute("create table if not exists " + cur_name + " (object text, hours text)")
        for i in range(0, self.tableWidget.rowCount()):
            object = self.tableWidget.item(i, 0).text()
            hours = self.tableWidget.item(i, 1).text()
            c.execute("INSERT into " + cur_name + " VALUES(?, ?)", (object, hours))
            para.add_run(object.upper() + ' ' + '(' + hours + ' н/ч), ')
        conn.commit()
        c.close()
        conn.close()
        mydoc.save('files/' + cur_mounth + cur_year + '.docx')
        self.tableWidget.setRowCount(0)
        self.lcd.display(0)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QMainWindow()
    test_1 = TestApp(dialog)
    dialog.show()
    dialog.setWindowTitle('Tabel v.1.0')
    dialog.setWindowIcon(QIcon('icon.ico'))
    sys.exit(app.exec_())
