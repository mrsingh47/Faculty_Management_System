from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import sqlite3
import datetime
from xlrd import *
from xlsxwriter import *
from PyQt5.uic import loadUiType

ui, _ = loadUiType('main.ui')


class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_Buttons()
        self.Handle_ui_changes()
        self.showAllFaculty()
        self.showCourse()
        self.showDepartment()
        self.showDesignation()
        self.showCourseCombobox()
        self.showDeptCombobox()
        self.showDesigCombobox()
        self.Theme()

    def Handle_ui_changes(self):
        self.tabWidget.tabBar().setVisible(False)

    def Handle_Buttons(self):
        self.pushButton.clicked.connect(self.openAddFaculty)
        self.pushButton_2.clicked.connect(self.openUpdateFaculty)
        self.pushButton_3.clicked.connect(self.openCourseFaculty)
        self.pushButton_8.clicked.connect(self.openDeptFaculty)
        self.pushButton_11.clicked.connect(self.openDesignationFaculty)
        self.pushButton_4.clicked.connect(self.addFaculty)
        self.pushButton_9.clicked.connect(self.addCourse)
        self.pushButton_10.clicked.connect(self.addDepartment)
        self.pushButton_12.clicked.connect(self.addDesignation)
        self.pushButton_13.clicked.connect(self.openshowAllFaculty)
        self.pushButton_7.clicked.connect(self.searchFaculty)
        self.pushButton_5.clicked.connect(self.updateFaculty)
        self.pushButton_6.clicked.connect(self.deleteFaculty)
        self.pushButton_14.clicked.connect(self.ExportData)

    ################################
    ######## opening tabs ##########
    def openshowAllFaculty(self):
        self.tabWidget.setCurrentIndex(5)

    def openAddFaculty(self):
        self.tabWidget.setCurrentIndex(0)

    def openUpdateFaculty(self):
        self.tabWidget.setCurrentIndex(1)

    def openCourseFaculty(self):
        self.tabWidget.setCurrentIndex(2)

    def openDeptFaculty(self):
        self.tabWidget.setCurrentIndex(3)

    def openDesignationFaculty(self):
        self.tabWidget.setCurrentIndex(4)

    ##################################
    ######## Faculty Manage ##########

    def addFaculty(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        Fac_no = self.lineEdit.text()
        name = self.lineEdit_2.text()
        course = self.comboBox.currentText()
        department = self.comboBox_2.currentText()
        designation = self.comboBox_3.currentText()
        phone = self.lineEdit_5.text()
        salary = self.lineEdit_3.text()
        doj = self.lineEdit_4.text()
        if Fac_no != '' and name != '' and phone != '':
            self.cur.execute(
                "insert into faculty(Faculty_Number,Name,Department,Course,Designation,salary,Phone_Number,Date_of_Joining) values('{}','{}','{}','{}','{}','{}','{}','{}')".format(Fac_no, name, department, course, designation, salary, phone, doj))
            self.con.commit()
            self.statusBar().showMessage('New Faculty Added Successfully..')
            self.con.close()

            self.lineEdit.setText('')
            self.lineEdit_2.setText('')
            self.comboBox.setCurrentIndex(0)
            self.comboBox_2.setCurrentIndex(0)
            self.comboBox_3.setCurrentIndex(0)
            self.lineEdit_5.setText('')
            self.lineEdit_3.setText('')
            self.lineEdit_4.setText('')
            self.showAllFaculty()
        else:
            self.statusBar().showMessage('Some Field are Empty.')

    def searchFaculty(self):

        global state
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        facNo = self.lineEdit_11.text()
        self.cur.execute(
            "select * from faculty where Faculty_Number='{}'".format(facNo))
        data = self.cur.fetchone()
        self.groupBox.setEnabled(False)
        if data != None:
            state = 1
        else:
            self.statusBar().showMessage('Data Not Found!!')
        if state == 1:
            self.groupBox.setEnabled(True)
            self.lineEdit_9.setText(data[1])
            self.lineEdit_10.setText(data[2])
            self.lineEdit_8.setText(data[7])
            self.lineEdit_6.setText(data[6])
            self.lineEdit_7.setText(data[8])
            self.comboBox_6.setCurrentText(data[4])
            self.comboBox_4.setCurrentText(data[3])
            self.comboBox_5.setCurrentText(data[5])
            state = 0

    def updateFaculty(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        Fac_no = self.lineEdit_9.text()
        name = self.lineEdit_10.text()
        course = self.comboBox_6.currentText()
        department = self.comboBox_4.currentText()
        designation = self.comboBox_5.currentText()
        phone = self.lineEdit_8.text()
        salary = self.lineEdit_6.text()
        doj = self.lineEdit_7.text()

        SearchFacultyNo = self.lineEdit_11.text()

        self.cur.execute("update faculty set Faculty_Number='{}',Name='{}',Department='{}',Course='{}',Designation='{}',salary='{}',Phone_Number='{}',Date_of_Joining='{}' where Faculty_Number='{}'".format(
            Fac_no, name, department, course, designation, salary, phone, doj, SearchFacultyNo))
        self.con.commit()
        self.statusBar().showMessage('Faculty Data Updated Successfully..')
        self.con.close()

        self.groupBox.setEnabled(False)
        self.lineEdit_9.setText('')
        self.lineEdit_10.setText('')
        self.lineEdit_8.setText('')
        self.lineEdit_6.setText('')
        self.lineEdit_7.setText('')
        self.comboBox_6.setCurrentText('')
        self.comboBox_4.setCurrentText('')
        self.comboBox_5.setCurrentText('')
        self.showAllFaculty()

    def deleteFaculty(self):
        SearchFacultyNo = self.lineEdit_11.text()
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        warn = QMessageBox.warning(self, 'Delete Faculty Record',
                                   "Are you sure u wanna Delete the Record of this Person", QMessageBox.Yes | QMessageBox.No)
        if warn == QMessageBox.Yes:
            self.cur.execute(
                "delete from faculty where Faculty_Number='{}'".format(SearchFacultyNo))
            self.con.commit()
            self.statusBar().showMessage('Faculty Data Deleted Successfully..')
            self.con.close()
            self.groupBox.setEnabled(False)
            self.lineEdit_9.setText('')
            self.lineEdit_10.setText('')
            self.lineEdit_8.setText('')
            self.lineEdit_6.setText('')
            self.lineEdit_7.setText('')
            self.comboBox_6.setCurrentText('')
            self.comboBox_4.setCurrentText('')
            self.comboBox_5.setCurrentText('')
            self.showAllFaculty()

    ####################################################
    ############ Add New Course/Department #############

    def addCourse(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        course = self.lineEdit_12.text()

        self.cur.execute(
            "insert into course(Course) values('{}');".format(course))
        self.con.commit()
        self.statusBar().showMessage('New Course Added Successfully')
        self.lineEdit_12.setText('')
        self.con.close()
        self.showCourse()
        self.showCourseCombobox()

    def addDepartment(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        dept = self.lineEdit_13.text()

        self.cur.execute(
            "insert into department(Department_Name) values('{}');".format(dept))
        self.con.commit()
        self.statusBar().showMessage('New Department Added Successfully')
        self.lineEdit_13.setText('')
        self.con.close()
        self.showDepartment()
        self.showDeptCombobox()

    def addDesignation(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        desig = self.lineEdit_14.text()

        self.cur.execute(
            "insert into designation(Designation) values('{}');".format(desig))
        self.con.commit()
        self.statusBar().showMessage('New Designation Added Successfully')
        self.lineEdit_14.setText('')
        self.con.close()
        self.showDesignation()
        self.showDesigCombobox()

    #################################################
    ############ Show Course/Department #############
    def showCourse(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute("select Course from course")
        data = self.cur.fetchall()
        if data:
            # Clear empty rows each time showing (add new course)
            self.tableWidget.setRowCount(0)
            self.tableWidget.insertRow(0)
            for row, i in enumerate(data):
                for column, j in enumerate(i):
                    self.tableWidget.setItem(
                        row, column, QTableWidgetItem(str(j)))
                    column += 1
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
        self.con.close()

    def showDepartment(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute("select Department_Name from department")
        data = self.cur.fetchall()
        if data:
            # Clear empty rows each time showing (add new course)
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, i in enumerate(data):
                for column, j in enumerate(i):
                    self.tableWidget_2.setItem(
                        row, column, QTableWidgetItem(str(j)))
                    column += 1
                rowPosition = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(rowPosition)
        self.con.close()

    def showDesignation(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute("select Designation from designation")
        data = self.cur.fetchall()
        if data:
            # Clear empty rows each time showing (add new course)
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row, i in enumerate(data):
                for column, j in enumerate(i):
                    self.tableWidget_3.setItem(
                        row, column, QTableWidgetItem(str(j)))
                    column += 1
                rowPosition = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(rowPosition)
        self.con.close()

    #####################################################################
    ############### show courses/dept/desig in comboBox #################
    def showCourseCombobox(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute("select Course from course order by Course")
        data = self.cur.fetchall()
        self.comboBox.clear()
        self.comboBox_6.clear()
        for course in data:
            self.comboBox.addItem(course[0])
            self.comboBox_6.addItem(course[0])
        self.con.close()

    def showDeptCombobox(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute(
            "select Department_Name from department order by Department_Name")
        data = self.cur.fetchall()
        self.comboBox_2.clear()
        self.comboBox_4.clear()
        for dept in data:
            self.comboBox_2.addItem(dept[0])
            self.comboBox_4.addItem(dept[0])
        self.con.close()

    def showDesigCombobox(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()
        self.cur.execute(
            "select Designation from designation order by Designation")
        data = self.cur.fetchall()
        self.comboBox_3.clear()
        self.comboBox_5.clear()
        for desig in data:
            self.comboBox_3.addItem(desig[0])
            self.comboBox_5.addItem(desig[0])
        self.con.close()

    ##################################################
    ############### Show All Faculty #################

    def showAllFaculty(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        self.cur.execute(
            "select Faculty_Number,Name,Department,Course,Designation,salary,Phone_Number,Date_of_Joining from faculty")
        data = self.cur.fetchall()
        self.tableWidget_4.insertRow(0)
        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)
        for row, i in enumerate(data):
            for column, j in enumerate(i):
                self.tableWidget_4.setItem(
                    row, column, QTableWidgetItem(str(j)))
                column += 1
            rowPosition = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(rowPosition)
        self.con.close()

    ###############################
    ########### Theme #############

    def Theme(self):
        style = open('themes/dark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    #############################################
    ############### Export Data #################
    def ExportData(self):
        self.con = sqlite3.connect('faculty.db')
        self.cur = self.con.cursor()

        self.cur.execute(
            "select Faculty_Number,Name,Department,Course,Designation,salary,Phone_Number,Date_of_Joining from faculty")
        data = self.cur.fetchall()

        currentTime = str(datetime.datetime.now().time())[:-7]
        a = "_".join(currentTime.split(':'))
        wb = Workbook('All_Faculty_'+a+'.xlsx')
        sheet = wb.add_worksheet()
        sheet.write(0, 0, 'Faculty Number')
        sheet.write(0, 1, 'Name')
        sheet.write(0, 2, 'Department')
        sheet.write(0, 3, 'Course')
        sheet.write(0, 4, 'Designation')
        sheet.write(0, 5, 'Salary')
        sheet.write(0, 6, 'Phone_Number')
        sheet.write(0, 7, 'Date_of_Joining')

        rowNumber = 1
        for row in data:
            columnNumber = 0
            for item in row:
                sheet.write(rowNumber, columnNumber, str(item))
                columnNumber += 1
            rowNumber += 1
        wb.close()
        self.con.close()
        self.statusBar().showMessage('Exported Successfully..')


def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()


global state
state = 0
if __name__ == "__main__":
    main()
