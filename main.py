import sqlite3
import sys
import os
import time
import shutil
import threading
import pandas as pd
import xlrd
from xlutils.copy import copy
import csv

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSql import *


class Viewer(QWidget):
    def __init__(self, parent=None):
        super(Viewer, self).__init__(parent)
        self.hlayout = QHBoxLayout()
        self.hlayout1 = QHBoxLayout()
        self.layout = QVBoxLayout()

        self.pici = ''
        self.proname = ''
        self.proname1 = ''
        self.main_path = 'E:/Gene'

        self.query = QSqlQuery()
        self.db = QSqlDatabase.addDatabase("QSQLITE")
        self.resize(1200, 700)
        self.setWindowTitle("基因测序小工具")

        # 第二行
        self.searchEdit = QLineEdit()
        self.searchEdit.setFixedHeight(32)
        self.searchEdit.setStyleSheet("border:1px solid #708090; border-radius:5px; margin-right: 20px")
        font = QFont()
        font.setPixelSize(14)
        self.searchEdit.setFont(font)

        self.searchButton = QPushButton("搜索")
        self.searchEdit.setPlaceholderText("支持样本名称、科室、开单日期模糊搜索")
        self.searchButton.setFixedSize(150, 32)
        self.searchButton.setFont(font)
        self.searchButton.setStyleSheet("color:white; background-color:#708090; font:bold 10pt;margin-right:50px; border-radius:5px")

        self.deleteButton = QPushButton("删除")
        self.deleteButton.setFixedSize(100, 32)
        self.deleteButton.setFont(font)
        self.deleteButton.setStyleSheet("color: white; background-color: 	#708090 ;font: bold 10pt;border-radius:5px")

        # 第一行
        self.label0 = QLabel("测序计划表")
        self.label0.setFixedWidth(135)
        self.label0.setStyleSheet("font-size: 25px; padding-right:0px;")

        self.label = QLabel("样本信息表")
        self.label.setFixedWidth(135)
        self.label.setStyleSheet("font-size: 25px; padding-right:0px;")

        self.addButton = QPushButton("上传")
        self.addButton.setFixedSize(100, 32)
        self.addButton.setFont(font)
        self.addButton.setStyleSheet("color: white; background-color:#708090; font: bold 10pt;border-radius:5px")

        self.addButton_0 = QPushButton("上传")
        self.addButton_0.setFixedSize(100, 32)
        self.addButton_0.setFont(font)
        self.addButton_0.setStyleSheet("color: white; background-color:#708090; font: bold 10pt;border-radius:5px")

        self.startButton = QPushButton("开始")
        self.startButton.setFixedSize(100, 32)
        self.startButton.setFont(font)
        self.startButton.setStyleSheet("color: white; background-color:#708090; font: bold 10pt;border-radius:5px")

        self.processLabel = QLabel("")
        self.processLabel.setFixedWidth(400)
        self.processLabel.setStyleSheet("color:red")
        self.processLabel.setAlignment(Qt.AlignCenter)

        # 空label填充
        self.label3 = QLabel("")
        self.label4 = QLabel("")
        self.label5 = QLabel("")
        self.label3.setFixedHeight(32)
        self.label4.setFixedHeight(1)
        self.label5.setFixedHeight(1)

        self.hlayout.addWidget(self.label0)
        self.hlayout.addWidget(self.addButton_0)
        self.hlayout.addStretch(1)
        self.hlayout.addWidget(self.label)
        self.hlayout.addWidget(self.addButton)
        self.hlayout.addStretch(2)
        self.hlayout.addWidget(self.processLabel)
        self.hlayout.addStretch(1)
        self.hlayout.addWidget(self.startButton)

        self.hlayout1.addWidget(self.searchEdit)
        self.hlayout1.addWidget(self.searchButton)
        self.hlayout1.addWidget(self.deleteButton)

        # 开始按钮状态
        # if os.path.exists('{}/data/data4/Cancer_Project/Blood_tumor/DNA/{}/Report/Result/{}.zip'.format(self.main_path, self.proname, self.proname)):
        #     self.startButton.setEnabled(False)
        #     self.startButton.setStyleSheet("background-color:rgba(0,0,0,0);color:rgba(0,0,0,0)")
        # self.processLabel.setText('报告已生成，程序结束')

        conn = sqlite3.connect('db/sample.db')
        cur = conn.cursor()
        sql = 'select * from people'
        cur.execute(sql)
        data = cur.fetchall()
        row = len(data)
        cur.close()
        conn.close()

        # 表格
        self.vol = 8
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(self.vol + 1)
        self.tableWidget.setRowCount(row)

        # 获取行数、列数
        for i in range(row):
            for j in range(self.vol):
                temp_data = data[i][j]
                if temp_data is None:
                    temp_data = ''
                real_data = QTableWidgetItem(str(temp_data))
                self.tableWidget.setItem(i, j, real_data)
                self.tableWidget.item(i, j).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        # self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tableWidget.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setFrameShape(QFrame.NoFrame)  # 无边框
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 隐藏横向滚动条
        self.tableWidget.setFocusPolicy(Qt.NoFocus)  # 选中无虚线框
        self.tableWidget.horizontalHeader().setHighlightSections(False)  # 表头不塌陷
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)  # 双击可编辑
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)  # 设置选中整行
        self.tableWidget.setSelectionMode(QAbstractItemView.SingleSelection)  # 可以选中单个
        self.tableWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)  # 设置能选中多行
        self.tableWidget.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget.setColumnWidth(1, 90)
        self.tableWidget.setColumnWidth(2, 90)
        self.tableWidget.setColumnWidth(3, 90)
        self.tableWidget.setColumnWidth(4, 90)
        self.tableWidget.setColumnWidth(5, 200)
        self.tableWidget.setColumnWidth(6, 100)
        self.tableWidget.setColumnWidth(7, 100)
        self.tableWidget.setColumnWidth(8, 350)

        self.tableWidget.setStyleSheet("alternate-background-color:#C0C0C0")
        self.tableWidget.horizontalHeader().setStyleSheet("QHeaderView:section{color: white; font: bold 10pt; background-color:#008B8B; border:0px solid;}")
        self.tableWidget.setHorizontalHeaderLabels(['', '样本编号', '样本名称', '年龄', '性别', '开单日期', '科室', '状态 ', '修改/下载'])  # 设置表头信息, '       状态       ', '修改/下载'])  # 设置表头信息
        for k in range(row):
            self.tableWidget.setCellWidget(k, self.vol, self.buttonForRow())  # 在最后一个单元格中加入修改、删除按钮
        for b in range(row):
            self.checkbox = QTableWidgetItem()
            self.checkbox.setCheckState(Qt.Unchecked)
            self.tableWidget.setItem(b, 0, self.checkbox)  # 在第一个单元格加入复选框
        # for m in range(row):
        #     self.tableWidget.item(m, 1).setFlags(Qt.ItemIsEnabled)
        #     self.tableWidget.item(m, 7).setFlags(Qt.ItemIsEnabled)

        # 布局与点击事件
        self.layout.addLayout(self.hlayout)
        self.layout.addWidget(self.label5)  # 空的label
        self.layout.addLayout(self.hlayout1)
        self.layout.addWidget(self.label4)  # 空的label
        self.layout.addWidget(self.tableWidget)
        self.layout.addWidget(self.label3)  # 空的label
        self.setLayout(self.layout)

        self.searchButton.clicked.connect(self.searchButtonClicked)
        self.searchEdit.returnPressed.connect(self.searchButtonClicked)
        self.addButton.clicked.connect(self.addButtonClicked)
        self.addButton_0.clicked.connect(self.planAddButtonClicked)
        self.deleteButton.clicked.connect(self.delButtonClicked)
        self.startButton.clicked.connect(self.startButtonClicked)

    # 按钮
    def buttonForRow(self):
        widget = QWidget()
        # 修改
        button_change = QPushButton('修改')
        button_change.setStyleSheet("color: white; background-color:#008B8B; font: bold 10pt;border-radius:5px")
        button_change.clicked.connect(self.changeButtonClicked)
        button_change.setFixedHeight(28)
        button_change.setMaximumWidth(105)
        button_change.setMinimumWidth(60)

        # word下载
        button_word = QPushButton('WORD')
        button_word.setStyleSheet("color: white; background-color:#008B8B; font: bold 10pt;border-radius:5px")
        button_word.clicked.connect(lambda: self.downloadButtonClicked(button_word.text()))
        button_word.setFixedHeight(28)
        button_word.setMaximumWidth(105)
        button_word.setMinimumWidth(60)

        # excel下载
        button_excel = QPushButton('EXCEL')
        button_excel.setStyleSheet("color: white; background-color:#008B8B; font: bold 10pt;border-radius:5px")
        button_excel.clicked.connect(lambda: self.downloadButtonClicked(button_excel.text()))
        button_excel.setFixedHeight(28)
        button_excel.setMaximumWidth(105)
        button_excel.setMinimumWidth(60)

        # bam下载
        button_bam = QPushButton('BAM')
        button_bam.setStyleSheet("color: white; background-color:#008B8B; font: bold 10pt;border-radius:5px")
        button_bam.clicked.connect(lambda: self.downloadButtonClicked(button_bam.text()))
        button_bam.setFixedHeight(28)
        button_bam.setMaximumWidth(105)
        button_bam.setMinimumWidth(60)

        layout = QHBoxLayout()
        layout.addWidget(button_change)
        layout.addWidget(button_word)
        layout.addWidget(button_excel)
        layout.addWidget(button_bam)
        layout.setContentsMargins(5, 2, 5, 2)
        widget.setLayout(layout)
        return widget

    # 修改
    def changeButtonClicked(self):
        try:
            conn = sqlite3.connect('db/sample.db')
            cur = conn.cursor()
            index = self.tableWidget.selectedItems()
            if len(index) == 0:
                return
            number = index[1].text()
            new_name = index[2].text()
            new_age = index[3].text()
            new_sex = index[4].text()
            new_date = index[5].text()
            new_keshi = index[6].text()

            file1 = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleInfo/CX320.xlsx'
            file2 = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/NextSeq550-CX320-SampleSheet.csv'
            rb = xlrd.open_workbook(file1)
            table = rb.sheets()[0]
            cb = copy(rb)
            s = cb.get_sheet(0)
            n = 0
            while True:
                if table.cell(n, 0).value == number:
                    s.write(n, 1, new_date)
                    s.write(n, 4, new_name)
                    s.write(n, 6, new_sex)
                    s.write(n, 7, new_age)
                    s.write(n, 11, new_keshi)
                    break
                n += 1
            cb.save(file1)

            os.rename(file2, 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv')
            f1 = open('E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv', 'r')
            f2 = open(file2, 'w', newline='')
            fr = csv.reader(f1)
            fw = csv.writer(f2)
            for line in fr:
                if line[0] == number:
                    line[1] = new_name
                fw.writerow(line)
            f1.close()
            f2.close()
            os.remove('E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv')

            sql = "update people set name ='" + new_name + "', age='" + new_age + "', sex ='" + new_sex + "', date = '" + new_date + "', keshi = '" + new_keshi + "' where number = '" + number + "'"
            if cur.execute(sql):
                conn.commit()
                QMessageBox.information(self, "提醒", "修改成功", QMessageBox.Yes, QMessageBox.Yes)
            else:
                QMessageBox.information(self, "提醒", "修改失败", QMessageBox.Yes, QMessageBox.Yes)
        except Exception as e:
            print('异常信息为:', e)

    # word/excel/bam下载
    def downloadButtonClicked(self, s_btn):
        try:
            if s_btn == 'WORD':
                s = self.sender()
                row = self.tableWidget.indexAt(s.parent().pos()).row()  # 获按钮行号
                number = self.tableWidget.item(row, 1).text()  # 获取样本编号

                dir_path = QFileDialog.getSaveFileName(self, '', '{}'.format(number), '文件(*.docx)')
                real_path = dir_path[0]
                if real_path == "":
                    QMessageBox.information(self, "提醒", "请选择下载路径", QMessageBox.Yes, QMessageBox.Yes)
                else:
                    conn = sqlite3.connect('db/sample.db')
                    cur = conn.cursor()
                    sql = "select word from people where number = '" + number + "' "
                    cur.execute(sql)
                    real_sql_path = cur.fetchall()[0][0]
                    if real_sql_path is None:
                        QMessageBox.information(self, "提醒", "word未上传", QMessageBox.Yes, QMessageBox.Yes)
                        return
                    shutil.copy(real_sql_path, real_path)
                    QMessageBox.information(self, "提醒", "word下载成功", QMessageBox.Yes, QMessageBox.Yes)
            elif s_btn == 'EXCEL':
                s = self.sender()
                row = self.tableWidget.indexAt(s.parent().pos()).row()  # 获按钮行号
                number = self.tableWidget.item(row, 1).text()  # 获取样本编号

                dir_path = QFileDialog.getSaveFileName(self, '', '{}'.format(number), '文件(*.xlsx)')
                real_path = dir_path[0]
                if real_path == "":
                    QMessageBox.information(self, "提醒", "请选择下载路径", QMessageBox.Yes, QMessageBox.Yes)
                else:
                    conn = sqlite3.connect('db/sample.db')
                    cur = conn.cursor()
                    sql = "select excel from people where number = '" + number + "' "
                    cur.execute(sql)
                    real_sql_path = cur.fetchall()[0][0]
                    if real_sql_path is None:
                        QMessageBox.information(self, "提醒", "excel未上传", QMessageBox.Yes, QMessageBox.Yes)
                        return
                    shutil.copy(real_sql_path, real_path)
                    QMessageBox.information(self, "提醒", "excel下载成功", QMessageBox.Yes, QMessageBox.Yes)
            else:
                s = self.sender()
                row = self.tableWidget.indexAt(s.parent().pos()).row()  # 获按钮行号
                number = self.tableWidget.item(row, 1).text()  # 获取样本编号

                dir_path = QFileDialog.getSaveFileName(self, '', 'bam')
                real_path = os.path.split(dir_path[0])[0]
                if real_path == "":
                    QMessageBox.information(self, "提醒", "请选择下载路径", QMessageBox.Yes, QMessageBox.Yes)
                else:
                    conn = sqlite3.connect('db/sample.db')
                    cur = conn.cursor()
                    sql = "select bam from people where number = '" + number + "' "
                    cur.execute(sql)
                    real_sql_path = cur.fetchall()[0][0]
                    if real_sql_path is None:
                        QMessageBox.information(self, "提醒", "bam数据未上传", QMessageBox.Yes, QMessageBox.Yes)
                        return
                    bam = real_sql_path + '/{}.merge.bam'.format(number)
                    bai = real_sql_path + '/{}.merge.bam.bai'.format(number)
                    shutil.copy2(bam, real_path)
                    shutil.copy2(bai, real_path)
                    QMessageBox.information(self, "提醒", "bam数据下载成功", QMessageBox.Yes, QMessageBox.Yes)
        except Exception as e:
            print('异常', e)
            QMessageBox.information(self, "提醒", "下载失败，请选择正确的文件格式", QMessageBox.Yes, QMessageBox.Yes)

    # 查询
    def searchButtonClicked(self):
        try:
            txt = self.searchEdit.text().strip()
            if u'\u4e00' <= txt <= u'\u9fa5':
                conn = sqlite3.connect('db/sample.db')
                cur = conn.cursor()
                sql = "select * from people where name like '%" + txt + "%' or keshi like '%" + txt + "%'"
                cur.execute(sql)
                data_x = cur.fetchall()
                self.tableWidget.clearContents()
                row_4 = len(data_x)
                # 查询到的更新到表格当中
                for i_x in range(row_4):
                    for j_y in range(self.vol):
                        temp_data_1 = data_x[i_x][j_y]
                        if temp_data_1 is None:
                            temp_data_1 = ''
                        data_1 = QTableWidgetItem(str(temp_data_1))
                        self.tableWidget.setItem(i_x, j_y, data_1)
                        self.tableWidget.item(i_x, j_y).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        self.tableWidget.setRowCount(row_4)
                for k in range(row_4):
                    self.tableWidget.setCellWidget(k, self.vol, self.buttonForRow())
                    self.checkbox = QTableWidgetItem()
                    self.checkbox.setCheckState(Qt.Unchecked)
                    self.tableWidget.setItem(k, 0, self.checkbox)
            elif txt == "":
                conn = sqlite3.connect('db/sample.db')
                cur = conn.cursor()
                sql = 'select * from people'
                cur.execute(sql)
                data = cur.fetchall()
                row = len(data)
                for i in range(row):
                    for j in range(self.vol):
                        temp_data = data[i][j]
                        if temp_data is None:
                            temp_data = ''
                        real_data = QTableWidgetItem(str(temp_data))
                        self.tableWidget.setItem(i, j, real_data)
                        self.tableWidget.setRowCount(row)
                        self.tableWidget.item(i, j).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                for r in range(row):
                    self.tableWidget.setCellWidget(r, self.vol, self.buttonForRow())
                    self.checkbox = QTableWidgetItem()
                    self.checkbox.setCheckState(Qt.Unchecked)
                    self.tableWidget.setItem(r, 0, self.checkbox)
            else:
                conn = sqlite3.connect('db/sample.db')
                cur = conn.cursor()
                sql = "select * from people where date like '%" + txt + "%' "
                cur.execute(sql)
                data_y = cur.fetchall()
                self.tableWidget.clearContents()
                row_5 = len(data_y)
                for i_x_1 in range(row_5):
                    for j_y_1 in range(self.vol):
                        temp_data_2 = data_y[i_x_1][j_y_1]
                        if temp_data_2 is None:
                            temp_data_2 = ''
                        data_2 = QTableWidgetItem(str(temp_data_2))
                        self.tableWidget.setItem(i_x_1, j_y_1, data_2)
                        self.tableWidget.setRowCount(row_5)
                        self.tableWidget.item(i_x_1, j_y_1).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                for k in range(row_5):
                    self.tableWidget.setCellWidget(k, self.vol, self.buttonForRow())
                    self.checkbox = QTableWidgetItem()
                    self.checkbox.setCheckState(Qt.Unchecked)
                    self.tableWidget.setItem(k, 0, self.checkbox)
        except Exception as e:
            print('异常', e)

    # 上传测序计划
    def planAddButtonClicked(self):
        try:
            file_name_0 = QFileDialog.getOpenFileName(self, 'open file', '', 'csv文件(*.csv)')
            f_name_0 = file_name_0[0]
            if f_name_0 == "":
                QMessageBox.information(self, "提醒", "请选择要上传的文件", QMessageBox.Yes, QMessageBox.Yes)
            else:
                shutil.copy2(f_name_0, 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet')
                QMessageBox.information(self, "提醒", "上传成功", QMessageBox.Yes, QMessageBox.Yes)
        except Exception as e:
            print('异常', e)
            QMessageBox.information(self, "提醒", "上传失败，请检查文件", QMessageBox.Yes, QMessageBox.Yes)

    # 上传样本信息
    def addButtonClicked(self):
        try:
            file_name = QFileDialog.getOpenFileName(self, 'open file', '', 'excel文件(*.xls *.xlsx)')
            f_name = file_name[0]
            if f_name == "":
                QMessageBox.information(self, "提醒", "请选择要上传的文件", QMessageBox.Yes, QMessageBox.Yes)
            else:
                df = pd.DataFrame(pd.read_excel(f_name))
                df2 = df.rename(
                    columns={"Unnamed: 0": "number", "开单日期": "date", "签收日期": "date_end", "编号（S）": "simpleid", "姓名": "name", "标本类型": "biaoben", "性别": "sex", "年龄": "age", "床号": "chuanghao",
                             "病案号": "binanhao", "医院": "yiyuan", "科室": "keshi", "开单大夫": "daifu", "检测项目": "xiangmu", "诊断": "zhenduan"})
                conn = sqlite3.connect('db/sample.db')
                df2.to_sql('people', con=conn, if_exists='append', index=False)  # 修改为不可重复上传

                cur = conn.cursor()
                sql = "update people set state = '待启动'"
                cur.execute(sql)
                conn.commit()

                shutil.copy2(f_name, 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleInfo')

                self.pici = os.path.splitext(os.path.basename(f_name))[0]
                times = time.strftime('%Y%m%d', time.localtime(time.time()))
                self.proname = '.'.join([self.pici, times])
                self.proname1 = '.'.join([self.pici, times, times])

                QMessageBox.information(self, "提醒", "上传成功", QMessageBox.Yes, QMessageBox.Yes)
        except Exception as e:
            print('异常', e)
            QMessageBox.information(self, "提醒", "上传失败,请检查文件", QMessageBox.Yes, QMessageBox.Yes)

    # 删除
    def delButtonClicked(self):
        try:
            row = self.tableWidget.rowCount()
            reply = QMessageBox.information(self, '确认', '确定删除数据?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                for i in range(row):
                    if self.tableWidget.item(i, 0).checkState() == Qt.Checked:
                        number = self.tableWidget.item(i, 1).text()
                        conn = sqlite3.connect('db/sample.db')
                        cur = conn.cursor()
                        sql = "update people set state = '已删除' where number = '"+number+"'"
                        cur.execute(sql)
                        conn.commit()
                        self.tableWidget.setItem(i, 7, QTableWidgetItem('已删除'))

                        file1 = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleInfo/CX320.xlsx'
                        file2 = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/NextSeq550-CX320-SampleSheet.csv'
                        df = pd.DataFrame(pd.read_excel(file1, index_col=0))
                        n = 0
                        while True:
                            if df.values[n][2] == number:
                                df = df.drop(number)
                                break
                            n += 1
                        df.to_excel(file1)

                        os.rename(file2, 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv')
                        f1 = open('E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv', 'r')
                        f2 = open(file2, 'w', newline='')
                        fr = csv.reader(f1)
                        fw = csv.writer(f2)
                        for line in fr:
                            if line[0] == number:
                                continue
                            fw.writerow(line)
                        f1.close()
                        f2.close()
                        os.remove('E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/Config/SampleSheet/b.csv')
        except Exception as e:
            print('异常', e)

    def brush(self):
        QApplication.processEvents()

    def run_docker(self):
        start_cmd = "docker start hm00"
        os.system(start_cmd)

        exec_cmd = "docker exec -i hm00 bash /data3/tmp/cx.sh"
        os.system(exec_cmd)

    # 开始
    def startButtonClicked(self):
        try:
            if not (os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\Config\SampleInfo\CX320.xlsx') and os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\Config\SampleSheet\NextSeq550-CX320-SampleSheet.csv')):
                QMessageBox.information(self, "提醒", "测序表或样本表未上传！", QMessageBox.Yes, QMessageBox.Yes)
                return
            self.startButton.setEnabled(False)

            num = self.tableWidget.rowCount()
            status = [0, 0, 0, 0, 0]
            num_list = []
            state = []
            for i in range(num):
                self.tableWidget.setItem(i, 7, QTableWidgetItem('1/8处理中'))
                state.append([0, 0, 0, 0, 0, 0])
                num_list.append(self.tableWidget.item(i, 1).text())

            t = threading.Thread(target=self.run_docker)
            t.start()

            while True:
                for n in range(num):
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\Rawdata2Fastq\{}\{}\{}_L1_1.fq.gz'.format(self.proname, num_list[n], num_list[n])) and state[n][0] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('2/8已拆分'))
                        state[n][0] = 1
                        self.brush()
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\QC\{}\{}_{}_{}_L1_1.clean.fq.gz'.format(self.proname, num_list[n], num_list[n], num_list[n], self.proname)) and state[n][1] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('3/8已清洗'))
                        state[n][1] = 1
                        self.brush()
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Mapping\{}\{}_{}_{}_L1-1.sam'.format(self.proname, num_list[n], num_list[n], num_list[n], self.proname)) and state[n][2] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('4/8已比对'))
                        state[n][2] = 1
                        self.brush()
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Mapping\{}\{}.final.bam'.format(self.proname, num_list[n], num_list[n])) and state[n][3] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('5/8已去重'))
                        state[n][3] = 1
                        self.brush()
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Result\{}\VarDict\{}.vars.vcf'.format(self.proname, num_list[n], num_list[n])) and state[n][4] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('6/8已检测'))
                        state[n][4] = 1
                        self.brush()
                    if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Result\Annotate\Sample_{}.out.xlsx'.format(self.proname, num_list[n])) and state[n][5] == 0:
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('7/8已注释'))
                        state[n][5] = 1
                        self.brush()

                if status[4] == 0:
                    self.processLabel.setText('拆分数据中...')
                    status[4] = 1
                    self.brush()
                if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\qc_mapping_somatic.job'.format(self.proname)) and status[0] == 0:
                    self.processLabel.setText('job文件已生成，运行中...')
                    status[0] = 1
                    self.brush()
                if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Report\qc\qc_complete.txt'.format(self.proname)) and status[1] == 0:
                    self.processLabel.setText('质控已完成，正在进行比对去重...')
                    status[1] = 1
                    self.brush()
                if os.path.exists(
                        r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Report\mapping\qc_mapping_summary.xls'.format(self.proname)) and status[2] == 0:
                    self.processLabel.setText('比对去重已完成，正在进行变异检测...')
                    status[2] = 1
                    self.brush()
                if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Blood_tumor\DNA\{}\Report\variation\variation_complete.txt'.format(self.proname)) and status[3] == 0:
                    self.processLabel.setText('变异检测已完成，正在生成报告...')
                    status[3] = 1
                    self.brush()
                if os.path.exists(r'E:\Gene\data\data4\Cancer_Project\Project_Stat\Blood_Project_Stat\{}\Project_data_statistics.txt'.format(self.proname1)):
                    self.processLabel.setText('报告已生成，程序结束')
                    conn = sqlite3.connect('db/sample.db')
                    for n in range(num):
                        self.tableWidget.setItem(n, 7, QTableWidgetItem('8/8已完成'))
                        cur = conn.cursor()
                        sql = "select date_end,name from people where number = '"+num_list[n]+"'"
                        cur.execute(sql)
                        data = cur.fetchall()
                        date = data[0][0]
                        name = data[0][1]
                        word = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/{}/Result/Annotate/{}-{}-{}.docx'.format(self.proname, num_list[n], name, date)
                        excel = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/{}/Result/Annotate/Sample_{}.out.xlsx'.format(self.proname, num_list[n])
                        bam = 'E:/Gene/data/data4/Cancer_Project/Blood_tumor/DNA/{}/Mapping/{}'.format(self.proname, num_list[n])
                        cur1 = conn.cursor()
                        sql1 = "update people set word = '"+word+"', excel = '"+excel+"', bam = '"+bam+"', state = '已完成' where number = '"+num_list[n]+"'"
                        cur1.execute(sql1)
                    conn.commit()
                    self.brush()
                    break
                self.brush()
        except Exception as e:
            print('异常', e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = Viewer()
    mainWindow.show()
    sys.exit(app.exec_())
