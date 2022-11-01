from PyQt5.QtWidgets import *
import xml.etree.ElementTree as ET
import pymysql
import sys
import csv
import json


class DB_Utils:

    def queryExecutor(self, db, sql, params):
        conn = pymysql.connect(host='localhost', user='guest', password='bemyguest', db='classicmodels', charset='utf8')

        try:
            # dictionary based cursor
            with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                cursor.execute(sql, params)
                tuples = cursor.fetchall()
                return tuples
        except Exception as e:
            print(e)
            print(type(e))
        finally:
            conn.close()


class DB_Queries:

    def selectCustomerName(self):
        sql = "SELECT name, customerId FROM customers order by name asc"
        params = ()

        util = DB_Utils()
        tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return tuples

    def selectCustomerCountry(self):
        sql = "SELECT distinct country FROM customers order by country asc"
        params = ()

        util = DB_Utils()
        tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return tuples

    def selectCustomerCity(self):
        sql = "SELECT DISTINCT city FROM customers order by city asc"
        params = ()

        util = DB_Utils()
        tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return tuples

    def updateCustomerCity(self, country):
        if country == "ALL":
            sql = "SELECT DISTINCT city FROM customers order by city ASC"
            params = ()
            util = DB_Utils()
            tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        else :
            sql = "SELECT DISTINCT city FROM customers where country = %s order by city ASC"
            params = (country)
            util = DB_Utils()
            tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return tuples

    def selectDatabyKey(self, customerName, country, city):
        if customerName == 'ALL' and country == 'ALL' and city == 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) "
            params = ()
        if customerName != 'ALL' and country == 'ALL' and city == 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) WHERE name = %s "
            params = (customerName)
        if customerName == 'ALL' and country != 'ALL' and city == 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) WHERE country = %s "
            params = (country)
        if customerName == 'ALL' and country == 'ALL' and city != 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) WHERE city = %s "
            params = (city)
        if customerName == 'ALL' and country != 'ALL' and city != 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) WHERE country = %s and city = %s "
            params = (country, city)
        if customerName != 'ALL' and country != 'ALL' and city == 'ALL':
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                    "orders LEFT JOIN customers USING(customerID) WHERE name = %s and country = %s "
            params = (customerName, country)
        if customerName != 'ALL' and country == 'ALL' and city != 'ALL':
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                    "orders LEFT JOIN customers USING(customerID) WHERE name = %s and city = %s "
            params = (customerName, city)
        if customerName != 'ALL' and country != 'ALL' and city != 'ALL' :
            sql = "SELECT orderNo, orderDate, requiredDate, shippedDate, status, name as customer, comments FROM " \
                  "orders LEFT JOIN customers USING(customerID) WHERE name = %s and country = %s and city = %s "
            params = (customerName, country, city)
        util = DB_Utils()
        tuples = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return tuples

    def selectDetailDatabyKey(self, orderNo):
        sql = "SELECT orderLineNo, productCode, name as productName, quantity, convert(priceEach, char) as priceEach, convert(quantity * priceEach, char) as 상품주문액" \
              " FROM orderDetails LEFT JOIN products USING(productCode) WHERE orderNo = %s"
        params = (orderNo)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows


class SubWindow(QDialog, QWidget) :
    def __init__(self,orderNo):
        super().__init__()
        self.setupUI(orderNo)
        self.show()
        self.fileType = "csv"

    def setupUI(self,orderNo):

        self.setWindowTitle("SubWindow")
        self.setGeometry(100, 100, 1200, 600)

        # 주문 상세 위젯
        self.subLabel = QLabel("주문 상세 내역", self)
        self.orderNumberLabel = QLabel("주문번호: ",self)
        self.orderNumberPrintLabel = QLabel()
        self.orderNumberPrintLabel.setText(orderNo)
        self.orderCountLabel = QLabel("상품개수: ",self)
        self.orderCountPrintLabel = QLabel()
        self.orderCountPrintLabel.setText('0')
        self.orderSumLabel = QLabel("주문액: ",self)
        self.orderSumPrintLabel = QLabel()
        self.orderSumPrintLabel.setText('0')
        self.btnSave = QPushButton("저장")

        # 파일 출력 라디오버튼
        self.exportLabel = QLabel("파일 출력",self)
        self.csvBtn = QRadioButton("CSV", self)
        self.csvBtn.setChecked(True)
        self.csvBtn.clicked.connect(self.fileRadioBtn_Clicked)

        self.jsonBtn = QRadioButton("JSON", self)
        self.jsonBtn.clicked.connect(self.fileRadioBtn_Clicked)

        self.xmlBtn = QRadioButton("XML", self)
        self.xmlBtn.clicked.connect(self.fileRadioBtn_Clicked)

        #테이블
        self.tableLabel = QLabel("주문에 포함된 상품 리스트")
        self.tableWidget = QTableWidget(self)
        self.tableWidget.resize(800, 600)

        query = DB_Queries()
        rows = query.selectDetailDatabyKey(orderNo)

        rows.sort(key=lambda x: x['orderLineNo'])
        self.btnSave.clicked.connect(lambda: self.btnSave_Clicked(orderNo, rows))
        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(len(rows))
        self.tableWidget.setColumnCount(len(rows[0]))
        columnNames = list(rows[0].keys())
        self.tableWidget.setHorizontalHeaderLabels(columnNames)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.sum = 0
        for rowIDX, value in enumerate(rows):
            self.sum += float(value['상품주문액'])

            for columnIDX, (k, v) in enumerate(value.items()):
                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.tableWidget.setItem(rowIDX, columnIDX, item)
        self.sum = round(self.sum, 2)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

        self.orderCountPrintLabel.setText("%s" % (len(rows)))
        self.orderSumPrintLabel.setText("%s" % self.sum)

        # 레이아웃
        self.subLayout = QHBoxLayout()
        self.subLayout.addWidget(self.subLabel)

        self.subLayout_1 = QHBoxLayout()
        self.subLayout_1.addWidget(self.orderNumberLabel)
        self.subLayout_1.addWidget(self.orderNumberPrintLabel)
        self.subLayout_1.addWidget(self.orderCountLabel)
        self.subLayout_1.addWidget(self.orderCountPrintLabel)
        self.subLayout_1.addWidget(self.orderSumLabel)
        self.subLayout_1.addWidget(self.orderSumPrintLabel)

        self.subLayout_2 = QHBoxLayout()
        self.subLayout_2.addWidget(self.tableLabel)

        self.subLayout_3 = QHBoxLayout()
        self.subLayout_3.addWidget(self.tableWidget)

        self.subLayout_4 = QHBoxLayout()
        self.subLayout_4.addWidget(self.exportLabel)

        self.subLayout_5 = QHBoxLayout()
        self.subLayout_5.addWidget(self.csvBtn)
        self.subLayout_5.addWidget(self.jsonBtn)
        self.subLayout_5.addWidget(self.xmlBtn)

        self.subLayout_6 = QHBoxLayout()
        self.subLayout_6.addWidget(self.btnSave)

        self.secondLayout = QVBoxLayout()
        self.secondLayout.addLayout(self.subLayout)
        self.secondLayout.addLayout(self.subLayout_1)
        self.secondLayout.addLayout(self.subLayout_2)
        self.secondLayout.addLayout(self.subLayout_3)
        self.secondLayout.addLayout(self.subLayout_4)
        self.secondLayout.addLayout(self.subLayout_5)
        self.secondLayout.addLayout(self.subLayout_6)

        self.setLayout(self.secondLayout)

    # SubWindow 메소드
    def fileRadioBtn_Clicked(self):
        self.fileType = "csv"
        if self.csvBtn.isChecked():
            self.fileType = "csv"
        elif self.jsonBtn.isChecked():
            self.fileType = "json"
        elif self.xmlBtn.isChecked():
            self.fileType = "xml"

    def btnSave_Clicked(self, orderNo,rows):
        if self.csvBtn.isChecked():
            self.writeCSV(orderNo,rows)
        elif self.jsonBtn.isChecked():
            self.writeJSON(orderNo,rows)
        elif self.xmlBtn.isChecked():
            self.writeXML(orderNo,rows)

    def writeCSV(self, orderNo,rows):
        with open(f'./{orderNo}.csv', 'w', encoding='utf-8', newline='') as f:
            wr = csv.writer(f)
            columnNames = list(rows[0].keys())
            wr.writerow(columnNames)
            for row in rows:
                orderlist = list(row.values())
                wr.writerow(orderlist)

    def writeJSON(self, orderNo, rows):
        newDict = dict(orderNo=rows)
        with open(f'./{orderNo}.json', 'w', encoding='utf-8') as f:
            json.dump(newDict, f, indent=4, ensure_ascii=False)

    def writeXML(self, orderNo, rows):
        newDict = dict(orderNo=rows)

        tableName = list(newDict.keys())[0]
        tableRows = list(newDict.values())[0]

        rootElement = ET.Element('Table')
        rootElement.attrib['name'] = tableName

        for row in tableRows:
            rowElement = ET.Element('Row')
            rootElement.append(rowElement)

            for columnName in list(row.keys()):
                if row[columnName] == None:
                    rowElement.attrib[columnName] = ''
                elif type(row[columnName]) in (int, float):
                    rowElement.attrib[columnName] = str(row[columnName])
                else:
                    rowElement.attrib[columnName] = row[columnName]

        ET.ElementTree(rootElement).write(f'./{orderNo}.xml', encoding='utf-8', xml_declaration=True)
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUI()
        self.customerName = "ALL"
        self.country = "ALL"
        self.city = "ALL"

    def setupUI(self):
        # 윈도우 설정
        self.setWindowTitle("MainWindow")
        self.setGeometry(100, 100, 1440, 900)

        self.mainLabel = QLabel("주문검색", self)
        # 고객
        self.customerNameLabel = QLabel("고객 :", self)
        self.customerNameComboBox = QComboBox(self)

        query = DB_Queries()
        rows = query.selectCustomerName()
        columnName = list(rows[0].keys())[0]
        items = [row[columnName] for row in rows]
        items.sort()
        self.customerNameComboBox.addItem("ALL")
        self.customerNameComboBox.addItems(items)
        self.customerNameComboBox.activated.connect(self.customerNameComboBox_Activated)
        self.customerNameComboBox.setCurrentIndex(0)
        self.customerNameComboBox.setFixedWidth(300)

        # 국가
        self.countryLabel = QLabel("국가 :", self)
        self.countryComboBox = QComboBox(self)

        query = DB_Queries()
        rows = query.selectCustomerCountry()
        columnName = list(rows[0].keys())[0]
        items = [row[columnName] for row in rows]
        items.sort()
        self.countryComboBox.addItem("ALL")
        self.countryComboBox.addItems(items)
        self.countryComboBox.setCurrentIndex(0)
        self.countryComboBox.activated.connect(self.countryComboBox_Activated)
        self.countryComboBox.setFixedWidth(300)

        # 도시
        self.cityLabel = QLabel("도시 :", self)
        self.cityComboBox = QComboBox(self)

        query = DB_Queries()
        rows = query.selectCustomerCity()
        columnName = list(rows[0].keys())[0]
        items = [row[columnName] for row in rows]
        items.sort()
        self.cityComboBox.addItem("ALL")
        self.cityComboBox.addItems(items)
        self.cityComboBox.activated.connect(self.cityComboBox_Activated)
        self.cityComboBox.setCurrentIndex(0)
        self.cityComboBox.setFixedWidth(300)

        self.customerSearchLabel = QLabel("검색된 주문의 개수 :", self)
        self.countlabel = QLabel(self)

        # 버튼
        self.btnClear = QPushButton("초기화", self)
        self.btnClear.resize(80,30)
        self.btnClear.setToolTip("초기화")
        self.btnClear.clicked.connect(self.btnClear_Clicked)

        self.btnSearch = QPushButton("검색", self)
        self.btnSearch.resize(80,30)
        self.btnSearch.setToolTip("검색")
        self.btnSearch.clicked.connect(self.btnSearch_Clicked)

        self.table_label = QLabel("검색된 주문 리스트", self)
        self.tableWidget = QTableWidget(self)

        # 테이블 클릭시 이벤트
        self.tableWidget.cellClicked.connect(self.cell_Clicked)
        self.tableWidget.resize(1000, 500)

        # 레이아웃
        self.Layout = QHBoxLayout()
        self.Layout.addWidget(self.mainLabel)

        self.Layout_1 = QHBoxLayout()
        self.Layout_1.addWidget(self.customerNameLabel)
        self.Layout_1.addWidget(self.customerNameComboBox)
        self.Layout_1.addStretch(1)
        self.Layout_1.addWidget(self.countryLabel)
        self.Layout_1.addWidget(self.countryComboBox)
        self.Layout_1.addStretch(1)
        self.Layout_1.addWidget(self.cityLabel)
        self.Layout_1.addWidget(self.cityComboBox)
        self.Layout_1.addStretch(2)
        self.Layout_1.addWidget(self.btnSearch)

        self.Layout_2 = QHBoxLayout()
        self.Layout_2.addWidget(self.customerSearchLabel)
        self.Layout_2.addWidget(self.countlabel)
        self.Layout_2.addStretch(2)
        self.Layout_2.addWidget(self.btnClear)

        self.Layout_3 = QHBoxLayout()
        self.Layout_3.addWidget(self.table_label)

        self.Layout_4 = QHBoxLayout()
        self.Layout_4.addWidget(self.tableWidget)

        self.mainLayout = QVBoxLayout()
        self.mainLayout.addLayout(self.Layout)
        self.mainLayout.addLayout(self.Layout_1)
        self.mainLayout.addLayout(self.Layout_2)
        self.mainLayout.addLayout(self.Layout_3)
        self.mainLayout.addLayout(self.Layout_4)

        self.setLayout(self.mainLayout)

    #  MainWindow 메소드
    def customerNameComboBox_Activated(self):
        self.customerName = self.customerNameComboBox.currentText()

    def countryComboBox_Activated(self):
        self.country = self.countryComboBox.currentText()
        query = DB_Queries()
        rowsCity = query.updateCustomerCity(self.country)
        columnNameCity = list(rowsCity[0].keys())[0]
        cityItems = [row[columnNameCity] for row in rowsCity]
        cityItems.sort()
        self.cityComboBox.clear()
        self.cityComboBox.addItem("ALL")
        self.cityComboBox.addItems(cityItems)
        self.city = self.cityComboBox.currentText()
        self.cityComboBox.activated.connect(self.cityComboBox_Activated)


    def cityComboBox_Activated(self):
        self.city = self.cityComboBox.currentText()

    def btnClear_Clicked(self):
        msg = "초기화하시겠습니까?"
        self.countlabel.setText('0')
        self.customerName = "ALL"
        self.customerNameComboBox.setCurrentIndex(0)
        self.country = "ALL"
        self.countryComboBox.setCurrentIndex(0)
        self.city = "ALL"
        self.cityComboBox.setCurrentIndex(0)

    def btnSearch_Clicked(self):
        query = DB_Queries()
        rows = query.selectDatabyKey(self.customerName,self.country,self.city)
        if rows:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(rows))
            self.tableWidget.setColumnCount(len(rows[0]))
            columnNames = list(rows[0].keys())
            self.tableWidget.setHorizontalHeaderLabels(columnNames)
            self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

            rows.sort(key = lambda x : x['orderNo'])
            for rowIDX, value in enumerate(rows):
                for columnIDX, (k, v) in enumerate(value.items()):
                    if v == None:
                        continue
                    else:
                        item = QTableWidgetItem(str(v))

                    self.tableWidget.setItem(rowIDX, columnIDX, item)

            self.tableWidget.resizeColumnsToContents()
            self.tableWidget.resizeRowsToContents()
            self.countlabel.setText("%s" % (len(rows)))

        else:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            self.countlabel.setText('0')

    def cell_Clicked(self):
        row = self.tableWidget.currentIndex().row()
        orderNo = self.tableWidget.item(row,0)
        self.hide()
        SW = SubWindow(orderNo.text())
        SW.exec()
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    mainWindow.btnSearch_Clicked()
    app.exec_()