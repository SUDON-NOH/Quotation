import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTableWidget, QHBoxLayout, QVBoxLayout, QTableWidgetItem, QComboBox, QMessageBox
from PyQt5.QtGui import QIcon
import pandas as pd

class search(QWidget):
    def __init__(self):
        super().__init__()
        self.combo_text = ''
        self.searchUI()

    def searchUI(self):
        self.combo = QComboBox()
        self.label = QLabel("검색:")
        self.lineEdit = QLineEdit()
        self.pushbtn = QPushButton("확인")
        self.table = QTableWidget()

        self.table.resize(700, 400)
        self.table.setAlternatingRowColors(True)

        self.pushbtn.clicked.connect(self.search)
        self.lineEdit.returnPressed.connect(self.search)
        self.combo.currentIndexChanged.connect(self.combo_act)

        # combobox
        self.combo.addItem('TYPE')
        self.combo.addItem('ITEM')
        self.combo.addItem('PU')
        # self.combo.addItem('LR')
        self.combo.addItem('LQ')
        # self.combo.addItem('PSN')
        # self.combo.addItem('STN')
        # self.combo.addItem('W')

        Hbox = QHBoxLayout()
        Hbox.addWidget(self.label)
        Hbox.addWidget(self.combo)
        Hbox.addWidget(self.lineEdit)
        Hbox.addWidget(self.pushbtn)

        Vbox = QVBoxLayout()
        Vbox.addLayout(Hbox)
        Vbox.addWidget(self.table)

        self.setLayout(Vbox)

        self.setWindowTitle('Ybio')
        self.setWindowIcon(QIcon('Ybio.png'))
        self.resize(750, 450)
        self.show()

    def search(self):
        search_text = self.lineEdit.text()
        search_text = search_text.upper()

        if self.combo_text == 'TYPE':
            reply = QMessageBox(self)
            reply.question(self, 'ERROR', 'TYPE을 입력하세요.', QMessageBox.Yes)
            DF = pd.DataFrame()

        if self.combo_text == 'ITEM':
            func = self.ITEM_f()
            df = self.ITEM
        elif self.combo_text == 'PU':
            df = self.PU_f()
        elif self.combo_text == 'LQ':
            df = self.LQ_f()

        if self.combo_text != 'TYPE':
            df_search = df['NAME'].str.contains(search_text)
            DF = df[df_search]

        if DF.empty:
             self.reply = QMessageBox(self)
             self.reply.question(self, 'Error', '검색되지 않습니다.', QMessageBox.Yes)
        else:
            self.numROW = len(DF)
            self.numCOL = len(DF.columns)


            self.table.setRowCount(self.numROW)
            self.table.setColumnCount(self.numCOL)

            self.table.setHorizontalHeaderLabels(DF.columns.tolist())
            self.v_list = DF.values.tolist()

            for m, n in zip(self.v_list, range(self.numROW)):
                for a, b in zip(range(self.numCOL), m):
                    b = str(b)
                    self.table.setItem(n, a, QTableWidgetItem(b))


    def ITEM_f(self):
        self.ITEM_3 = pd.read_excel("Z:/ANRT/물품구매/INVENTORY(2.0).xlsx", sheet_name='M')
        self.ITEM_2 = self.ITEM_3.drop(self.ITEM_3.index[0:6])
        self.ITEM_2.columns = ['1', '2', '3', 'PSNID', 'ITID', 'VENID', 'CATID', '판매자', 'SUPPLYER', 'CAT', 'NAME',
                               'QUAN', '단가', '구매요청일', '담당자', '15', '16', '17', '18']
        self.ITEM = self.ITEM_2[['PSNID', 'CAT', '구매요청일', 'NAME', 'SUPPLYER', 'CATID', '담당자', '판매자']]
        self.ITEM = self.ITEM.fillna('-')
        P = self.ITEM['NAME'] != 0
        self.ITEM = self.ITEM[P]
        self.ITEM = self.ITEM.sort_values(["PSNID"], ascending = [False])
        self.ITEM = self.ITEM.reset_index(drop=True)
        for i in range(len(self.ITEM)):
            X = self.ITEM['구매요청일'][i]
            x = str(X)[0:10]
            self.ITEM['구매요청일'][i] = x
        self.ITEM['NAME'] = self.ITEM['NAME'].str.upper()
        return self.ITEM

    def PU_f(self):
        self.PU_3 = pd.read_excel("Z:/REGI/REGI/REGI008PU(2.0).xlsx", sheet_name='M')
        self.PU_2 = self.PU_3.drop(self.PU_3.index[0:6])
        self.PU_2.columns = ['1', '2', '3', 'PUID', 'TRID', 'SID', 'NAME', 'VOL', 'PRODUCTIVITY', 'DATE', 'CONC',
                             'AMOUNT', 'RESEARCHER', 'ELUTION', 'DIALYSIS', 'METHOD', 'COMMENT', 'BATCH']
        self.PU = self.PU_2[['PUID', 'TRID', 'NAME', 'SID', 'DATE']]
        self.PU['DATE'] = self.PU.loc[:,'DATE'].dt.strftime('%Y-%m-%d')
        self.PU = self.PU.sort_values(["PUID"], ascending = [False])
        self.PU = self.PU.reset_index(drop=True)
        self.PU = self.PU.fillna('-')
        self.PU['NAME'] = self.PU['NAME'].str.upper()
        return self.PU



    def LQ_f(self):
        self.LQ_3 = pd.read_excel("Z:/REGI/REGI/REGI008LQV2.xlsx", sheet_name='Main')
        self.LQ_2 = self.LQ_3.drop(self.LQ_3.index[0:6])
        self.LQ = self.LQ_2[['Unnamed: 4', 'Unnamed: 5', 'Unnamed: 8', 'Unnamed: 6', 'Unnamed: 7']]
        self.LQ.columns = ['LQID', 'TARGET', 'NAME', 'PUID', 'SID']
        self.LQ = self.LQ.sort_values(["LQID"], ascending = [False])
        self.LQ = self.LQ.reset_index(drop = True)
        self.LQ['NAME'] = self.LQ['NAME'].str.upper()
        return self.LQ

    def combo_act(self):
        self.combo_text = self.combo.currentText()

    # def LR_f(self):
    # def STN_f(self):
    # def W_f(self):

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = search()
    sys.exit(app.exec_())