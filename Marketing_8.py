import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QGridLayout, QPushButton, \
    QMessageBox, QDesktopWidget, QTabWidget, QTextEdit, QVBoxLayout, QGroupBox, QMainWindow, \
    QHBoxLayout, QTableWidget, QRadioButton, QTableWidgetItem, QAbstractItemView, QDialog
from PyQt5.QtGui import QIcon
import excel_file as ef
import pandas as pd
import openpyxl


class exe_func(QMainWindow):
    def __init__(self):
        super().__init__()
        login = Login()
        login.exec_()
        self.login_Ycode = login.Ycode
        self.login_Name = login.NAME

        # login_info가 비어있는 경우 main을 실행하지 않고 닫기
        if not login.login_info.empty:
            self.mainUI()
        else:
            # 시스템 종료
            sys.exit()

    def mainUI(self):
        self.setWindowTitle('Ybio')
        self.setWindowIcon(QIcon('ybio.png'))
        self.left = 300
        self.top = 100
        self.width = 1200
        self.height = 700

        # 상태바에 표시
        Message = '       CODE:   ' + self.login_Ycode +'        ' + '       NAME :   ' +self.login_Name
        self.statusBar().showMessage(Message)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.tab_widgets = tab_widget()
        self.setCentralWidget(self.tab_widgets)

        self.show()

class Login(QDialog):
    def __init__(self):
        super().__init__()

        self.Ycode = None
        self.NAME = None
        self.login_info = pd.DataFrame()

        self.loginUI()

    # 로그인 UI
    def loginUI(self):
        # Ycode와 NAME을 입력할 Label과 Line을 생성
        self.label1 = QLabel("코 드: ")
        self.label2 = QLabel("이 름: ")
        self.lineEdit1 = QLineEdit()
        self.lineEdit2 = QLineEdit()
        self.pushButton = QPushButton("Login")

        # 위치 입력
        layout = QGridLayout()
        layout.addWidget(self.label1, 0, 0)
        layout.addWidget(self.lineEdit1, 0, 1)
        layout.addWidget(self.label2, 1, 0)
        layout.addWidget(self.lineEdit2, 1, 1)
        layout.addWidget(self.pushButton, 1, 2)

        self.setLayout(layout)

        self.pushButton.clicked.connect(self.pushButtonCliked)

        self.setWindowTitle('Ybio')
        self.resize(300, 100)
        self.setWindowIcon(QIcon('ybio.png'))
        # self.center()
        self.show()

    # 로그인 버튼 입력시 발생
    def pushButtonCliked(self):

        # 로그인 정보가 있는 EXCEL을 불러오기 위치 찾기
        excel_LDB = ef.excel_pd()
        login_DB = excel_LDB.load_LDB()

        self.login_db = pd.read_excel(login_DB)

        # Line에 입력된 값을 텍스트로 불러옴
        self.Ycode = self.lineEdit1.text()
        self.NAME = self.lineEdit2.text()

        # Ycode와 이름이 일치할 경우 DF 생성, 아닐경우 빈 DF 생성으로 로그인 성공과 실패를 결정
        self.login_info = self.login_db.loc[(self.login_db['Ycode'] == self.Ycode) & (self.login_db['Name'] == self.NAME)]

        # 위치에 없으면 다시 입력하는 칸이 나오고, 있을 경우 close 하고 다음 실행
        if not self.login_info.empty:
            self.close()

        else:
            reply = QMessageBox(self)
            reply.question(self, 'Error', '사번 및 이름을 확인하세요.', QMessageBox.Yes)

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

class tab_widget(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        self.exeUI()

    def exeUI(self):
        # tab 설정

        self.tabs = QTabWidget()

        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tab4 = QWidget()

        # 각 TAB 이름 설정 및 추가

        self.tabs.addTab(self.tab1, 'Quotation')
        self.tabs.addTab(self.tab2, 'Progress')
        self.tabs.addTab(self.tab3, 'Result')
        self.tabs.addTab(self.tab4, 'Report')

        # tab1의 실행 명령어

        self.tab1_play = self.tab1_f()

        # tab 전체 위젯을 완성시키는 명령어

        self.layout.addWidget(self.tabs)

        self.setLayout(self.layout)

    # tab1 실행 함수
    def tab1_f(self):
        # tab1 의 layout 을 그리드로 설정
        self.tab1.layout = QGridLayout()

        self.tab1_func1 = self.code_c_search()
        self.tab1_func2 = self.Quotation()

        self.tab1.setLayout(self.tab1.layout)

    # tab1 의 code_c_search
    def code_c_search(self):
        self.c_code_search = QGroupBox('C-CODE SEARCH')
        # GropBox 크기 조절
        self.c_code_search.setMaximumSize(430, 800)
        self.tab1.layout.addWidget(self.c_code_search, 0, 0)

        self.c_code_label = QLabel('검 색:')
        self.c_code_line = QLineEdit()
        self.c_code_btn = QPushButton('확인')
        self.c_code_btn2 = QPushButton('입력')
        self.c_code_table = QTableWidget(self.c_code_search)

        # TableWidget 설정
        self.c_code_table.resize(300, 700)
        # Table 짝수번째 색 변화
        self.c_code_table.setAlternatingRowColors(True)
        # Table 셀을 선택할 때 전체 행을 선택하도록 설정
        self.c_code_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        # Table을 수정하지 못하도록 설정
        self.c_code_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # Table에서 더블클릭을 하면 실행되는 기능
        self.c_code_table.doubleClicked.connect(self.move_toQ)


        # 버튼을 누르면 시작
        self.c_code_btn.clicked.connect(self.tab1_ccode_btn)
        self.c_code_btn2.clicked.connect(self.tab1_ccode_btn2)
        self.c_code_line.returnPressed.connect(self.tab1_ccode_btn)

        # c_code_search 의 GroupBox 에 대한 정렬코드
        # label 과 line 의 위치를 설정
        Hbox = QHBoxLayout()
        Vbox = QVBoxLayout()
        Hbox.addStretch(1)
        Hbox.addWidget(self.c_code_label)
        Hbox.addWidget(self.c_code_line)
        Hbox.addWidget(self.c_code_btn)
        Hbox.addStretch(1)
        # Vbox.addStretch(0)
        Vbox.addLayout(Hbox)
        # Vbox.addStretch(20)
        Vbox.addWidget(self.c_code_table)
        Vbox.addWidget(self.c_code_btn2)
        self.c_code_search.setLayout(Vbox)

    # c_code 의 확인 button
    def tab1_ccode_btn(self):
        df_fun = ef.excel_pd()
        df = df_fun.search_ccode()

        ccode_text = self.c_code_line.text()
        df_search = df['COMPANY'].str.contains(ccode_text)
        DF = df[df_search]

        if DF.empty:
            reply = QMessageBox(self)
            reply.question(self, 'Error', '검색되지 않습니다.', QMessageBox.Yes)
        else:
            self.numROW = len(DF)
            self.numCOL = len(DF.columns)
            # ROW와 COLUMN 수 지정
            self.c_code_table.setRowCount(self.numROW)
            self.c_code_table.setColumnCount(self.numCOL)
            # COLUMN 지정
            self.c_code_table.setHorizontalHeaderLabels(DF.columns.tolist())
            # 요소 넣기
            self.v_list = DF.values.tolist()
            for m, n in zip(self.v_list, range(self.numROW)):
                for a, b in zip(range(self.numCOL), m):
                    self.c_code_table.setItem(n, a, QTableWidgetItem(b))

    # c_code 의 입력 button
    def tab1_ccode_btn2(self):
        tab1_btn2_play = InputDialog()
        tab1_btn2_play.exec_()

    # 검색기능
    def move_toQ(self):
        # 더블 클릭된 행을 불러와서 각각 이름을 선언한다.
        row = self.c_code_table.currentRow()
        self.tb_code = self.c_code_table.item(row, 0).text()
        self.tb_comp = self.c_code_table.item(row, 1).text()
        self.tb_name = self.c_code_table.item(row, 2).text()

        # N NUMBER 추출
        mkfunc = ef.excel_pd()
        self.Nnumber = mkfunc.MK_DB()
        self.Nrow = len(self.Nnumber)
        last = self.Nnumber[self.Nnumber.columns[0]][-1:].values.tolist()
        # 마지막 N 번호 다음 번호를 붙일 땐 +1 해서 사용
        self.last = int(last[0][6:])
        self.N = str(self.last + 1)

        Message1 = 'C-CODE:  ' + self.tb_code
        Message2 = 'COMPANY:  ' + self.tb_comp
        Message3 = 'NAME:  ' + self.tb_name
        Message4 = 'Quotation Number:  ' + self.N

        self.Quotation_label1.setText(Message1)
        self.Quotation_label2.setText(Message2)
        self.Quotation_label3.setText(Message3)
        self.Quotation_number.setText(Message4)

    # tab1의 Quotation
    def Quotation(self):
        self.Quotation_1 = QGroupBox('Quotation')
        self.Quotation_1.setStyleSheet('border-style: solid;'
                                       'border-width: 2px;'
                                       'border-color: #9E9E9E;'
                                       'border-radius: 3px')
        self.Quotation_1.setMaximumSize(700, 100)
        self.tab1.layout.addWidget(self.Quotation_1, 0, 1)

        self.Quotation_label1 = QLabel('C-CODE: ')
        self.Quotation_label2 = QLabel('C-COMP: ')
        self.Quotation_label3 = QLabel('C-NAME: ')
        self.Quotation_number = QLabel('Quotation Number: ')

        layout = QGridLayout()
        layout.addWidget(self.Quotation_number, 0, 1)
        layout.addWidget(self.Quotation_label1, 1, 0)
        layout.addWidget(self.Quotation_label2, 1, 1)
        layout.addWidget(self.Quotation_label3, 1, 2)

        self.Quotation_1.setLayout(layout)

# c_code popup창
class InputDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setupUI()

        self.COMPANY = None
        self.NAME = None

    def setupUI(self):
        self.label1 = QLabel("COMPANY")
        self.label2 = QLabel("NAME")
        self.line1 = QLineEdit()
        self.line2 = QLineEdit()
        self.btn = QPushButton('저장')

        # 위치 입력
        layout = QGridLayout()
        layout.addWidget(self.label1, 0, 0)
        layout.addWidget(self.line1, 0, 1)
        layout.addWidget(self.label2, 1, 0)
        layout.addWidget(self.line2, 1, 1)
        layout.addWidget(self.btn, 1, 2)

        self.setLayout(layout)

        # Enter를 치면 아래의 명령이 실행되도록록
        self.btn.clicked.connect(self.pushButtonCliked)
        self.line2.returnPressed.connect(self.pushButtonCliked)

        self.setWindowTitle('Ybio')
        self.resize(300, 100)
        self.setWindowIcon(QIcon('ybio.png'))
        # CENTER 기능이 없음
        self.show()

    def pushButtonCliked(self):
        adr_a = ef.excel_pd()
        df = adr_a.search_ccode()
        # 여기서 x는 마지막행을 말한다.
        p_x = len(df)
        y = str(p_x + 1)
        x = str(p_x + 2)
        wb = openpyxl.load_workbook(adr_a.DB_Cfile)
        w_sheet = wb['DB']
        w_sheet['A' + x] = self.line1.text()
        w_sheet['B' + x] = self.line2.text()
        w_sheet['C' + x] = 'C' + y.zfill(4)
        wb.save(adr_a.DB_Cfile)
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    first = exe_func()
    sys.exit(app.exec_())