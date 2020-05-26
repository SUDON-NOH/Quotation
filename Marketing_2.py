import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QGridLayout, QPushButton, \
    QMessageBox, QDesktopWidget, QTabWidget, QTextEdit, QVBoxLayout, QGroupBox, QMainWindow, QHBoxLayout
from PyQt5.QtGui import QIcon
import excel_file as ef
import pandas as pd

class Login(QWidget):
    def __init__(self):
        super().__init__()
        self.loginUI()
        # 로그인에 사용할 Ycode 와 NAME
        self.Ycode = None
        self.NAME = None

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

        # Enter를 치면 아래의 명령이 실행되도록록
        self.pushButton.clicked.connect(self.pushButtonCliked)
        self.lineEdit2.returnPressed.connect(self.pushButtonCliked)

        self.setWindowTitle('Ybio')
        self.resize(300, 100)
        self.setWindowIcon(QIcon('ybio.png'))
        self.center()
        self.show()

    # 로그인 버튼 입력시 발생
    def pushButtonCliked(self):

        # Line에 입력된 값을 텍스트로 불러옴
        self.Ycode = self.lineEdit1.text()
        self.NAME = self.lineEdit2.text()

        # 로그인 정보가 있는 EXCEL을 불러오기 위치 찾기
        excel_LDB = ef.excel_pd()
        login_DB = excel_LDB.load_LDB()
        self.login_db = pd.read_excel(login_DB)
        # Ycode와 이름이 일치할 경우 DF 생성, 아닐경우 빈 DF 생성으로 로그인 성공과 실패를 결정
        self.login_info = self.login_db.loc[(self.login_db['Ynumber'] == self.Ycode) & (self.login_db['Name'] == self.NAME)]

        # 위치에 없으면 다시 입력하는 칸이 나오고, 있을 경우 close 하고 다음 실행
        if self.login_info.empty:
            reply = QMessageBox(self)
            reply.question(self, 'Error', '사번과 이름이 일치하지 않습니다.', QMessageBox.Yes)
        else:
            self.close()
            # 다음 위젯을 실행
            self.exe_load = exe_func()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

class exe_func(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ybio')
        self.setWindowIcon(QIcon('ybio.png'))
        self.left = 300
        self.top = 100
        self.width = 1360
        self.height = 800
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.tab_widgets = tab_widget()
        self.setCentralWidget(self.tab_widgets)

        self.show()

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
        self.tab1.layout.addWidget(self.c_code_search, 0, 0)

        self.c_code_label = QLabel('검 색:')
        self.c_code_line = QLineEdit()
        self.c_code_btn = QPushButton('확인')

        # c_code_search 의 GroupBox 에 대한 정렬코드
        # label 과 line 의 위치를 설정
        Hbox = QHBoxLayout()
        Vbox = QVBoxLayout()
        Hbox.addStretch(1)
        Hbox.addWidget(self.c_code_label)
        Hbox.addWidget(self.c_code_line)
        Hbox.addStretch(1)
        Vbox.addStretch(1)
        Vbox.addLayout(Hbox)
        Vbox.addStretch(20)
        self.c_code_search.setLayout(Vbox)

    # tab1의 Quotation
    def Quotation(self):
        self.Quotation_1 = QGroupBox('Quotation')
        self.tab1.layout.addWidget(self.Quotation_1, 0, 1)






    #     # C_code 칸
    #     self.c_code_search = self.code_search()
    #     self.tab1.layout.addWidget(self.c_code_search, 0, 0)
    #     self.tab1.layout.addWidget(self.Quatation(), 1, 0)
    #
    #     self.tab1.layout.addWidget(self.tabs)
    #     self.setLayout(self.tab1.layout)

        # # Table 생성
        # self.tablewidget = QTableWidget(self.tab1)
        # self.tablewidget.resize(290, 290)
        # self.tablewidget.setRowCount(2)
        # self.tablewidget.setColumnCount(2)
        # self.setTableWidgetData()

        # def setTableWidgetData(self):
        #     self.tablewidget.setItem(0, 0, QTableWidgetItem("노수돈 0"))
        #     self.tablewidget.setItem(0, 1, QTableWidgetItem("노수돈 1"))
        #     self.tablewidget.setItem(1, 0, QTableWidgetItem("노수돈 2"))
        #     self.tablewidget.setItem(1, 1, QTableWidgetItem("노수돈 3"))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    first = Login()
    sys.exit(app.exec_())