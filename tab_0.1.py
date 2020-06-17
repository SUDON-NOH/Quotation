import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QGridLayout, QPushButton, \
    QMessageBox, QDesktopWidget, QTabWidget, QTextEdit, QVBoxLayout, QGroupBox, QMainWindow, \
    QHBoxLayout, QTableWidget, QRadioButton, QTableWidgetItem, QAbstractItemView, QDialog, QComboBox, QSpinBox
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt
import excel_file as ef
import pandas as pd
import openpyxl
import datetime


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
        self.left = 120
        self.top = 200
        self.width = 1600
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
        self.Q_spin1_text = 1
        self.date = datetime.date.today().isoformat()
        self.setFont(QFont("Times New Roman ", 10, QFont.Bold))
        self.exeUI()

    def exeUI(self):

        # ---------------------------------------------------------------- 위젯
        left_groupbox = QGroupBox("C-CODE SEARCH")
        left_groupbox.setMaximumSize(400, 800)
        right_groupbox = QGroupBox("Quotation")
        right_groupbox_tbl = QGroupBox("입력")

        # ---------------------------------------------------------------- C-CODE SEARCH
        # left_groupbox
        lal1 = QLabel("검색: ")
        self.c_lin1 = QLineEdit()
        self.c_pbtn1 = QPushButton("확인")
        self.c_pbtn2 = QPushButton("입력")
        self.c_tbl1 = QTableWidget()

        # Table 짝수번째 색 변화
        self.c_tbl1.setAlternatingRowColors(True)
        # Table 셀을 선택할 때 전체 행을 선택하도록 설정
        self.c_tbl1.setSelectionBehavior(QAbstractItemView.SelectRows)
        # Table을 수정하지 못하도록 설정
        self.c_tbl1.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # Table에서 더블클릭을 하면 실행되는 기능
        self.c_tbl1.doubleClicked.connect(self.move_toQ)
        # Table 크기 조절

        # 버튼을 누르면 시작
        self.c_pbtn1.clicked.connect(self.tab1_ccode_btn)
        self.c_pbtn2.clicked.connect(self.tab1_ccode_btn2)
        self.c_lin1.returnPressed.connect(self.tab1_ccode_btn)

        leftInnerLayout_top = QHBoxLayout()
        leftInnerLayout_top.addWidget(lal1)
        leftInnerLayout_top.addWidget(self.c_lin1)
        leftInnerLayout_top.addWidget(self.c_pbtn1)
        leftInnerLayout_top.addWidget(self.c_pbtn2)

        leftInnerLayout_btm = QVBoxLayout()
        leftInnerLayout_btm.addLayout(leftInnerLayout_top)
        leftInnerLayout_btm.addWidget(self.c_tbl1)

        left_groupbox.setLayout(leftInnerLayout_btm)

        # ---------------------------------------------------------------- Quotation
        # right_groupbox
        lal1 = QLabel('CODE:')
        lal2 = QLabel('COMPANY:')
        lal3 = QLabel('NAME:')
        lal4 = QLabel('Quotation Number:')
        lal5 = QLabel('CATEGORY:')
        lal6 = QLabel('Number Of Type:')
        lal_date = QLabel(self.date)
        lal1.setFixedSize(100, 20)
        lal2.setFixedSize(150, 20)
        lal3.setFixedSize(100, 20)
        lal4.setFixedSize(150, 20)
        lal5.setFixedSize(100, 20)
        lal6.setFixedSize(150, 20)

        self.Q_lal1 = QLabel('')
        self.Q_lal2 = QLabel('')
        self.Q_lal3 = QLabel('')
        self.Q_lal4 = QLabel('')
        self.Q_com1 = QComboBox()
        self.Q_spin1 = QSpinBox()
        self.Q_pbtn1 = QPushButton("확인")
        self.Q_lal1.setFixedSize(100, 20)
        self.Q_lal2.setFixedSize(150, 20)
        self.Q_lal3.setFixedSize(100, 20)
        self.Q_lal4.setFixedSize(150, 20)
        self.Q_com1.setFixedSize(100, 20)
        self.Q_spin1.setFixedSize(150, 20)
        self.Q_pbtn1.setFixedSize(60, 30)

        # Quotation Groupbox
        Q_table_groupbox = QVBoxLayout()
        self.Q_tbl = QTableWidget()
        self.Q_tbl.setAlternatingRowColors(True)
        Q_table_groupbox.addWidget(self.Q_tbl)
        right_groupbox_tbl.setLayout(Q_table_groupbox)

        # lal 설정
        self.Q_lal1.setStyleSheet(
            "color: black;"
            # "border-style: solid;"
            # "border-width: 2px;"
            "background-color: #BFB1B1;"
            # "border-radius: 3px;"
            # 포트설정 !!!
            "font: bold large 'Malgun Gothic'"
            )
        self.Q_lal2.setStyleSheet(
            "color: black;"
            # "border-style: solid;"
            # "border-width: 2px;"
            "background-color: #BFB1B1;"
            # "border-radius: 3px;"
            # 포트설정 !!!
            "font: bold large 'Malgun Gothic'"
            )
        self.Q_lal3.setStyleSheet(
            "color: black;"
            # "border-style: solid;"
            # "border-width: 2px;"
            "background-color: #BFB1B1;"
            # "border-radius: 3px;"
            # 포트설정 !!!
            "font: bold large 'Malgun Gothic'"
            )
        self.Q_lal4.setStyleSheet(
            "color: black;"
            # "border-style: solid;"
            # "border-width: 2px;"
            "background-color: #BFB1B1;"
            # "border-radius: 3px;"
            # 포트설정 !!!
            "font: bold large 'Malgun Gothic'"
            )
        lal_date.setStyleSheet(
            "color: black;"
            "background-color: #E0E0E0;"
            "font: bold large 'Malgun Gothic';"
            "font-size: 30px"
            )

        # COMBOBOX ITEM
        self.Q_com1.addItem('PQ')
        self.Q_com1.addItem('NA')
        self.Q_com1.currentIndexChanged.connect(self.combo_act)
        # SPINBOX
        self.Q_spin1.setMinimum(1)
        self.Q_spin1.setMaximum(10)
        self.Q_spin1.valueChanged.connect(self.spin_act)
        # Q_pbtn1
        self.Q_pbtn1.clicked.connect(self.Q_tbl_show)

        rightInnerLayout_1 = QHBoxLayout()
        # rightInnerLayout_1.addStretch(0)
        rightInnerLayout_1.addWidget(lal1)
        rightInnerLayout_1.addWidget(self.Q_lal1)
        rightInnerLayout_1.addWidget(lal2)
        rightInnerLayout_1.addWidget(self.Q_lal2)
        # rightInnerLayout_1.addStretch(1)

        rightInnerLayout_2 = QHBoxLayout()
        rightInnerLayout_2.addWidget(lal3)
        rightInnerLayout_2.addWidget(self.Q_lal3)
        rightInnerLayout_2.addWidget(lal4)
        rightInnerLayout_2.addWidget(self.Q_lal4)

        rightInnerLayout_3 = QHBoxLayout()
        rightInnerLayout_3.addWidget(lal5)
        rightInnerLayout_3.addWidget(self.Q_com1)
        rightInnerLayout_3.addWidget(lal6)
        rightInnerLayout_3.addWidget(self.Q_spin1)

        rightInnerLayout_4 = QHBoxLayout()
        rightInnerLayout_4.addStretch(10)
        rightInnerLayout_4.addWidget(self.Q_pbtn1)
        rightInnerLayout_4.addStretch(1)

        rightInnerLayout_V = QVBoxLayout()
        rightInnerLayout_V.addLayout(rightInnerLayout_1)
        rightInnerLayout_V.addLayout(rightInnerLayout_2)
        rightInnerLayout_V.addLayout(rightInnerLayout_3)
        rightInnerLayout_V.addLayout(rightInnerLayout_4)

        right_groupbox.setLayout(rightInnerLayout_V)

        tablayout1 = QHBoxLayout()
        tablayout1.addWidget(right_groupbox)
        tablayout1.addWidget(lal_date)

        tablayout2 = QVBoxLayout()
        tablayout2.addLayout(tablayout1)
        tablayout2.addWidget(right_groupbox_tbl)

        tablayout3 = QHBoxLayout()
        tablayout3.addWidget(left_groupbox)
        tablayout3.addLayout(tablayout2)

        self.setLayout(tablayout3)

    # c_code 의 확인 button
    def tab1_ccode_btn(self):
        df_fun = ef.excel_pd()
        df = df_fun.search_ccode()

        ccode_text = self.c_lin1.text()
        df_search = df['COMPANY'].str.contains(ccode_text)

        DF = df[df_search]

        if DF.empty:
            reply = QMessageBox(self)
            reply.question(self, 'Error', '검색되지 않습니다.', QMessageBox.Yes)
        else:
            self.numROW = len(DF)
            self.numCOL = len(DF.columns)
            # ROW와 COLUMN 수 지정
            self.c_tbl1.setRowCount(self.numROW)
            self.c_tbl1.setColumnCount(self.numCOL)
            # COLUMN 지정
            self.c_tbl1.setHorizontalHeaderLabels(DF.columns.tolist())
            # 요소 넣기
            self.v_list = DF.values.tolist()
            for m, n in zip(self.v_list, range(self.numROW)):
                for a, b in zip(range(self.numCOL), m):
                    self.c_tbl1.setItem(n, a, QTableWidgetItem(b))

    # c_code 의 입력 button
    def tab1_ccode_btn2(self):
        tab1_btn2_play = InputDialog()
        tab1_btn2_play.exec_()

    # 검색기능
    def move_toQ(self):
        # 더블 클릭된 행을 불러와서 각각 이름을 선언한다.
        row = self.c_tbl1.currentRow()
        self.tb_code = self.c_tbl1.item(row, 0).text()
        self.tb_comp = self.c_tbl1.item(row, 1).text()
        self.tb_name = self.c_tbl1.item(row, 2).text()

        # N NUMBER 추출
        mkfunc = ef.excel_pd()
        self.Nnumber = mkfunc.MK_DB()
        self.Nrow = len(self.Nnumber)
        last = self.Nnumber[self.Nnumber.columns[0]][-1:].values.tolist()
        # 마지막 N 번호 다음 번호를 붙일 땐 +1 해서 사용
        self.last = int(last[0][6:])
        self.N = str(self.last + 1)

        Message1 = self.tb_code
        Message2 = self.tb_comp
        Message3 = self.tb_name
        Message4 = self.N

        self.Q_lal1.setText(Message1)
        self.Q_lal2.setText(Message2)
        self.Q_lal3.setText(Message3)
        self.Q_lal4.setText(Message4)

    # Quotation COMBOBOX
    def combo_act(self):
        self.Q_com1_text = self.Q_com1.currentText()

    # Quotation SPINBOX
    def spin_act(self):
        # INT로 활용하기 위해 value로 값을 받음
        self.Q_spin1_text = self.Q_spin1.value()

    def Q_tbl_show(self):
        self.Q_tbl.setRowCount(self.Q_spin1_text)
        self.Q_tbl.setColumnCount(6)
        column_list = ['CLASS', 'NAME', 'SUB', '규격', '수량', '단가']
        self.Q_tbl.setHorizontalHeaderLabels(column_list)
        self.Q_tbl.setColumnWidth(0, 250)
        self.Q_tbl.setColumnWidth(1, 400)
        self.Q_tbl.setColumnWidth(2, 200)
        self.Q_tbl.setColumnWidth(3, 80)
        self.Q_tbl.setColumnWidth(4, 40)
        self.Q_tbl.setColumnWidth(5, 140)
        # self.Q_tbl.resizeColumnsToContents() 컬럼사이즈에 맞게

        self.combo_class_list = ['선택', 'RECOMBINANT PROTEIN', 'ANTIBODY', 'HYBRIDOMA', 'CULTURE MEDIA', 'BIACORE', 'SCREENING',
                                 'HUMANIZATION', 'AFFINITY MATURATION', 'IMMUNE LIBRARY', 'RCB', 'ELISA', 'FACS', 'ETC', 'SEC-HPLC']
        self.combo_size_list = ['VOL', 'ML', 'EA']
        self.combo_sub = {'선택':['NA'],
                          'RECOMBINANT PROTEIN':['1ST-PURI', '2ND-PURI', 'ENDOTOXIN', 'SEC-HPLC', '배양액'],
                          'ANTIBODY':['1ST-PURI', '2ND-PURI', 'CHIMERIC(hIGg1/Hkappa'],
                          'HYBRIDOMA':['IDENTIFICATION', '1ST-PURI'],
                          'CULTURE MEDIA':['NA'],
                          'BIACORE':['AFFINITY', 'MAPPING'],
                          'SCREENING':['BASIC BINDER', 'CUSTOM', 'FULL BINDER', 'BLOCKER', 'RAPID BINDER'],
                          'HUMANIZATION':['CDR GRAFTING', 'GUIDED SELECTION', 'TOTAL SOLUTION'],
                          'AFFINITY MATURATION':['LC SHUFFLING', 'HOT SPOT MUTATION', 'CORE PACKING', 'TOTAL SOLUTION'],
                          'IMMUNE LIBRARY':['IMMUNE LIBRARY ONLY', 'SCREENING'],
                          'RCB':['FULL SERVICE', 'POOL GENERATION'],
                          'ELISA':['AFFINITY', 'MAPPING'],
                          'FACS':['NA'],
                          'ETC':['BUFFER', 'ENDOTOXIN', 'SEC-HPLC'],
                          'SEC-HPLC':['PURITY']
                          }

        for i in range(self.Q_spin1_text):

            globals()['combo_size{}'.format(i)] = QComboBox()
            globals()['combo_class{}'.format(i)] = QComboBox()
            self.spin = QSpinBox()
            self.spin.setMinimum(1)
            self.spin.setMaximum(1000)

            globals()['combo_class{}'.format(i)].addItems(self.combo_class_list)
            globals()['combo_size{}'.format(i)].addItems(self.combo_size_list)

            self.Q_tbl.setCellWidget(i, 0, globals()['combo_class{}'.format(i)])
            # self.Q_tbl.setCellWidget(i, 2, globals()['combo_sub{}'.format(i)])
            self.Q_tbl.setCellWidget(i, 3, globals()['combo_size{}'.format(i)])
            self.Q_tbl.setCellWidget(i, 4, self.spin)


        # 각 콤보에서 필요한 기능 추가 -> 콤보가 변경될 때마다 작동하는 함수 필요
        # 콤보변화에 따른 하위 콤보의 변화
        if self.Q_spin1_text == 1:
            combo_class0.currentTextChanged.connect(self.combo_change0)
        elif self.Q_spin1_text == 2:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
        elif self.Q_spin1_text == 3:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
        elif self.Q_spin1_text == 4:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
        elif self.Q_spin1_text == 5:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
        elif self.Q_spin1_text == 6:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
            combo_class5.currentTextChanged.connect(self.combo_change5)
        elif self.Q_spin1_text == 7:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
            combo_class5.currentTextChanged.connect(self.combo_change5)
            combo_class6.currentTextChanged.connect(self.combo_change6)
        elif self.Q_spin1_text == 8:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
            combo_class5.currentTextChanged.connect(self.combo_change5)
            combo_class6.currentTextChanged.connect(self.combo_change6)
            combo_class7.currentTextChanged.connect(self.combo_change7)
        elif self.Q_spin1_text == 9:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
            combo_class5.currentTextChanged.connect(self.combo_change5)
            combo_class6.currentTextChanged.connect(self.combo_change6)
            combo_class7.currentTextChanged.connect(self.combo_change7)
            combo_class8.currentTextChanged.connect(self.combo_change8)
        elif self.Q_spin1_text == 10:
            combo_class0.currentTextChanged.connect(self.combo_change0)
            combo_class1.currentTextChanged.connect(self.combo_change1)
            combo_class2.currentTextChanged.connect(self.combo_change2)
            combo_class3.currentTextChanged.connect(self.combo_change3)
            combo_class4.currentTextChanged.connect(self.combo_change4)
            combo_class5.currentTextChanged.connect(self.combo_change5)
            combo_class6.currentTextChanged.connect(self.combo_change6)
            combo_class7.currentTextChanged.connect(self.combo_change7)
            combo_class8.currentTextChanged.connect(self.combo_change8)
            combo_class9.currentTextChanged.connect(self.combo_change9)

    def combo_change0(self):
        CD0 = QComboBox()
        list0 = self.combo_sub[combo_class0.currentText()]
        CD0.addItems(list0)
        self.Q_tbl.setCellWidget(0, 2, CD0)
    def combo_change1(self):
        CD1 = QComboBox()
        list1 = self.combo_sub[combo_class1.currentText()]
        CD1.addItems(list1)
        self.Q_tbl.setCellWidget(1, 2, CD1)
    def combo_change2(self):
        CD2 = QComboBox()
        list2 = self.combo_sub[combo_class2.currentText()]
        CD2.addItems(list2)
        self.Q_tbl.setCellWidget(2, 2, CD2)
    def combo_change3(self):
        CD3 = QComboBox()
        list3 = self.combo_sub[combo_class3.currentText()]
        CD3.addItems(list3)
        self.Q_tbl.setCellWidget(3, 2, CD3)
    def combo_change4(self):
        CD4 = QComboBox()
        list4 = self.combo_sub[combo_class4.currentText()]
        CD4.addItems(list4)
        self.Q_tbl.setCellWidget(4, 2, CD4)
    def combo_change5(self):
        CD5 = QComboBox()
        list5 = self.combo_sub[combo_class5.currentText()]
        CD5.addItems(list5)
        self.Q_tbl.setCellWidget(5, 2, CD5)
    def combo_change6(self):
        CD6 = QComboBox()
        list6 = self.combo_sub[combo_class6.currentText()]
        CD6.addItems(list6)
        self.Q_tbl.setCellWidget(6, 2, CD6)
    def combo_change7(self):
        CD7 = QComboBox()
        list7 = self.combo_sub[combo_class7.currentText()]
        CD7.addItems(list7)
        self.Q_tbl.setCellWidget(7, 2, CD7)
    def combo_change8(self):
        CD8 = QComboBox()
        list8 = self.combo_sub[combo_class8.currentText()]
        CD8.addItems(list8)
        self.Q_tbl.setCellWidget(0, 2, CD8)
    def combo_change9(self):
        CD9 = QComboBox()
        list9 = self.combo_sub[combo_class9.currentText()]
        CD9.addItems(list9)
        self.Q_tbl.setCellWidget(0, 2, CD9)

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    first = exe_func()
    sys.exit(app.exec_())