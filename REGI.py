import time
import pandas as pd
import openpyxl
from datetime import datetime

"""


개인 혹은 팀 REGI에 접속해서
파일에 없는 내용을 업데이트


"""
# 팀 경로
TEAM_list = [TEAM02_load, TEAM04_load, TEAM06_load, TEAM07_load, TEAM08_load]
# TEAM_list = [TEAM07_load, TEAM08_load]
# 개인 REGI 파일 이름
REGI = [TEAM02, TEAM04, TEAM06, TEAM07, TEAM08]
# REGI = [TEAM07, TEAM08]
# SHEET의 COLUMNS
columns = [FA_columns, FT_columns, FS_columns, FR_columns, FD_columns]
# SHEET의 COLUMNS의 제한
limit_column = [26, 11, 6, 8, 12]
# 컬럼의 위치
alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
        'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM']

for team, regi in zip(TEAM_list, REGI):
    # 접근 주소는 team + regi[ynum]
    for ynum in range(len(regi)):
        # print("ynum: ", ynum)
        # sheet 번호
        for sheet, col, lcol in zip(SHEET, columns, limit_column):

            Total_1 = pd.read_excel(Total_load, sheet_name=sheet)
            Total = Total_1[col[0]].values.tolist()
            # print(sheet)
            # print(regi[ynum], sheet)
            # print("오류1")
            file_1 = pd.read_excel(team+regi[ynum], sheet_name=sheet)
            # print("오류2")
            file_2 = file_1.drop(file_1.index[0:6])
            # print(team+regi[ynum], '  ',sheet)
            # print("오류3")
            file_3 = file_2[file_2.columns[3:lcol]]
            # print("오류4")
            file_3.columns = col
            # print("오류5")
            # NaN 값을 제거 후 비어있는 데이터프레임이 되면 pass 아니면 진행
            # 모든 행의 값이 NaN인 경우에 행 삭제
            file_4 = file_3.dropna(how = 'all')
            file_5 = file_4[~file_4[col[0]].isin(Total)]
            # print(file_5)
            # print("오류6")

            if file_5.empty:
                # print("오류7")
                print("\n\n\n\n\n========     업데이트 된 목록이 없습니다.     ========\n\n", "     TEAM: ", team[19:25], '\n',
                      "     REGI: ", regi[ynum], '\n', "     SHEET: ", sheet, "\n\n====================================================\n\n\n\n")
                time.sleep(2)
                pass
            else:
                print(file_5)
                # print("오류8")
                # print("오류9")
                file = file_5.values.tolist()
                # Read_OR 참고해서, 길이는 alpha 를 3:lcol 로 잘라서 활용, 시작점은 len(Total_1) 을 사용
                # print("오류10")
                COL = alpha[0:lcol-3]
                # print("오류11")
                ROW = str(len(Total_1) + 2)
                # print("오류12")
                wb = openpyxl.load_workbook(Total_load)
                # print("오류13")
                w_sheet = wb[sheet]
                # print("오류14")
                for k in file:
                    for m, i in zip(COL, k):
                        z = m + ROW
                        w_sheet[z] = i
                        if m == COL[-1]:
                            DT = datetime.today().strftime("%Y-%m-%d %H:%M")
                            if sheet == "FA":
                                w_sheet["X" + ROW] = DT
                            elif sheet == "FT":
                                w_sheet["I" + ROW] = DT
                            elif sheet == "FS":
                                w_sheet["D" + ROW] = regi[ynum][0:5]
                                w_sheet["E" + ROW] = DT
                            elif sheet == "FR":
                                w_sheet["F" + ROW] = regi[ynum][0:5]
                                w_sheet["G" + ROW] = DT
                            else:
                                w_sheet["J" + ROW] = DT
                            ROW = int(ROW)
                            ROW += 1
                            ROW = str(ROW)
                # print("오류15")
                wb.save(Total_load)
                print("\n\n\n\n\n================  SAVE  =====================\n\n", "     TEAM: ", team[19:25], '\n',
                      "     REGI: ", regi[ynum], '\n', "     SHEET: ", sheet, "\n\n=============================================\n\n\n\n\n\n")
                print("==============       Sleep 10 seconds from now on...       ==============")
                time.sleep(1)
                print("                                   1")
                time.sleep(1)
                print("                                   2")
                time.sleep(1)
                print("                                   3")
                time.sleep(1)
                print("                                   4")
                time.sleep(1)
                print("                                   5")
                time.sleep(1)
                print("                                   6")
                time.sleep(1)
                print("                                   7")
                time.sleep(1)
                print("                                   8")
                time.sleep(1)
                print("                                   9")
                time.sleep(1)
                print("                                   10")


                print("=====================            Wake up            =====================\n\n\n\n\n\n\n\n")

print("\n\n\n\n\n\n\n================================              FINISHED              ================================")
