import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt, QThread
from CrawlingUi import Ui_DomecallCrawling
import time

# 보안할 부분
# 1. Key Value로 데이터 넣기
# 2. 쓰레드 10개로 탐색
# 2.5 쓰레드 올리지말고 크롬드라이버 생성 for문으로 100개 실행하고 i순서대로 값 받으면 됌
# 3 소박스 대박스 try except문으로 있으면 실행 없으면 없음 키 벨류에 넣기

def requestCrawling(startNum, EndNum, itemList, number):
    j = 0
    items = {'num': ['productCode', 'price', 'bacode', 'bigBoxCount', 'smallBoxCount', 'origin']}

    for i in range(startNum, EndNum):
        j += 1 
        url = f"https://www.domecall.net/goods/goods_view.php?goodsNo={itemList[i - startNum]}"
        request = requests.get(url)
        soup = BeautifulSoup(request.text, 'html.parser')
        price = str(soup.select_one('#frmView > div > div.item > ul > li.price > div > strong')).strip("</strong>")
        bacode = str(soup.select_one('#frmView > div > div.item > ul > li:nth-child(2) > div')).strip("</div>")
        productCode = str(soup.select_one('#frmView > div > div.item > ul > li:nth-child(3) > div')).strip("</div>")
        origin = str(soup.select_one('#frmView > div > div.item > ul > li:nth-child(4) > div')).strip("</div>")
        bigBoxCount = str(soup.select_one('#frmView > div > div.item > ul > li:nth-child(5) > div > span')).strip("</span>")
        smallBoxCount = str(soup.select_one('#frmView > div > div.item > ul > li:nth-child(6) > div > span')).strip("</span>")
        items[i] = [productCode, price, bacode, bigBoxCount, smallBoxCount, origin]
        textBrowser.append(f'쓰레드:{number}, {j}번 상품번호:{productCode}, 가격:{price}, 바코드번호:{bacode}, 큰박스:{bigBoxCount}, 작은박스:{smallBoxCount}')
    # 모두 끝남
    bol = True

    return items, bol


def merge_two_dicts(dicA, dicB):
    mergedic = dicA.copy()

    mergedic.update(dicB)

    return mergedic


# data_only=Ture로 해줘야 수식이 아닌 값으로 받아온다.
load_wb = load_workbook("productNumbers.xlsx", data_only=True)

# 시트 이름으로 불러오기
load_ws = load_wb['Sheet1']

def ListDivion(Num, Ui_startText, Ui_endText):
    # Excel item 불러 리스트 담기
    get_cells = load_ws[Ui_startText:Ui_endText]
    for row in get_cells:
        for cell in row:
            itemCodeTread[Num].append(cell.value)
    print(itemCodeTread)


class Example(QMainWindow, Ui_DomecallCrawling):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        # 프로그램이 항상 최상단에 위치하도록 지정(크롬에 가리지 않게)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.show()

    # start 쓰레드 1(예매페이지 접속)
    class Tread1(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 0
                global bol1
                global items1
                bol1 = False
                items1, bol1 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items1)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread2(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 1
                global bol2
                global items2
                bol2 = False
                items2, bol2 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items2)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread3(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 2
                global bol3
                global items3
                bol3 = False
                items3, bol3 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items3)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread4(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 3
                global bol4
                global items4
                bol4 = False
                items4, bol4 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items4)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread5(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 4
                global bol5
                global items5
                bol5 = False
                items5, bol5 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items5)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread6(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 5
                global bol6
                global items6
                bol6 = False
                items6, bol6 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items6)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread7(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 6
                global bol7
                global items7
                bol7 = False
                items7, bol7 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items7)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread8(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 7
                global bol8
                global items8
                bol8 = False
                items8, bol8 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items8)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread9(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 8
                global bol9
                global items9
                bol9 = False
                items9, bol9 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items9)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread10(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 9
                global bol10
                global items10
                bol10 = False
                items10, bol10 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items10)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread11(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 10
                global bol11
                global items11
                bol11 = False
                items11, bol11 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items11)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread12(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 11
                global bol12
                global items12
                bol12 = False
                items12, bol12 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items12)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread13(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 12
                global bol13
                global items13
                bol13 = False
                items13, bol13 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items13)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread14(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 13
                global bol14
                global items14
                bol14 = False
                items14, bol14 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items14)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread15(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 14
                global bol15
                global items15
                bol15 = False
                items15, bol15 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items15)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread16(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 15
                global bol16
                global items16
                bol16 = False
                items16, bol16 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items16)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread17(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 16
                global bol17
                global items17
                bol17 = False
                items17, bol17 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items17)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread18(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 17
                global bol18
                global items18
                bol18 = False
                items18, bol18 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items18)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread19(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 18
                global bol19
                global items19
                bol19 = False
                items19, bol19 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items19)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

    # start 쓰레드 1(예매페이지 접속)
    class Tread20(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                global merges
                number = 19
                global bol20
                global items20
                bol20 = False
                items20, bol20 = requestCrawling(startNum[number], endNum[number], itemCodeTread[number], number)
                print(f"Tread{number + 1}")
                print(items20)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            while True:
                if True == bol1 == bol2 == bol3 == bol4 == bol5 == bol6 == bol7 == bol8 == bol9 == bol10 == bol11 == bol12 == bol13 == bol14 == bol15 == bol16 == bol17 == bol18 == bol19 == bol20:
                    merges = merge_two_dicts(items1, items2)
                    merges = merge_two_dicts(merges, items1)
                    merges = merge_two_dicts(merges, items2)
                    merges = merge_two_dicts(merges, items3)
                    merges = merge_two_dicts(merges, items4)
                    merges = merge_two_dicts(merges, items5)
                    merges = merge_two_dicts(merges, items6)
                    merges = merge_two_dicts(merges, items7)
                    merges = merge_two_dicts(merges, items8)
                    merges = merge_two_dicts(merges, items9)
                    merges = merge_two_dicts(merges, items10)
                    merges = merge_two_dicts(merges, items11)
                    merges = merge_two_dicts(merges, items12)
                    merges = merge_two_dicts(merges, items13)
                    merges = merge_two_dicts(merges, items14)
                    merges = merge_two_dicts(merges, items15)
                    merges = merge_two_dicts(merges, items16)
                    merges = merge_two_dicts(merges, items17)
                    merges = merge_two_dicts(merges, items18)
                    merges = merge_two_dicts(merges, items19)
                    merges = merge_two_dicts(merges, items20)
                    print(merges)
                    break
            end_time = time.time()
            textBrowser.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
            j = 1
            for key, value in merges.items():
                load_ws[f"B{j}"] = key
                load_ws[f"C{j}"] = value[0]
                load_ws[f"D{j}"] = value[1]
                load_ws[f"E{j}"] = value[2]
                load_ws[f"F{j}"] = value[3]
                load_ws[f"G{j}"] = value[4]
                load_ws[f"H{j}"] = value[5]
                j += 1

            load_wb.save(filename="CrawlingData.xlsx")

    # start 쓰레드 1(예매페이지 접속)
    class stop(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            sys.exit(0)

    # start 쓰레드 1(예매페이지 접속)
    class stop(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            sys.exit(0)

    # start
    def start(self):
        # 타이머 변수
        global start_time
        global end_time

        # Ui 변수
        global textBrowser

        # 데이터 변수
        global itemBacode
        global itemCodeTread
        global startNum
        global endNum


        try:
            # Ui 변수 선언
            Ui_startText = self.lineEdit_groupcode.text()
            Ui_endText = self.lineEdit_datecode.text()
            textBrowser = self.textBrowser

            # 2차원 쓰레드 20개 대응하는 리스트
            itemCodeTread = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]

            # item 10개의 리스트로 나누기
            StartListNum = int(Ui_startText.replace('A', ''))
            EndListNum = int(Ui_endText.replace('A', ''))

            # item 쓰레드 나눠서 담기
            count = EndListNum - StartListNum + 1
            divion = count // 20
            remainder = count % 20

            # startNum 리스트
            startNum = [StartListNum, StartListNum+(divion*1), StartListNum+(divion*2), StartListNum+(divion*3),
                        StartListNum+(divion*4), StartListNum+(divion*5), StartListNum+(divion*6),
                        StartListNum+(divion*7), StartListNum+(divion*8), StartListNum+(divion*9),
                        StartListNum+(divion*10), StartListNum+(divion*11), StartListNum+(divion*12),
                        StartListNum+(divion*13), StartListNum+(divion*14), StartListNum+(divion*15),
                        StartListNum+(divion*16), StartListNum+(divion*17), StartListNum+(divion*18),
                        StartListNum+(divion*19)]
            
            endNum = [StartListNum+(divion*1), StartListNum+(divion*2), StartListNum+(divion*3), 
                      StartListNum+(divion*4), StartListNum+(divion*5), StartListNum+(divion*6), 
                      StartListNum+(divion*7), StartListNum+(divion*8), StartListNum+(divion*9), 
                      StartListNum+(divion*10), StartListNum+(divion*11), StartListNum+(divion*12),
                      StartListNum+(divion*13), StartListNum+(divion*14), StartListNum+(divion*15),
                      StartListNum+(divion*16), StartListNum+(divion*17), StartListNum+(divion*18),
                      StartListNum+(divion*19), StartListNum+(divion*20)+remainder]

            # item 리스트 번호 부여 및 itemCodeTread 번호 담기
            itemDivionListStart = [Ui_startText, f'A{StartListNum + divion}', f'A{StartListNum + (divion * 2)}',
                                   f'A{StartListNum + (divion * 3)}', f'A{StartListNum + (divion * 4)}',
                                   f'A{StartListNum + (divion * 5)}', f'A{StartListNum + (divion * 6)}',
                                   f'A{StartListNum + (divion * 7)}', f'A{StartListNum + (divion * 8)}',
                                   f'A{StartListNum + (divion * 9)}', f'A{StartListNum + (divion * 10)}',
                                   f'A{StartListNum + (divion * 11)}', f'A{StartListNum + (divion * 12)}',
                                   f'A{StartListNum + (divion * 13)}', f'A{StartListNum + (divion * 14)}',
                                   f'A{StartListNum + (divion * 15)}', f'A{StartListNum + (divion * 16)}',
                                   f'A{StartListNum + (divion * 17)}', f'A{StartListNum + (divion * 18)}',
                                   f'A{StartListNum + (divion * 19)}'
                                   ]

            itemDivionListEnd = [f'A{StartListNum + divion-1}', f'A{StartListNum + (divion * 2)-1}',
                                 f'A{StartListNum + (divion * 3)-1}', f'A{StartListNum + (divion * 4)-1}',
                                 f'A{StartListNum + (divion * 5)-1}', f'A{StartListNum + (divion * 6)-1}',
                                 f'A{StartListNum + (divion * 7)-1}', f'A{StartListNum + (divion * 8)-1}',
                                 f'A{StartListNum + (divion * 9)-1}', f'A{StartListNum + (divion * 10)-1}',
                                 f'A{StartListNum + (divion * 11)-1}', f'A{StartListNum + (divion * 12)-1}',
                                 f'A{StartListNum + (divion * 13)-1}', f'A{StartListNum + (divion * 14)-1}',
                                 f'A{StartListNum + (divion * 15)-1}', f'A{StartListNum + (divion * 16)-1}',
                                 f'A{StartListNum + (divion * 17)-1}', f'A{StartListNum + (divion * 18)-1}',
                                 f'A{StartListNum + (divion * 19)-1}',
                                 f'A{StartListNum + (divion * 20) + remainder - 1}'
                                 ]

            for i in range(20):
                ListDivion(i, itemDivionListStart[i], itemDivionListEnd[i])


            start_time = time.time()
            Tread1 = Example.Tread1(self)
            Tread1.start()
            Tread2 = Example.Tread2(self)
            Tread2.start()
            Tread3 = Example.Tread3(self)
            Tread3.start()
            Tread4 = Example.Tread4(self)
            Tread4.start()
            Tread5 = Example.Tread5(self)
            Tread5.start()
            Tread6 = Example.Tread6(self)
            Tread6.start()
            Tread7 = Example.Tread7(self)
            Tread7.start()
            Tread8 = Example.Tread8(self)
            Tread8.start()
            Tread9 = Example.Tread9(self)
            Tread9.start()
            Tread10 = Example.Tread10(self)
            Tread10.start()
            Tread11 = Example.Tread11(self)
            Tread11.start()
            Tread12 = Example.Tread12(self)
            Tread12.start()
            Tread13 = Example.Tread13(self)
            Tread13.start()
            Tread14 = Example.Tread14(self)
            Tread14.start()
            Tread15 = Example.Tread15(self)
            Tread15.start()
            Tread16 = Example.Tread16(self)
            Tread16.start()
            Tread17 = Example.Tread17(self)
            Tread17.start()
            Tread18 = Example.Tread18(self)
            Tread18.start()
            Tread19 = Example.Tread19(self)
            Tread19.start()
            Tread20 = Example.Tread20(self)
            Tread20.start()

        except Exception as error:
            print(error)

    def stop(self):
        stop = Example.stop(self)
        stop.start()


app = QApplication([])
ex = Example()
sys.exit(app.exec_())



