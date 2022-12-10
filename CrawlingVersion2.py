import logging
from openpyxl import load_workbook
import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt, QThread
from CrawlingUi import Ui_DomecallCrawling
from chrome_autoinstall import chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import math
import time

# 보안할 부분
# 1. Key Value로 데이터 넣기
# 2. 쓰레드 10개로 탐색
# 2.5 쓰레드 올리지말고 크롬드라이버 생성 for문으로 100개 실행하고 i순서대로 값 받으면 됌
# 3 소박스 대박스 try except문으로 있으면 실행 없으면 없음 키 벨류에 넣기


# data_only=Ture로 해줘야 수식이 아닌 값으로 받아온다.
load_wb = load_workbook("data.xlsx", data_only=True)

# 시트 이름으로 불러오기
load_ws = load_wb['Sheet1']

# chrome_autoinstaller
driver1 = chromedriver()
driver2 = chromedriver()
driver3 = chromedriver()
driver4 = chromedriver()
driver5 = chromedriver()
driver6 = chromedriver()
driver7 = chromedriver()
driver8 = chromedriver()
driver9 = chromedriver()
driver10 = chromedriver()


def ListDivion(Num, Ui_startText, Ui_endText):
    # Excel item 불러 리스트 담기
    get_cells = load_ws[Ui_startText:Ui_endText]
    for row in get_cells:
        for cell in row:
            itemCodeTread[Num].append(cell.value)


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
                treadN = 0
                number = 0
                driver = driver1
                textBrowser = textBrowser1

                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF1
                    TreadTF1 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread2(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 1
                driver = driver2
                textBrowser = textBrowser2
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF2
                    TreadTF2 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread3(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 2
                driver = driver3
                textBrowser = textBrowser3
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF3
                    TreadTF3 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread4(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 3
                driver = driver4
                textBrowser = textBrowser4
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF4
                    TreadTF4 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread5(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 4
                driver = driver5
                textBrowser = textBrowser5
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF5
                    TreadTF5 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread6(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 5
                driver = driver6
                textBrowser = textBrowser6
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF6
                    TreadTF6 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread7(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 6
                driver = driver7
                textBrowser = textBrowser7
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF7
                    TreadTF7 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread8(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 7
                driver = driver8
                textBrowser = textBrowser8
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF8
                    TreadTF8 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread9(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 8
                driver = driver9
                textBrowser = textBrowser9
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    global TreadTF9
                    TreadTF9 = True
                    break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class Tread10(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                treadN = 0
                number = 9
                driver = driver10
                textBrowser = textBrowser10
                while True:
                    try:
                        for i in range(treadN, len(itemCodeTread[number])):
                            driver.get(
                                f'https://www.domecall.net/goods/goods_view.php?goodsNo={itemCodeTread[number][i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH,
                                                               "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            itemBacode[number].append(bacode)
                            try:
                                bigBoxText = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                                if bigBoxText == "대박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    bigBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        smallBox[number].append(smallBox_XPATH)
                                    except:
                                        smallBox[number].append("정보 없음")

                                elif bigBoxText == "소박스":
                                    bigBox_XPATH = driver.find_element(By.XPATH,
                                                                       "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                                    smallBox[number].append(bigBox_XPATH)
                                    try:
                                        smallBox_XPATH = driver.find_element(By.XPATH,
                                                                             "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                                        bigBox[number].append(smallBox_XPATH)
                                    except:
                                        bigBox[number].append("정보 없음")
                            except:
                                bigBox[number].append("정보 없음")
                                smallBox[number].append("정보 없음")

                            textBrowser.append(
                                f'{i + 1}번 상품번호:{itemCodeTread[number][i]}, 바코드번호:{itemBacode[number][i]}, 큰박스:{bigBox[number][i]}, 작은박스:{smallBox[number][i]}')
                    except:
                        print(f"스레드 1번 {itemCodeTread[number][i]} 구매불가 상품")
                        itemBacode[number].append("구매 불가")
                        smallBox[number].append("구매 불가")
                        bigBox[number].append("구매 불가")
                        treadN = i + 1
                    break

                while True:
                    if True == TreadTF1 == TreadTF2 == TreadTF3 == TreadTF4 == TreadTF5 == TreadTF6 == TreadTF7 == TreadTF8 == TreadTF9:
                        end_time = time.time()
                        totalitemCode = itemCodeTread[0] + itemCodeTread[1] + itemCodeTread[2] + itemCodeTread[3] + \
                                        itemCodeTread[4] \
                                        + itemCodeTread[5] + itemCodeTread[6] + itemCodeTread[7] + itemCodeTread[8] + \
                                        itemCodeTread[9]
                        totalBacode = itemBacode[0] + itemBacode[1] + itemBacode[2] + itemBacode[3] + itemBacode[4] + \
                                      itemBacode[5] \
                                      + itemBacode[6] + itemBacode[7] + itemBacode[8] + itemBacode[9]
                        totalSmallBox = smallBox[0] + smallBox[1] + smallBox[2] + smallBox[3] + smallBox[4] + smallBox[
                            5] + \
                                        smallBox[6] \
                                        + smallBox[7] + smallBox[8] + smallBox[9]
                        totalbigBox = bigBox[0] + bigBox[1] + bigBox[2] + bigBox[3] + bigBox[4] + bigBox[5] + bigBox[
                            6] + \
                                      bigBox[7] \
                                      + bigBox[8] + bigBox[9]

                        for j in range(len(totalitemCode)):
                            load_ws[f"B{j + 2}"] = totalitemCode[j]
                        for j in range(len(totalBacode)):
                            load_ws[f"C{j + 2}"] = totalBacode[j]
                        for j in range(len(totalSmallBox)):
                            load_ws[f"D{j + 2}"] = totalSmallBox[j]
                        for j in range(len(totalbigBox)):
                            load_ws[f"E{j + 2}"] = totalbigBox[j]

                        load_wb.save(filename="data_bacode.xlsx")
                        textBrowser1.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser2.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser3.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser4.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser5.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser6.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser7.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser8.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser9.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        textBrowser10.append(f"엑셀 저장 완료 {end_time - start_time:.5f} sec")
                        break
            except Exception as error:
                print(error)

    # start 쓰레드 1(예매페이지 접속)
    class stop(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            driver1.quit()
            driver2.quit()
            driver3.quit()
            driver4.quit()
            driver5.quit()
            driver6.quit()
            driver7.quit()
            driver8.quit()
            driver9.quit()
            driver10.quit()
            sys.exit(0)

    # start
    def start(self):
        # 타이머 변수
        global start_time
        global end_time
        # Ui 변수
        global Ui_startText
        global Ui_endText
        global textBrowser1
        global textBrowser2
        global textBrowser3
        global textBrowser4
        global textBrowser5
        global textBrowser6
        global textBrowser7
        global textBrowser8
        global textBrowser9
        global textBrowser10
        # 데이터 변수
        global itemCode
        global itemBacode
        global itemCodeTread
        global smallBox
        global bigBox

        try:
            # Ui 변수 선언
            Ui_startText = self.lineEdit_groupcode.text()
            Ui_endText = self.lineEdit_datecode.text()
            textBrowser1 = self.textBrowser_1
            textBrowser2 = self.textBrowser_2
            textBrowser3 = self.textBrowser_3
            textBrowser4 = self.textBrowser_4
            textBrowser5 = self.textBrowser_5
            textBrowser6 = self.textBrowser_6
            textBrowser7 = self.textBrowser_7
            textBrowser8 = self.textBrowser_8
            textBrowser9 = self.textBrowser_9
            textBrowser10 = self.textBrowser_10

            # 2차원 쓰레드 10개 대응하는 리스트
            itemCodeTread = [[], [], [], [], [], [], [], [], [], []]
            itemBacode = [[], [], [], [], [], [], [], [], [], []]
            smallBox = [[], [], [], [], [], [], [], [], [], []]
            bigBox = [[], [], [], [], [], [], [], [], [], []]

            # item 10개의 리스트로 나누기
            StartListNum = int(Ui_startText.replace('A', ''))
            EndListNum = int(Ui_endText.replace('A', ''))

            # item 쓰레드 나눠서 담기
            count = EndListNum - StartListNum
            divion = count // 10
            remainder = count % 10

            # item 리스트 번호 부여 및 itemCodeTread 번호 담기
            itemDivionListStart = [Ui_startText, f'A{StartListNum + divion + 1}', f'A{StartListNum + (divion * 2) + 1}',
                                   f'A{StartListNum + (divion * 3) + 1}', f'A{StartListNum + (divion * 4) + 1}',
                                   f'A{StartListNum + (divion * 5) + 1}', f'A{StartListNum + (divion * 6) + 1}',
                                   f'A{StartListNum + (divion * 7) + 1}', f'A{StartListNum + (divion * 8) + 1}',
                                   f'A{StartListNum + (divion * 9) + 1}']

            itemDivionListEnd = [f'A{StartListNum + divion}', f'A{StartListNum + (divion * 2)}',
                                 f'A{StartListNum + (divion * 3)}', f'A{StartListNum + (divion * 4)}',
                                 f'A{StartListNum + (divion * 5)}', f'A{StartListNum + (divion * 6)}',
                                 f'A{StartListNum + (divion * 7)}', f'A{StartListNum + (divion * 8)}',
                                 f'A{StartListNum + (divion * 9)}', f'A{StartListNum + (divion * 10) + remainder}']

            for i in range(10):
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

        except Exception as error:
            print(error)

    def stop(self):
        stop = Example.stop(self)
        stop.start()


app = QApplication([])
ex = Example()
sys.exit(app.exec_())
