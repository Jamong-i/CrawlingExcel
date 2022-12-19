from openpyxl import load_workbook
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt, QThread
from CrawlingUi import Ui_DomecallCrawling
from chrome_autoinstall import chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time

# 보안할 부분e
# 1. Key Value로 데이터 넣기
# 2. 쓰레드 10개로 탐색
# 2.5 쓰레드 올리지말고 크롬드라이버 생성 for문으로 100개 실행하고 i순서대로 값 받으면 됌e
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

def seleniumLogin(driver, textBrowser, num):
    while True:
        try:
            url = "https://www.domecall.net/member/login.php"
            driver.get(url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'input-info')))
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(num * 10)
            driver.find_element(By.XPATH, "/html/body/div[2]/div[3]/div/div/div/form/div[1]/div/div[1]/input").send_keys('userid')
            driver.find_element(By.XPATH, "/html/body/div[2]/div[3]/div/div/div/form/div[1]/div/div[2]/input").send_keys('userpwd' + Keys.RETURN)
            time.sleep(1)
            time.sleep(5)
            loginCheck = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[2]/div/div[2]/ul/li[2]/a").text
            print(loginCheck)
            if loginCheck == '로그아웃':
                textBrowser.append("로그인 완료")
                break
        except:
            print("로그인 오류")


def seleniumCrawling(driver, startNum, EndNum, itemList, number, textBrowser):
    j = 0
    items = {'num': ['productCode', 'price', 'bacode', 'bigBoxCount', 'smallBoxCount', 'origin']}
    startNumFix = startNum
    while True:
        try:
            for i in range(startNum, EndNum):
                j += 1
                url = f"https://www.domecall.net/goods/goods_view.php?goodsNo={itemList[i - startNumFix]}"
                driver.get(url)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                # price 찾기
                try:
                    price = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[1]/div/strong").text
                except:
                    price = "None"
                # bacode 찾기
                try:
                    bacode = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[2]/div").text
                except:
                    bacode = "None"
                # productCode 찾기
                try:
                    productCode = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[3]/div").text
                except:
                    productCode = "None"
                # origin 찾기
                try:
                    origin = driver.find_element(By.XPATH, "//html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[4]/div").text
                except:
                    origin = "None"
                # BigBoxCount 찾기
                try:
                    bigBoxText = driver.find_element(By.XPATH,
                                                     "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/strong").text
                    if bigBoxText == "대박스":
                        bigBoxCount = driver.find_element(By.XPATH,
                                                           "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                        try:
                            smallBoxCount = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                        except:
                            smallBoxCount = "None"

                    elif bigBoxText == "소박스":
                        smallBoxCount = driver.find_element(By.XPATH,
                                                           "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[5]/div/span").text
                        try:
                            bigBoxCount = driver.find_element(By.XPATH,
                                                                 "/html/body/div[2]/div[2]/div/div[1]/div[2]/form/div/div[2]/ul/li[6]/div/span").text
                        except:
                            bigBoxCount = "None"
                    else:
                        bigBoxCount = "None"
                        smallBoxCount = "None"
                except:
                    bigBoxCount = "None"
                    smallBoxCount = "None"
                items[i] = [productCode, price, bacode, bigBoxCount, smallBoxCount, origin]
                textBrowser.append(
                    f'쓰레드:{number}, {j}번 상품번호:{productCode}, 가격:{price}, 바코드번호:{bacode}, 큰박스:{bigBoxCount}, 작은박스:{smallBoxCount}')

            break

        except:
            try:
                productCode = 'None'
                price = 'None'
                bacode = 'None'
                bigBoxCount = 'None'
                smallBoxCount = 'None'
                origin = 'None'
                items[i] = [productCode, price, bacode, bigBoxCount, smallBoxCount, origin]
                textBrowser.append(
                    f'쓰레드:{number}, {j}번 상품번호:{productCode}, 가격:{price}, 바코드번호:{bacode}, 큰박스:{bigBoxCount}, 작은박스:{smallBoxCount}')
                startNum = i + 1
                i = startNum
            except:
                print("에러7")

    # 모두 끝남
    bol = True

    return items, bol


def merge_two_dicts(dicA, dicB):
    mergedic = dicA.copy()

    mergedic.update(dicB)

    return mergedic


def ListDivion(Num, Ui_startText, Ui_endText):
    # Excel item 불러 리스트 담기
    get_cells = load_ws[Ui_startText:Ui_endText]
    for row in get_cells:
        for cell in row:
            itemCodeTread[Num].append(cell.value)
    for i in range(10):
        print(f"{Num}번 {itemCodeTread[i]}")

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
                driver = driver1
                textBrowser = textBrowser1
                global bol1
                global items1
                bol1 = False
                seleniumLogin(driver, textBrowser, number)
                items1, bol1 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items1)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 1번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread2(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 1
                driver = driver2
                textBrowser = textBrowser2
                global bol2
                global items2
                bol2 = False
                seleniumLogin(driver, textBrowser, number)
                items2, bol2 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items2)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 2번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread3(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 2
                driver = driver3
                textBrowser = textBrowser3
                global bol3
                global items3
                bol3 = False
                seleniumLogin(driver, textBrowser, number)
                items3, bol3 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items3)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 3번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread4(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 3
                driver = driver4
                textBrowser = textBrowser4
                global bol4
                global items4
                bol4 = False
                seleniumLogin(driver, textBrowser, number)
                items4, bol4 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items4)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 4번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread5(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 4
                driver = driver5
                textBrowser = textBrowser5
                global bol5
                global items5
                bol5 = False
                seleniumLogin(driver, textBrowser, number)
                items5, bol5 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items5)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 5번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread6(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 5
                driver = driver6
                textBrowser = textBrowser6
                global bol6
                global items6
                bol6 = False
                seleniumLogin(driver, textBrowser, number)
                items6, bol6 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items6)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 6번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread7(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 6
                driver = driver7
                textBrowser = textBrowser7
                global bol7
                global items7
                bol7 = False
                seleniumLogin(driver, textBrowser, number)
                items7, bol7 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items7)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 7번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread8(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 7
                driver = driver8
                textBrowser = textBrowser8
                global bol8
                global items8
                bol8 = False
                seleniumLogin(driver, textBrowser, number)
                items8, bol8 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items8)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 8번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread9(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 8
                driver = driver9
                textBrowser = textBrowser9
                global bol9
                global items9
                bol9 = False
                seleniumLogin(driver, textBrowser, number)
                items9, bol9 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items9)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 9번 크롬 종료")

    # start 쓰레드 1(예매페이지 접속)
    class Tread10(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                number = 9
                driver = driver10
                textBrowser = textBrowser10
                global bol10
                global items10
                bol10 = False
                seleniumLogin(driver, textBrowser, number)
                items10, bol10 = seleniumCrawling(driver, startNum[number], endNum[number], itemCodeTread[number], number, textBrowser)
                print(f"Tread{number + 1}")
                print(items10)
            except Exception as error:
                print(f"Tread{number+1} error : {error}")

            driver.quit()
            textBrowser.append("쓰레드 10번 크롬 종료")

            while True:
                print(bol1, bol2, bol3, bol4, bol5, bol6, bol7, bol8, bol9, bol10)
                if True == bol1 == bol2 == bol3 == bol4 == bol5 == bol6 == bol7 == bol8 == bol9 == bol10:
                    merges = merge_two_dicts(items1, items2)
                    merges = merge_two_dicts(merges, items3)
                    merges = merge_two_dicts(merges, items4)
                    merges = merge_two_dicts(merges, items5)
                    merges = merge_two_dicts(merges, items6)
                    merges = merge_two_dicts(merges, items7)
                    merges = merge_two_dicts(merges, items8)
                    merges = merge_two_dicts(merges, items9)
                    merges = merge_two_dicts(merges, items10)
                    print(merges)
                    break
            end_time = time.time()
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

    # start
    def start(self):
        # 타이머 변수
        global start_time
        global end_time

        # Ui 변수
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
        global itemBacode
        global itemCodeTread
        global startNum
        global endNum

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

            # 2차원 쓰레드 20개 대응하는 리스트
            itemCodeTread = [[], [], [], [], [], [], [], [], [], []]

            # item 10개의 리스트로 나누기
            StartListNum = int(Ui_startText.replace('A', ''))
            EndListNum = int(Ui_endText.replace('A', ''))

            # item 쓰레드 나눠서 담기
            count = EndListNum - StartListNum + 1
            divion = count // 10
            remainder = count % 10

            # startNum 리스트
            startNum = [StartListNum, StartListNum + (divion * 1), StartListNum + (divion * 2),
                        StartListNum + (divion * 3),
                        StartListNum + (divion * 4), StartListNum + (divion * 5), StartListNum + (divion * 6),
                        StartListNum + (divion * 7), StartListNum + (divion * 8), StartListNum + (divion * 9)]

            endNum = [StartListNum + (divion * 1), StartListNum + (divion * 2), StartListNum + (divion * 3),
                      StartListNum + (divion * 4), StartListNum + (divion * 5), StartListNum + (divion * 6),
                      StartListNum + (divion * 7), StartListNum + (divion * 8), StartListNum + (divion * 9),
                      StartListNum + (divion * 10) + remainder]

            # item 리스트 번호 부여 및 itemCodeTread 번호 담기
            itemDivionListStart = [Ui_startText, f'A{StartListNum + divion}', f'A{StartListNum + (divion * 2)}',
                                   f'A{StartListNum + (divion * 3)}', f'A{StartListNum + (divion * 4)}',
                                   f'A{StartListNum + (divion * 5)}', f'A{StartListNum + (divion * 6)}',
                                   f'A{StartListNum + (divion * 7)}', f'A{StartListNum + (divion * 8)}',
                                   f'A{StartListNum + (divion * 9)}'
                                   ]

            itemDivionListEnd = [f'A{StartListNum + divion - 1}', f'A{StartListNum + (divion * 2) - 1}',
                                 f'A{StartListNum + (divion * 3) - 1}', f'A{StartListNum + (divion * 4) - 1}',
                                 f'A{StartListNum + (divion * 5) - 1}', f'A{StartListNum + (divion * 6) - 1}',
                                 f'A{StartListNum + (divion * 7) - 1}', f'A{StartListNum + (divion * 8) - 1}',
                                 f'A{StartListNum + (divion * 9) - 1}', f'A{StartListNum + (divion * 10) + remainder - 1}'
                                 ]

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
