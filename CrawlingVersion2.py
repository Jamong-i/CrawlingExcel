import logging
from openpyxl import load_workbook
import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt, QThread
from CrawlingUi import Ui_MainWindow
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


#data_only=Ture로 해줘야 수식이 아닌 값으로 받아온다.
load_wb = load_workbook("data.xlsx", data_only=True)

#시트 이름으로 불러오기
load_ws = load_wb['Sheet1']

# chrome_autoinstaller
driver = chromedriver()

class Example(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        # 프로그램이 항상 최상단에 위치하도록 지정(크롬에 가리지 않게)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.show()


    def Crawling(count):
        data = [count]
        return 0

    # start 쓰레드 1(예매페이지 접속)
    class start_clock(QThread):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def run(self):
            try:
                global n
                start_time = time.time()
                data = []
                data_bacode = []
                n = 0
                data_bacode
                get_cells = load_ws[Ui_startText:Ui_endText]
                for row in get_cells:
                    for cell in row:
                        data.append(cell.value)

                # String A없애주기
                new_str = Ui_startText.replace('A', '')
                new_end = Ui_endText.replace('A', '')
                while True:
                    try:
                        for i in range(n, int(new_end) - int(new_str) + 1):
                            driver.get(f'https://www.domecall.net/goods/goods_view.php?goodsNo={data[i]}')
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'item')))
                            bacode_XPATH = driver.find_element(By.XPATH, "/ html / body / div[2] / div[2] / div / div[1] / div[2] / form / div / div[2] / ul / li[2] / div")
                            bacode = bacode_XPATH.text
                            data_bacode.append(bacode)
                            textBrowser.append(f'{i+1}번 상품번호:{data[i]}, 바코드번호:{bacode}')
                    except Exception as error:
                        print(error, data[i])
                        data_bacode.append("구매 불가")
                        n = i + 1

                    print(data_bacode)

                    for j in range(0, int(new_end) - int(new_str) + 1):
                        load_ws[f"B{j+2}"] = data_bacode[j]

                    load_wb.save(filename="data_bacode.xlsx")
                    textBrowser.append("엑셀 저장 완료")
                    end_time = time.time()
                    textBrowser.append(f"{end_time - start_time:.5f} sec")
                    break

            except Exception as error:
                print(error)

    # start
    def start(self):
        global start_time
        global end_time
        global Ui_startText
        global Ui_endText
        global textBrowser

        try:
            Ui_startText = self.lineEdit_groupcode.text()
            Ui_endText = self.lineEdit_datecode.text()
            textBrowser = self.textBrowser

            start_clock = Example.start_clock(self)
            start_clock.start()
        except Exception as error:
            print(error)

    def stop(self):
        driver.quit()
        sys.exit(0)


app = QApplication([])
ex = Example()
sys.exit(app.exec_())
