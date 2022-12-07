import chromedriver_autoinstaller
from selenium import webdriver


def chromedriver():
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]  # 크롬드라이버 버전 확인

    try:
        # options = webdriver.ChromeOptions()
        # # 창 숨기는 옵션 추가
        # options.add_argument("headless")
        # options.add_argument('window-size=1920x1080')
        # options.add_argument("disable-gpu")

        chromedriver_autoinstaller.install(True)
        # driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver', chrome_options=options)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver')
        return driver
    except:
        # options = webdriver.ChromeOptions()
        # # 창 숨기는 옵션 추가
        # options.add_argument("headless")
        # options.add_argument('window-size=1920x1080')
        # options.add_argument("disable-gpu")

        chromedriver_autoinstaller.install(True)
        # driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver', chrome_options=options)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver')
        return driver

