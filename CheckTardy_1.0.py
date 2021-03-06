import sys
from datetime import datetime
import time
import chromedriver_autoinstaller
import pandas as pd

from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
from html_table_parser import parser_functions as parser


chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0] #크롬드라이버 버전 확인

options = webdriver.ChromeOptions()
options.add_argument("headless")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option("detach", True)

try:
    browser = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=options)
except:
    chromedriver_autoinstaller.install(True)
    browser = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=options)

browser.implicitly_wait(10)

def FindnClick_span_by_xpath(driver, xpath, text):
    for elem in driver.find_elements_by_xpath(xpath):
        if elem.text != "":
            print(elem.text)
        if elem.text == text:
            elem.click()
            break

def ParseTardyInfo(driver):
    try:
        time.sleep(1)
        result_html = driver.page_source
        result_soup = BeautifulSoup(result_html, 'html.parser')
        tags = result_soup.find_all("table", attrs={"class":"data data4 eff1"})[0]

        html_table = parser.make2d(tags)

        df = pd.DataFrame(html_table[1:], columns=html_table[0])
        #print(df.head())
        dtToday = str(int(datetime.today().strftime('%d')))
        strCol0 = "정렬변경"
        #with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
        #    print(df)
        #print(df['3'])
        for i in range(1, len(df[dtToday]), 3):
            if df[dtToday][i] == "00:00":
                if df[dtToday][i+2] == "출근전":
                    print(df[strCol0][i], df[dtToday][i+2], df[dtToday][i])
            elif df[dtToday][i] > "08:45":
                if df[dtToday][i+2] == "정상출근" or df[dtToday][i+2] == "출근전" or df[dtToday][i+2] == "지각":
                    print(df[strCol0][i], "지각", df[dtToday][i])
    except:
        print("Exception: ParseTardyInfo")

start_url = 'http://g.wisestone.kr'
browser.get(start_url)
browser.minimize_window()##maximize_window()
print("1: Start Chrome browser")


# 로그인 화면 인지 메인화면인지 구분
elemloginId = browser.find_element_by_id("emp_no")
if elemloginId != "":
    elemloginId.clear()
    elemloginId.send_keys("jjko")

    elemPass = browser.find_element_by_id("passwd")
    elemPass.clear()
    elemPass.send_keys("Wowjddl!@34")

    elemBtnLogin = browser.find_element_by_link_text("로그인")
    if elemBtnLogin != "":
        elemBtnLogin.click()
    else:
        print("ddd")
else:
    print("main")
print("2: Login Complete!")

try:
    element = browser.find_element_by_id("popupframe") #iframe 태그 엘리먼트 찾기
    if element != "":
        browser.switch_to.frame(element) #프레임 이동
        #browser.switch_to.frame("popupframe")
        time.sleep(1)
        elem0 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='no_today']")))
        if elem0 != "":
            elem0.click()
            browser.switch_to.parent_frame()
        time.sleep(1)
        print("2-1: End Popup")
except:
    browser.switch_to.parent_frame()

browser.switch_to.frame("contentsWrap")
time.sleep(1)
elem1 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@title='출·퇴근시간 자세히보기']")))
#for i in elem1:
#    print("="+i.text)
#    #attrs = browser.execute_script('var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;', i)
#    #print(attrs)
if elem1 != "":
    elem1.click()
browser.switch_to.parent_frame()
time.sleep(1)
print("3: Move Page")

# 근태관리 화면 > 소속출근기록 클릭
browser.switch_to.frame("leftmenuWrap")
time.sleep(1)
elem2 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@menu_cd='200147']")))
if elem2 != "":
    elem2.click()
browser.switch_to.parent_frame()
time.sleep(1)

# 월별 근태 테이블이 위치한 iframe의 src 값 추출
src1 = WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, "//div[@id='contents']/iframe[@src]"))).get_attribute("src")
print("4: Start Parsing tardy data...")

# 월별 근태 테이블 접근(1페이지)
browser.switch_to.frame("contentsWrap")
ParseTardyInfo(browser)

# 2페이지로 이동
elem3 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="container"]/div[2]/div[2]/a[1]""")))
elem3.click()
ParseTardyInfo(browser)

# 3페이지로 이동
elem4 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="container"]/div[2]/div[2]/a[3]""")))
elem4.click()
ParseTardyInfo(browser)
print("")

# 프로그램 종료 처리
browser.close()
browser.quit()
print("End check tardy")

while 1:
    a = input("종료를 위하여 q를 입력하세요.")
    if a == 'q':
        break

sys.exit()