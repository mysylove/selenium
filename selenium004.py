import datetime as dt
import time
import chromedriver_autoinstaller

from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0] #크롬드라이버 버전 확인

options = webdriver.ChromeOptions()
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

#delay = 0.1
#path = "D:\SRC\GitHub\selenium\chromedriver.exe"
#options = webdriver.ChromeOptions()
#options.add_experimental_option("excludeSwitches", ["enable-logging"])
#options.add_experimental_option("detach", True)
#browser = Chrome(path, options=options)
#browser.implicitly_wait(delay)

start_url = 'http://g.wisestone.kr'
browser.get(start_url)
browser.maximize_window()


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


browser.switch_to.frame("contentsWrap")
time.sleep(1)
wait = WebDriverWait(browser, 10)
elem1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@title='출·퇴근시간 자세히보기']")))
#for i in elem1:
#    print("="+i.text)
#    #attrs = browser.execute_script('var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;', i)
#    #print(attrs)
if elem1 != "":
    print("ttttttt")
    elem1.click()
browser.switch_to.parent_frame()
time.sleep(1)

# 근태관리 화면 > 소속출근기록 클릭
browser.switch_to.frame("leftmenuWrap")
time.sleep(1)
wait = WebDriverWait(browser, 10)
elem2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@menu_cd='200147']")))
if elem2 != "":
    print("gggggg")
    elem2.click()
browser.switch_to.parent_frame()
time.sleep(1)




time.sleep(10)

'''
elem1 = browser.find_element_by_id("loginForm1")
elem1.clear()
elem1.send_keys("jjko@wisestone.kr")

elem2 = browser.find_element_by_id("loginForm2")
elem2.clear()
elem2.send_keys("wowjddl!@34")

time.sleep(1)
print("find 로그인-Start")
FindnClick_span_by_xpath(browser, './/span[@class="ng-scope"]', '로그인')
browser.implicitly_wait(2)

time.sleep(1)
print("find 이슈 관리-Start")
FindnClick_span_by_xpath(browser, './/span[@class="ng-scope"]', '이슈 관리')
#browser.implicitly_wait(2)

time.sleep(1)
print("find 상세검색-Start")
FindnClick_span_by_xpath(browser, './/span[@class="ng-scope"]', '상세검색')
#browser.implicitly_wait(2)

time.sleep(3)
'''