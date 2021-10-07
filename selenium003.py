import datetime as dt
import time
import chromedriver_autoinstaller

from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0] #크롬드라이버 버전 확인

try:
    driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe')
except:
    chromedriver_autoinstaller.install(True)
    driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe')

driver.implicitly_wait(10)

def FindnClick_span_by_xpath(driver, xpath, text):
    for elem in driver.find_elements_by_xpath(xpath):
        if elem.text != "":
            print(elem.text)
        if elem.text == text:
            elem.click()
            break

delay = 0.1
path = "D:\SRC\GitHub\selenium\chromedriver.exe"
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option("detach", True)
browser = Chrome(path, options=options)
browser.implicitly_wait(delay)

start_url = 'http://www.owlsolution.co.kr:8080'
browser.get(start_url)
browser.maximize_window()

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
