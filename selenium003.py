import datetime as dt
import time

from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

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

print("find elemLogin-Start")
time.sleep(1)

for elemLogin in browser.find_elements_by_xpath('.//span[@class="ng-scope"]'):
    print(elemLogin.text)
    if elemLogin.text == "로그인":
        elemLogin.click()
        break

browser.implicitly_wait(1)

print("find elemIssueMgmt-Start")
time.sleep(1)

for elemIssueMgmt in browser.find_elements_by_xpath('.//span[@class="ng-scope"]'):
    print(elemIssueMgmt.text)
    if elemIssueMgmt.text == "이슈 관리":
        elemIssueMgmt.click()
        break

#issue_url = 'http://www.owlsolution.co.kr:8080/#/issues/issueList/'
#browser.get(issue_url)

#browser.find_element_by_xpath("//span[contains(@class, 'ng-scope')]").click()

#browser.find_element_by_class_name("ng-scope").click()

time.sleep(1)
