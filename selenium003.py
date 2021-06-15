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

body = browser.find_element_by_tag_name('body')
time.sleep(5)
