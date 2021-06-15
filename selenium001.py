from selenium import webdriver
import time

def open_chrome_driver():
    path = "D:\SRC\GitHub\selenium\chromedriver.exe"
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("detach", True)
    chrome_driver = webdriver.Chrome(path, options=options)
    return chrome_driver

driver = open_chrome_driver()

#driver = webdriver.Chrome(path)
#driver.get(url)
#browser = webdriver.Chrome(path, options=options)
#url = 'https://www.google.com'
#browser.get(url)

time.sleep(30)
driver.close()