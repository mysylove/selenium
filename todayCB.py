import sys
import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2

from datetime import datetime
import clipboard
import time
import pyautogui
import pywinauto
import pygetwindow as gw
import openpyxl as xl
from PyQt5.QtWidgets import * #QApplication, QWidget, QPushButton
from PyQt5.QtCore import QCoreApplication

##################################################
# CheckTardy(Selenium 관련)
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
##################################################

'''
# 유저가 현재 보고 있는 윈도우창 알아내기
import time, ctypes

time.sleep(3)    # 타이틀을 가져오고자 하는 윈도우를 활성화 하기위해 의도적으로 3초 멈춥니다. 
lib = ctypes.windll.LoadLibrary('user32.dll')
handle = lib.GetForegroundWindow()    # 활성화된 윈도우의 핸들얻음
buffer = ctypes.create_unicode_buffer(255)    # 타이틀을 저장할 버퍼
lib.GetWindowTextW(handle, buffer, ctypes.sizeof(buffer))    # 버퍼에 타이틀 저장

print(buffer.value)    # 버퍼출력
'''

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

        #self.btn_1.clicked.connect(self.button1Function)

    def initUI(self):
        self.btn_1 = QPushButton('Only Today Write', self)
        self.btn_1.move(10, 10)
        self.btn_1.resize(self.btn_1.sizeHint())
        #self.btn_1.clicked.connect(QCoreApplication.instance().quit)
        self.btn_1.clicked.connect(self.button1Function)

        self.label_1 = QLabel("클릭 후 OWL 내용 입력 창을 선택하면 [오늘날짜 이름]을 최상단에 입력", self)
        self.label_1.move(150, 15)

        self.btn_2 = QPushButton('Today+Clipboard Write', self)
        self.btn_2.move(10, 40)
        self.btn_2.resize(self.btn_2.sizeHint())
        self.btn_2.clicked.connect(self.button2Function)

        self.label_2 = QLabel("(미리 메일 내용을 복사하고)클릭 후 OWL 내용 입력 창을 선택하면 [오늘날짜 이름]과 메일내용을 최상단에 입력", self)
        self.label_2.move(150, 45)

        self.btn_3 = QPushButton('견적서 EXCEL 파일 분석', self)
        self.btn_3.move(10, 70)
        self.btn_3.resize(self.btn_3.sizeHint())
        self.btn_3.clicked.connect(self.button3Function)

        self.label_3 = QLabel("", self)
        self.label_3.move(170, 75)
        self.label_3.resize(300, 15)

        self.btn_4 = QPushButton('TEST', self)
        self.btn_4.move(10, 100)
        self.btn_4.resize(self.btn_4.sizeHint())
        self.btn_4.clicked.connect(self.button4Function)

        self.btn_5 = QPushButton('CheckTardy', self)
        self.btn_5.move(10, 130)
        self.btn_5.resize(self.btn_5.sizeHint())
        self.btn_5.clicked.connect(self.button5Function)

        self.text_edit_1 = QTextEdit(self)
        self.text_edit_1.setGeometry(200, 70, 500, 100)

        self.setWindowTitle('OWL Semi-Auto Write')
        self.setGeometry(300, 700, 900, 300)
        self.show()
    
    def button1Function(self):
        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 

        self.MySetFocusWindowWithTitle('OWL ITS - Chrome')

        time.sleep(1)
        pyautogui.hotkey('ctrl', 'home')
        pyautogui.hotkey('ctrl', 'enter')
        pyautogui.hotkey('ctrl', 'home')
        clipboard.copy(strTodayOWL)
        
        pyautogui.hotkey('ctrl', 'v')

    def button2Function(self):
        self.label_2.setText("(미리 메일 내용을 복사하고)클릭 후 OWL 내용 입력 창을 선택하면 [오늘날짜 이름]과 메일내용을 최상단에 입력")
        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 

        self.MySetFocusWindowWithTitle('OWL ITS - Chrome')

        time.sleep(1)
        strPaste = clipboard.paste()
        if len(strPaste) != 0:
            pyautogui.hotkey('ctrl', 'home')
            pyautogui.hotkey('ctrl', 'enter')
            pyautogui.hotkey('ctrl', 'home')
            pyautogui.hotkey('ctrl', 'v')

            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'home')
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl', 'home')
            clipboard.copy(strTodayOWL)
            pyautogui.hotkey('ctrl', 'v')
        else:
            self.label_2.setText("클립보드에 내용이 없습니다.")

    def button3Function(self):
        self.label_3.setText("")
        
        xl = clipboard.paste()
        strFind = '시험 제품명' #'수  신'
        if xl.find(strFind) == -1:
            self.label_3.setText("클립보드에 견적내용이 없습니다.")
            return
        '''
        pos = xl.find(strFind) + len(strFind)
        while xl[pos].isspace():
            pos = pos + 1
        nPosStart = pos
        while not xl[pos].isspace():
            pos = pos + 1
        nPosEnd = pos
        strCompany = xl[nPosStart:nPosEnd] # 수신 회사명
        
        self.MySetFocusWindowWithTitle(strCompany)

        time.sleep(1)
        pyautogui.hotkey('ctrl', 's')
        pyautogui.hotkey('alt', 'f')
        pyautogui.hotkey('alt', 'e')
        pyautogui.hotkey('alt', 'a')
        pyautogui.press('enter')
        '''
        strFind = '직접인건비'
        pos = xl.find(strFind) + len(strFind)
        while xl[pos].isspace():
            pos = pos + 1
        nPosStart = pos
        while xl[pos].isdigit() or xl[pos] == "." :
            pos = pos + 1
        nPosEnd = pos
        strTestDays = xl[nPosStart:nPosEnd]

        strFind = '할인율'
        pos = xl.find(strFind) + len(strFind)
        while not xl[pos].isdigit():
            pos = pos + 1
        nPosStart = pos
        while xl[pos].isdigit() or xl[pos] == "." or xl[pos] == "%":
            pos = pos + 1
        nPosEnd = pos
        strDiscount = xl[nPosStart:nPosEnd]

        pos = -1
        while xl[pos].isspace():
            pos = pos - 1
        nPosStart = pos + 1
        while xl[pos].isdigit() or xl[pos] == ",":
            pos = pos - 1
        nPosEnd = pos + 1
        strTotalMoney = xl[nPosEnd:nPosStart]

        print('시험소요일:', strTestDays)
        print('할인율:', strDiscount)

        strFind = '(부가세 포함)'
        if xl.find(strFind) == -1:
            strTotalMoney = strTotalMoney + '원'
            print('시험견적가:', strTotalMoney)
        else:
            strTotalMoney = strTotalMoney + '원' + strFind
            print('시험견적가:', strTotalMoney)

        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 
        strEst = strTodayOWL + ' 견적 발행' + '\n' + 'ㅇ ' + strTestDays + '일 ' + strDiscount + ' 할인 적용 ' + strTotalMoney + '; 가일정 : '

        self.MySetFocusWindowWithTitle('OWL ITS - Chrome')

        time.sleep(1)
        pyautogui.hotkey('ctrl', 'home')
        pyautogui.hotkey('ctrl', 'enter')
        pyautogui.hotkey('ctrl', 'home')
        clipboard.copy(strEst)
        pyautogui.hotkey('ctrl', 'v') 

    def button4Function(self):
        fname = QFileDialog.getOpenFileName(self, 'Open File', '',
                                        'All File(*);; html File(*.html *.htm)')
        if fname[0]:
            # 튜플 데이터에서 첫 번째 인자 값이 주소이다.
            self.pathLabel.setText(fname[0])
            print('filepath : ', fname[0])
            print('filesort : ', fname[1])

            # 텍스트 파일 내용 읽기
            f = open(fname[0], 'r', encoding='UTF8') # Path 정보로
            with f:
                data = f.read()
                self.textEdit.setText(data)
        else:
            QMessageBox.about(self, 'Warning', '파일을 선택하지 않았습니다.')

    def button5Function(self):
        self.text_edit_1.setText("")
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

        start_url = 'http://g.wisestone.kr'
        browser.get(start_url)
        browser.minimize_window()##maximize_window()
        self.text_edit_1.append("1: Start Chrome browser") #print("1: Start Chrome browser")


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
        self.text_edit_1.append("2: Login Complete!") #print("2: Login Complete!")

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
        self.text_edit_1.append("3: Move Page") #print("3: Move Page")

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
        self.text_edit_1.append("4: Start Parsing tardy data...") #print("4: Start Parsing tardy data...")

        # 월별 근태 테이블 접근(1페이지)
        browser.switch_to.frame("contentsWrap")
        self.ParseTardyInfo(browser)

        # 2페이지로 이동
        elem3 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="container"]/div[2]/div[2]/a[1]""")))
        elem3.click()
        self.ParseTardyInfo(browser)

        # 3페이지로 이동
        elem4 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="container"]/div[2]/div[2]/a[3]""")))
        elem4.click()
        self.ParseTardyInfo(browser)

        # 프로그램 종료 처리
        browser.close()
        browser.quit()        
        self.text_edit_1.append("End check tardy") #print("End check tardy")

    def MySetFocusWindowWithTitle(self, title):
        '''
        titles = gw.getWindowsWithTitle('')
        for x in titles:
            print(x)
        '''
        #win = gw.getWindowsWithTitle('OWL ITS - Chrome')[0] # 윈도우 타이틀에 Chrome 이 포함된 모든 윈도우 수집, 리스트로 리턴
        win = gw.getWindowsWithTitle(title)[0] # 윈도우 타이틀에 Chrome 이 포함된 모든 윈도우 수집, 리스트로 리턴
        if win.isActive == False:
            pywinauto.application.Application().connect(handle=win._hWnd).top_window().set_focus()
        win.activate() #윈도우 활성화

    ##################################################
    # CheckTardy(Selenium 관련)
    def FindnClick_span_by_xpath(self, driver, xpath, text):
        for elem in driver.find_elements_by_xpath(xpath):
            if elem.text != "":
                print(elem.text)
            if elem.text == text:
                elem.click()
                break

    def ParseTardyInfo(self, driver):
        try:
            time.sleep(1)
            result_html = driver.page_source
            result_soup = BeautifulSoup(result_html, 'html.parser')
            tags = result_soup.find_all("table", attrs={"class":"data data4 eff1"})[0]

            html_table = parser.make2d(tags)
            strLog = ""

            df = pd.DataFrame(html_table[1:], columns=html_table[0])
            #print(df.head())
            dtToday = str(int(datetime.today().strftime('%d')))
            strCol0 = "정렬변경"
            #print(df[dtToday])
            for i in range(1, len(df[dtToday]), 3):
                if df[dtToday][i] == "00:00":
                    if df[dtToday][i+2] == "출근전":
                        strLog = df[strCol0][i] + " " + df[dtToday][i+2] + " " + df[dtToday][i]
                        self.text_edit_1.append(strLog) #print(df[strCol0][i], df[dtToday][i+2], df[dtToday][i])
                elif df[dtToday][i] > "08:45":
                    if df[dtToday][i+2] == "정상출근" or df[dtToday][i+2] == "출근전":
                        strLog = df[strCol0][i] + " " + "지각" + " " + df[dtToday][i]
                        self.text_edit_1.append(strLog) #print(df[strCol0][i], "지각", df[dtToday][i])
        except:
            self.text_edit_1.append("Exception: ParseTardyInfo") #print("Exception: ParseTardyInfo")
    ##################################################

######################################################################################
# pyqt5 에서 exception 발생시 종료 방지 방법(https://4uwingnet.tistory.com/13)
def my_exception_hook(exctype, value, traceback):
    # Print the error and traceback
    print(exctype, value, traceback)
    # Call the normal Exception hook after
    sys._excepthook(exctype, value, traceback)
    # sys.exit(1)

# Back up the reference to the exceptionhook
sys._excepthook = sys.excepthook

# Set the exception hook to our wrapping function
sys.excepthook = my_exception_hook
######################################################################################

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())        


'''
        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 
        #yesterday = datetime.today() - timedelta(1)
        #print(yesterday.strftime("%Y-%m-%d")

        strPaste = clipboard.paste()
        if strPaste[0:5] == "안녕하세요":
            pyautogui.hotkey('ctrl', 'home')
            pyautogui.hotkey('ctrl', 'enter')
            pyautogui.hotkey('ctrl', 'home')
            strPaste = strTodayOWL + "\n" + strPaste
            clipboard.copy(strPaste)
            pyautogui.hotkey('ctrl', 'v')
        else:
            pyautogui.hotkey('ctrl', 'home')
            pyautogui.hotkey('ctrl', 'enter')
            pyautogui.hotkey('ctrl', 'home')
            clipboard.copy(strTodayOWL)
            pyautogui.hotkey('ctrl', 'v')
'''    