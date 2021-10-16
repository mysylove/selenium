from datetime import datetime
import clipboard
import time
import pyautogui
import sys
import openpyxl as xl
from PyQt5.QtWidgets import * #QApplication, QWidget, QPushButton
from PyQt5.QtCore import QCoreApplication

'''
# 유저가 현재 보고 있는 윈도우차 알아내기
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

        self.setWindowTitle('OWL Semi-Auto Write')
        self.setGeometry(300, 300, 900, 200)
        self.show()
    
    def button1Function(self):
        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 

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
        strFind = '직접인건비'
        if xl.find(strFind) == -1:
            self.label_3.setText("클립보드에 견적내용이 없습니다.")
            return
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
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'home')
        pyautogui.hotkey('ctrl', 'enter')
        pyautogui.hotkey('ctrl', 'home')
        clipboard.copy(strEst)
        pyautogui.hotkey('ctrl', 'v')            

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