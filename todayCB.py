from datetime import datetime
import clipboard
import time
import pyautogui
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
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
        btn_1 = QPushButton('Only Today Write', self)
        btn_1.move(10, 10)
        btn_1.resize(btn_1.sizeHint())
        #btn_1.clicked.connect(QCoreApplication.instance().quit)
        btn_1.clicked.connect(self.button1Function)

        btn_2 = QPushButton('Today+Clipboard Write', self)
        btn_2.move(10, 40)
        btn_2.resize(btn_2.sizeHint())
        btn_2.clicked.connect(self.button2Function)

        self.setWindowTitle('OWL Semi-Auto Write')
        self.setGeometry(300, 300, 300, 200)
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
        strToday = datetime.today().strftime("%Y-%m-%d")  # YYYY/mm/dd HH:MM:SS 형태의 시간 출
        strTodayOWL = "[" + strToday + " 고재정]" 

        time.sleep(1)
        strPaste = clipboard.paste()
        if strPaste[0:5] == "안녕하세요":
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