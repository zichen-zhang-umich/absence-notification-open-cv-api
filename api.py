import pyautogui
import time
import xlrd
import pyperclip

def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("Picture Not Found. Retry in 0.1 second")
            time.sleep(0.1)

    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)

    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("Repeat")
                i += 1
            time.sleep(0.1)

# How to understand cmdType.value: 1 stands for left click once, 2 stands for left click twice, 3 stands for right click once
# 4 stands for type, 5 stands for wait, 6 stands for scroll

# How to understand ctype: 0 stands for empty, 1 stands for string, 2 stands for number, 3 stands for date,
# 4 stands for boolean, 5 stands for error

def dataCheck(sheet1):
    checkCmd = True
    if sheet1.nrows<2:
        print("Warning: No Data Present")
        checkCmd = False
    i = 1
    while i < sheet1.nrows:
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('Error on row #',i+1," col #1")
            checkCmd = False
        cmdValue = sheet1.row(i)[1]
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        i += 1
    return checkCmd


def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("Left click",img)
        elif cmdType.value == 2.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("Left double click",img)
        elif cmdType.value == 3.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("Right click",img) 
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("Type:",inputValue)                                        
        elif cmdType.value == 5.0:
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("Wait",waitTime," second(s)")
        elif cmdType.value == 6.0:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("Scroll",int(scroll)," distance")                      
        i += 1

if __name__ == '__main__':
    file = 'cmd.xls'
    wb = xlrd.open_workbook(filename=file)
    sheet1 = wb.sheet_by_index(0)
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key=input('Enter 1: Run one time; Enter 2: Continue running\n')
        if key=='1':
            mainWork(sheet1)
        elif key=='2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("Wait for 0.1 second")    
    else:
        print('Program ends')
