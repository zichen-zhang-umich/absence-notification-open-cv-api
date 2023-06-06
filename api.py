# a program for messenger automation


import pyautogui
import time
import xlrd
import pyperclip


from onClick import onClick
from check import check


def mainWork(img):
    i = 1 # counter

    while i < sheet1.nrows:
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            onClick(1,"left",img,reTry)
            print("Left click",img)

        elif cmdType.value == 2.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            onClick(2,"left",img,reTry)
            print("Left double click",img)

        elif cmdType.value == 3.0:
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            onClick(1,"right",img,reTry)
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
    checkCmd = check(sheet1)
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