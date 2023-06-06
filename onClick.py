import pyautogui
import time


def onClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x/2,location.y/2,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("Picture Not Found. Retry in 0.1 second")
            time.sleep(0.1)

    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x/2,location.y/2,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)

    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x/2,location.y/2,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("Repeat")
                i += 1
            time.sleep(0.1)