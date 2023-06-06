# Messenger Absence Notification API

An idea originated from [RPA](https://www.automationanywhere.com/rpa/robotic-process-automation#:~:text=What%20is%20Robotic%20Process%20Automation,execute%20rules%2Dbased%20business%20processes.) Robotic Process Automation. This program is designed to mimic this concept.

This is a Python project designed for **MacOS PC** that helps people by controlling their computer mice to automate the process of sending absence notifications to those who sent their messages.

The default `sample1.png` will be auto-clicked when the program is executed on a MacOS PC. You can change the command by modifing `cmd.xls`.

## Install Modules

To run the program on your device, you should download the necessary modules by typing the following on your computer terminal:

```
pip install pyperclip
pip install xlrd
pip install pyautogui==0.9.50
pip install opencv-python
pip install pillow
```

## Windows Retina Adjustion
It's important that we adjust the program to fit into different PCs. To operate on a Windows computer, you should change the following function in `check.py` from
```Python
pyautogui.click(location.x/2,location.y/2,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
```
to:
```Python
pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
```

Additionally, macOS and Windows PCs use different keyboards. So, its necessary to change this function in `api.py`
```Python
pyautogui.hotkey('command','v')
```
to:
```Python
pyautogui.hotkey('cmd','v')
```