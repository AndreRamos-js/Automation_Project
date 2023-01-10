import pyautogui
import pyperclip
import pandas
import numpy
import openpyxl
import time

pyautogui.PAUSE = 2


# Step 1: Enter the company's system

pyautogui.hotkey('ctrl', 'alt', '0')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

# Step 2: Navigate to the report location

#teste