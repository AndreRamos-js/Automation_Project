import pyautogui
import pyperclip
import pandas
import numpy
import openpyxl
import time

pyautogui.PAUSE = 1


# Step 1: Enter the company's system

pyautogui.hotkey('ctrl', 'alt', '0')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# Step 2: Navigate to the report location

time.sleep(3)
pyautogui.click(x=519, y=416, clicks=2)

# Step 3: Export the report
