import pyautogui
import pyperclip
import pandas
import numpy
import openpyxl
import time

pyautogui.PAUSE = 1


# Step 1: Enter the company's system

pyautogui.hotkey('ctrl', 'alt', '0') # Open the browser
pyperclip.copy('https://drive.google.com/drive/u/5/folders/1-IO1PVmfVQkPxahLVI7vpyFeElwIgm1K')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# Step 2: Navigate to the report location

time.sleep(3) # Waiting time in seconds
pyautogui.click(x=539, y=480)

# Step 3: Export the report

time.sleep(3)
pyautogui.click(x=1609, y=236)
pyautogui.click(x=1408, y=797)

#Step 4: Calculate the indicators (income and quantity of products)

time.sleep(3)
table = pandas.read_excel(r'C:\Users\andre\Downloads\Vendas - Dez.xlsx')

billing = table['Valor Final'].sum()
amount = table['Quantidade'].sum()

# Step 5: Send an e-mail to board of directors

time.sleep(3)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/?ogbl#inbox') # Enter my email
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=123, y=255) # Click to write email
pyautogui.write('andreluis0703@gmail.com')
pyautogui.press('enter')
pyautogui.press('tab')
pyperclip.copy('Sales Report')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

text = f"""
Dears, good morning!

Yesterday's billing was: R${billing:,.2f}
The number of products was: {amount:,}

SY, André Ramos
"""

pyperclip.copy(text)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)
pyautogui.hotkey('ctrl', 'enter')
