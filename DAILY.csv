from pywinauto import application
import pyautogui
from datetime import date
import time
import os

try:
    os.remove('DAILY.csv')
except:
    pass

app = application.Application().start("ProControl.exe")
window = app.window(title="Pro Control")

window_rect = window.rectangle()
window_x, window_y = window_rect.left, window_rect.top


absolute_x = window_x + 30
absolute_y = window_y + 40

pyautogui.click(absolute_x, absolute_y)
pyautogui.click(absolute_x, absolute_y)
absolute_y = window_y + 50
pyautogui.click(absolute_x, absolute_y)
absolute_x = window_x + 1100
absolute_y = window_y + 360
pyautogui.click(absolute_x, absolute_y)
today = date.today()

# dd/mm/YY
d = today.strftime("%d")
m = today.strftime("%m")
# y = today.strftime("%Y")
y = today.strftime("%m")
new = app.window(title="File Export/Import")
new.type_keys(m)
new.type_keys("{RIGHT}")
new.type_keys(d)
# new.type_keys("{DOWN}")
new.type_keys("{RIGHT}")
new.type_keys(y)
new.type_keys("{TAB}")
new.type_keys(m)
new.type_keys("{RIGHT}")
new.type_keys(d)
new.type_keys("{RIGHT}")
new.type_keys(y)


absolute_x = window_x + 1110
absolute_y = window_y + 720
pyautogui.click(absolute_x, absolute_y)

new = app.window(title="Open")
new.type_keys('DAILY.DAT')
new.type_keys("{ENTER}")

time.sleep(12)
new = app.window(title="File Export/Import")
new.type_keys("{TAB}")
new.type_keys("{TAB}")
new.type_keys("{TAB}")
new.type_keys("{ENTER}")
new = app.window(title="Save As")
new.type_keys("{ENTER}")
time.sleep(12)
new = app.window(title="File Export/Import")
new.type_keys("{ESC}")
window.close()
