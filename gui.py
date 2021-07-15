from datetime import datetime
from pyautogui import press, hotkey, moveTo, click, rightClick, write
from time import sleep
from os import startfile

print(datetime.now())

def open_pims():
  startfile(r"C:\Users\Elvin.Bashirli\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Omega AS\Pims BP CMS.appref-ms")
  sleep(20)

def kt(k, t=1):
  for i in range(t):
    press(k)
    sleep(.2)

def login_pims():
  kt('tab', 6)
  press('enter')
  sleep(20)

def open_tab(tab_name):
  hotkey('alt', 'o')
  sleep(3)
  write(tab_name)
  sleep(2)
  press('tab')
  press('enter')
  sleep(15)

def save_excel(file_name):
  sleep(5)
  moveTo((702, 61))
  click()
  sleep(2)
  moveTo((702, 250))
  rightClick()
  kt('down', 10)
  press('right')
  kt('down', 4)
  press('enter')
  write(file_name)
  press('enter')
  hotkey('alt', 'y')
  sleep(2)
  press('tab')
  press('enter')

open_pims()
login_pims()

tab_file_names = ['systems (cms)', 'subsystems', 'packages']

for tab_file_name in tab_file_names: 
  open_tab(tab_file_name)
  save_excel(tab_file_name)
  sleep(5)

startfile(r"C:\Users\Elvin.Bashirli\Desktop\MC\New folder\updater.py")