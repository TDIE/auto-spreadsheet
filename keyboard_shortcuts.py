'''
Author: Tom Diederen
Summary: pyAutoGui can be used to automate keystrokes and mouse clicks. 
    This module combines commonly used shortcuts, making them available for re-use in other modules.
    It also adds delays to prevent the keystrokes or clicks from happening to fast for the system to keep up.
'''

import pyautogui
import time
import datetime

pause = 0.5  # Pause before and after keystrokes, in seconds. Needed so the system can keep up with fast keyboard and mouse input.
today = datetime.date.today()
year, week, day_of_week = today.isocalendar()

def add_pause_decorator(func):
    def wrapper():
        time.sleep(pause)
        func()
        time.sleep(pause)
    return wrapper

@add_pause_decorator
def alt_tab():
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

@add_pause_decorator
def alt_num(num):
    pyautogui.keyDown("alt")
    pyautogui.press(str(num))
    pyautogui.keyUp("alt")

@add_pause_decorator
def go_to_line(line):  #ctrl+g
    pyautogui.keyDown("ctrl")
    pyautogui.press("g")
    pyautogui.keyUp("ctrl")
    time.sleep(pause)    
    pyautogui.write(f"{line}")
    time.sleep(pause)  
    pyautogui.press("enter")

@add_pause_decorator
def ctrl_f(query):
    pyautogui.keyDown("ctrl")
    pyautogui.press("f")
    pyautogui.keyUp("ctrl")
    time.sleep(pause)    
    pyautogui.write(f"{query}")
    time.sleep(pause)  
    pyautogui.press("enter")

@add_pause_decorator
def alt_shift_plus():
    pyautogui.hotkey("alt", "shift", "+")

@add_pause_decorator
def alt_shift_minus():
    pyautogui.hotkey("alt", "shift", "-")

@add_pause_decorator
def copy():  # Ctrl+c)
    pyautogui.keyDown("ctrl")
    pyautogui.press("c")
    pyautogui.keyUp("ctrl")

def paste():  # Ctrl+v
    pyautogui.keyDown("ctrl")
    pyautogui.press("v")
    pyautogui.keyUp("ctrl")

def mouse_click_left(x, y):  # x: pixels from left side of the screen, y: pixels from top of the screen.
    time.sleep(pause)
    pyautogui.click(x=x, y=y)
    time.sleep(pause)

def take_screenshot(file_name, ROI):  
    top_left_x, top_left_y, width, height = ROI
    im = pyautogui.screenshot(f"Screenshot_{file_name}_{datetime.datetime.now().strftime('%Y-%m-%d')}.png", region=(top_left_x, top_left_y, width, height))

'''
Example
    1. Update a few values in Excel
    2. Click button to run macro
    3. Take screenshot of updated charts
'''
new_data = [10202, 39240, 9823, -2838, 29823, 8383, 23939, 3839, -198328, 2938, -9283, 3288]  # Sample data. Could be read from a file, URL, etc.

# Switch from VS Code to Excel
alt_tab()

# 1. Update values
pyautogui.click(x=192, y=230)  # Click first cell to past to (sample coordinates).  
time.sleep(0.25)

for num in new_data:
    paste(num)
    time.sleep(0.25)
    pyautogui.press("down")
    time.sleep(0.25)

# 2. Run Macro
pyautogui.click(x=949, y=1230)  #sample coordinates for macro button.  
time.sleep(0.25)

# 3. Take Screenshot of new charts
pyautogui.scroll(-10) #  Scroll down to charts
take_screenshot(f"Updated_Charts_wk{week}.{day_of_week}_{year}", (10, 50, 500, 1000))  # Take screenshot of 500 (width) * 1000 (height) pixels. Top left of the image is 10 pixels to the right and 50 pixels down from the top left corner of the screen.


