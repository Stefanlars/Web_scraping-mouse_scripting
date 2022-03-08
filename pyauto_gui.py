import pyautogui
import re
import os
import time
size = pyautogui.size()
from datetime import date
print(size)

#TODO: use os.startfile() to start the pdf script. This function opens the file in the default pdf viewer.
#TODO: after that, use pyautogui to edit the text in the document. Once that's finished, Move to a specific folder
#TODO: Once it's moved to a folder the emailing can be done manually

today = date.today()

d1 = today.strftime("%m/%d/%Y")

path = "c:\\hubbel_docs"

files = os.listdir(path)
for file in files:
    os.startfile(f"{path}\\{file}")
    pyautogui.moveTo(962,521, duration=0.25)
    pyautogui.click()
    pyautogui.click(339, 95)
    pyautogui.moveTo(962, 521, duration=0.25)
    time.sleep(2)
    pyautogui.scroll(10000)
    pyautogui.scroll(-6250)

    time.sleep(.2)
    pyautogui.click(182, 53)
    time.sleep(.25)
    pyautogui.click(685, 109)
    time.sleep(.25)
    pyautogui.click(1809, 270)
    time.sleep(.25)
    pyautogui.click(1562, 342)
    time.sleep(.25)
    pyautogui.click(1630, 607)
    time.sleep(.25)
    pyautogui.click(905, 528)
    time.sleep(.25)
    pyautogui.click(1138, 322)

    time.sleep(.25)
    pyautogui.moveTo(800, 438)
    pyautogui.dragTo(800, 470, .25, button='left')
    pyautogui.moveTo(808, 438)
    pyautogui.dragTo(808, 470, .25, button='left')
    pyautogui.moveTo(816, 438)
    pyautogui.dragTo(816, 470, .25, button='left')
    pyautogui.moveTo(824, 438)
    pyautogui.dragTo(824, 470, .25, button='left')
    pyautogui.moveTo(832, 438)
    pyautogui.dragTo(832, 470, .25, button='left')
    pyautogui.moveTo(840, 438)
    pyautogui.dragTo(840, 470, .25, button='left')
    pyautogui.moveTo(848, 438)
    pyautogui.dragTo(848, 470, .25, button='left')
    pyautogui.moveTo(856, 438)
    pyautogui.dragTo(856, 470, .25, button='left')
    pyautogui.moveTo(864, 438)
    pyautogui.dragTo(864, 470, .25, button='left')
    pyautogui.moveTo(872, 438)
    pyautogui.dragTo(872, 470, .25, button='left')
    pyautogui.moveTo(880, 438)
    pyautogui.dragTo(880, 470, .25, button='left')
    pyautogui.moveTo(888, 438)
    pyautogui.dragTo(888, 470, .25, button='left')
    pyautogui.moveTo(896, 438)
    pyautogui.dragTo(896, 470, .25, button='left')
    pyautogui.moveTo(904, 438)
    pyautogui.dragTo(904, 470, .25, button='left')
    pyautogui.moveTo(912, 438)
    pyautogui.dragTo(912, 470, .25, button='left')

    pyautogui.click(425, 108)
    pyautogui.click(800, 439)
    pyautogui.write(f'James Moore \njames.moore@wachter.com\n(479)361-7913\n{d1}')

    pyautogui.click(95, 51)
    pyautogui.click(1048, 116)
    pyautogui.click(544, 110)
    # pyautogui.click(544, 110)
    pyautogui.click(855, 418)

    completed_folder = 'C:\hubbel_docs_finished\n'

    pyautogui.click(28, 49)
    pyautogui.click(78, 236)
    pyautogui.click(412, 193)
    pyautogui.click(745, 690)
    pyautogui.click(925, 193)
    pyautogui.write(f"{completed_folder}")
    pyautogui.click(906, 670)
    pyautogui.write(f"{file}")
    pyautogui.click(1429, 810)

    pyautogui.click(1024, 589)
    time.sleep(2)
    pyautogui.click(1024, 589)
    time.sleep(2)
    pyautogui.click(1896, 15)
