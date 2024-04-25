# import keyboard
import os
import pyautogui as pg
from pyautogui import ImageNotFoundException
import time


# Prerequisites:
#  close firefox previous time so window is width you like
# Recommend the user take a screenshot to find upper part (left or right) and use this
#  in `locateOnScreen()` pg method.

class OpenFirefoxWindows:
    def __init__(self):
        self.screen_dim = None
        self.screen_w = None
        self.screen_h = None
        self.firefox_left = None
        self.firefox_top = None
        self.screenshot = None

    def get_monitor_dimensions(self):
        try:
            self.screen_w, self.screen_h = pg.size()
        except (AttributeError, IndexError, Exception) as error:
            print(error)

    @staticmethod
    def firefox_one_open():
        pg.press('win')
        time.sleep(1)
        pg.write('firefox')
        time.sleep(1)
        pg.press('enter')
        time.sleep(3)

    def firefox_one_locate(self):
        try:
            # Box(left=0, top=0, width=319, height=86)
            img_path = os.path.abspath(
                os.curdir) + "/static/firefox_locate.png"
            firefox_locator = pg.locateOnScreen(img_path)
            self.firefox_left = firefox_locator.left + 8
            self.firefox_top = firefox_locator.top + 8
        except ImageNotFoundException as infe:
            print(infe)

    def firefox_one_drag(self):
        pg.moveTo(self.firefox_left, self.firefox_top)
        pg.dragTo(0, 0, button='left')
        pg.mouseInfo()

    @staticmethod
    def firefox_open_two():
        print('hi')
        pg.hotkey('ctrl', 'n')
        print('hi')


if __name__ == '__main__':
    open_firefox_windows = OpenFirefoxWindows()
    open_firefox_windows.get_monitor_dimensions()
    open_firefox_windows.firefox_one_open()
    open_firefox_windows.firefox_one_locate()
    open_firefox_windows.firefox_one_drag()
    open_firefox_windows.firefox_open_two()
    print(open_firefox_windows.firefox_left)
    print(open_firefox_windows.firefox_top)
