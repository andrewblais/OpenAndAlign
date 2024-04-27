import pyautogui as pg
import sys
import time
from win32gui import GetForegroundWindow, GetWindowRect  # noqa


class OpenAndAlign:
    """
    A class to automate the process of opening an application and aligning its window.

    Arguments:
    :param app_name: The name of the application to open via the Start Menu.
    :type app_name: str
    :param night_light_window: Whether the opened app is a Windows Night Light.
                               Defaults to False.
    :type night_light_window: bool
    :param settings_val_01: Number or string to enter into opened Settings dialog.
    :type settings_val_01: int | str
    :param app_relocate: Whether or not to move the app window from initial position.
                         Defaults to False.
    :type app_relocate: bool
    :param image_file_name: The file name of the image to be located on the screen.
                            User creates and stores this image file in the 'static'
                            project folder.
    :type image_file_name: str
    :param desired_win_loc: The desired window location coordinates. Defaults to (0, 0).
    :type desired_win_loc: list, optional
    :param align_right: Whether to align the window to the right side of the screen.
                        Defaults to False.
    :type align_right: bool, optional
    :param confidence: The confidence threshold for image recognition. Defaults to 0.9.
    :type confidence: float, optional

    Attributes:
    :ivar app_name: The name of the application to open via the Start Menu.
    :type app_name: str
    :ivar night_light_window: Whether the opened app is a Windows Night Light.
                              Defaults to False.
    :type night_light_window: bool
    :ivar settings_val_01: Number or string to enter into opened Settings dialog.
    :type settings_val_01: int | str
    :ivar app_relocate: Whether or not to move the app window from initial position.
                        Defaults to False.
    :type app_relocate: bool
    :ivar image_file_name: The file name of the image to be located on the screen.
                           User creates and stores this image file in the 'static'
                           project folder.
    :type image_file_name: str
    :ivar desired_win_loc: The desired window location coordinates. Defaults to (0, 0).
    :type desired_win_loc: list
    :ivar align_right: Whether to align the window to the right side of the screen.
                       Defaults to False.
    :type align_right: bool
    :ivar screen_w: The width of the screen.
    :type screen_w: int
    :ivar app_left: The left coordinate of the located application window.
    :type app_left: int
    :ivar app_top: The top coordinate of the located application window.
    :type app_top: int
    :ivar corner_avoid_val: The value to avoid the corner of the screen.
    :type corner_avoid_val: int
    :ivar right_adjust_val: The value to adjust for aligning the window to the right side.
    :type right_adjust_val: int
    :ivar confidence: The confidence threshold for image recognition. Defaults to 0.9.
    :type confidence: float
    """

    def __init__(self,
                 app_name: str,
                 night_light_window: bool = False,
                 settings_val_01: int = 10,
                 app_relocate: bool = False,
                 image_file_name: str = "",
                 desired_win_loc: list = [0, 0],  # noqa
                 align_right: bool = False,
                 confidence: float = .9):
        self.app_name = app_name
        self.night_light_window = night_light_window
        self.settings_val_01 = settings_val_01
        self.app_relocate = app_relocate
        self.image_file_name = image_file_name
        self.desired_win_loc = desired_win_loc
        self.align_right = align_right
        self.screen_w = 1920
        self.app_left = 0
        self.app_top = 0
        self.corner_avoid_val = 10
        self.right_adjust_val = 15
        self.confidence = confidence

        self.run_program()

    @staticmethod
    def error_handling(exception: Exception) -> None:
        """
        Static method to handle errors that might arise.

        :return: None
        """
        print(f"Unexpected error: {exception}. Discretely exiting program...")
        sys.exit(1)

    def screen_w_val(self) -> None:
        """
        Gets the width of the screen.

        :return: None
        """
        try:
            self.screen_w = pg.size()[0]
        except (Exception,) as e:
            self.error_handling(e)

    def app_open(self, app_name: str) -> None:
        """
        Opens the specified application via the Start Menu.

        :param app_name: The name of the application.
        :type app_name: str

        :return: None
        """
        try:
            pg.press('win')
            time.sleep(1)
            pg.write(app_name)
            time.sleep(1)
            pg.press('enter')
            time.sleep(3)
        except (Exception,) as e:
            self.error_handling(e)

    def app_locate(self, image_file_name: str) -> None:
        """
        Locates the specified image on the screen. User must crop a screenshot an image
         representing the upper-left part of the app to identify the app to pyautogui.
        Store this image in the project 'static' folder.

        :param image_file_name: The file name of the image.
        :type image_file_name: str

        :return: None
        """
        try:
            img_path = f"./static/{image_file_name}"
            app_locator = pg.locateOnScreen(img_path, confidence=self.confidence)
            self.app_left = app_locator.left + 8 + self.corner_avoid_val
            self.app_top = app_locator.top + 8 + self.corner_avoid_val
        except (Exception,) as e:
            self.error_handling(e)

    def night_light_change(self) -> None:
        """
        Tabs, slides, etc... within a Settings dialog and adjusts its settings.

        :return: None
        """
        if self.night_light_window:
            time.sleep(2)
            pg.press('tab')  # Tabs to Night Light value slider
            time.sleep(1)
            pg.press('left', presses=100)  # Resets slider to 0
            time.sleep(1)
            pg.press('right', presses=self.settings_val_01)  # Moves slider to desired val
            time.sleep(1)
            pg.hotkey('alt', 'f4')  # Closes Settings window

    def right_move(self) -> int:
        """
        Adjusts values based on window width so window will be snug against right
         side of computer screen.

        :return: The adjusted x-coordinate.
        :rtype: int | float
        """
        try:
            window = GetForegroundWindow()
            window_rect = GetWindowRect(window)
            app_width = window_rect[2] - window_rect[0]
            return self.screen_w - app_width + self.right_adjust_val
        except (Exception,) as e:
            self.error_handling(e)

    def app_drag(self, desired_win_locations: list) -> None:
        """
        Moves mouse cursor to upper left of app window and drags to desired location.

        :param desired_win_locations: The desired window location coordinates.

        :return: None
        """
        try:
            pg.moveTo(self.app_left, self.app_top)
            x, y = desired_win_locations
            # Increment pixel distance to turn mouse cursor from 'resizer' to 'grabber':
            x += 8
            y += 10
            pg.dragTo(x, y, button='left', duration=1)
        except (Exception,) as e:
            self.error_handling(e)

    def run_program(self) -> None:
        """
        Runs the program and all relevant class methods to open the application and
         align its window.

        :return: None
        """
        try:
            self.screen_w_val()
            self.app_open(self.app_name)
            if self.night_light_window:
                self.night_light_change()
            if self.app_relocate:
                self.app_locate(self.image_file_name)
                if self.align_right:
                    self.desired_win_loc[0] += self.right_move()
                self.app_drag(self.desired_win_loc)
        except (Exception,) as e:
            self.error_handling(e)


# Instantiate and run for dual Firefox windows in the developers favorite config:
if __name__ == '__main__':
    try:
        firefox_01 = OpenAndAlign(app_name='firefox',
                                  app_relocate=True,
                                  image_file_name='firefox_locate.png')

        firefox_02 = OpenAndAlign(app_name='firefox',
                                  app_relocate=True,
                                  image_file_name='firefox_locate.png',
                                  align_right=True)

        # Include after packaging as .exe:
        # py_charm = OpenAndAlign(app_name='pycharm',
        #                         app_relocate=False)

        vs_code = OpenAndAlign(app_name='vscode',
                               app_relocate=False)

        night_light = OpenAndAlign(app_name='night light',
                                   night_light_window=True,
                                   settings_val_01=10,
                                   app_relocate=False)

    except (Exception,) as e_:
        print(f"Unexpected error instantiating OpenAndAlign class: {e_}. Exiting...")
        sys.exit(1)
