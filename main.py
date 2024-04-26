import pyautogui as pg
import sys
import time
from win32gui import GetForegroundWindow, GetWindowRect  # noqa


class OpenAndAlign:
    """
    A class to automate the process of opening an application and aligning its window.

    :param app_name: The name of the application to open via the Start Menu.
    :type app_name: str
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

    :ivar app_name: The name of the application to open via the Start Menu.
    :type app_name: str
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
                 image_file_name: str,
                 desired_win_loc: list = [0, 0],  # noqa
                 align_right: bool = False,
                 confidence: float = .9):
        self.app_name = app_name
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
            self.app_locate(self.image_file_name)
            if self.align_right:
                self.desired_win_loc[0] += self.right_move()
            self.app_drag(self.desired_win_loc)
        except (Exception,) as e:
            self.error_handling(e)


# Instantiate and run for dual Firefox windows in the developers favorite config:
if __name__ == '__main__':
    try:
        open_and_align_01 = OpenAndAlign(app_name='firefox',
                                         image_file_name='firefox_locate.png')

        open_and_align_02 = OpenAndAlign(app_name='firefox',
                                         image_file_name='firefox_locate.png',
                                         align_right=True)
    except (Exception,) as e_:
        print(f"Unexpected error instantiating OpenAndAlign class: {e_}. Exiting...")
        sys.exit(1)
