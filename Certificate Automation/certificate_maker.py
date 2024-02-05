from config import logger
from excel_data_extractor import yield_excel_row_data

import json
import time
import keyboard
import pyautogui
from typing import Callable

# Loading coordinates of things of canva
with open("canva_coordinates.json") as file:
    file_data = json.loads(file.read())
    # Canva Stuff
    filename_position = file_data["filename_position"]
    share_btn_position = file_data["share_btn_position"]
    download_menu_position = file_data["download_menu_position"]
    empty_area_position = file_data["empty_area_position"]
    # Certificate Stuff
    download_btn_position = file_data["download_btn_position"]
    student_name_position = file_data["student_name_position"]

download_btn_image: str = "Images/download_btn_ready.png"
share_text_image: str = "Images/completed.png"


# Utility Functions
def get_position_on_spacebar():
    while True:
        keyboard.wait("space")
        break
    return pyautogui.position()


def wait_till_image_on_screen(image_path: str):
    while True:
        if pyautogui.locateOnScreen(image_path, confidence=0.7) is not None:
            break


def click_empty_area_on_end(function: Callable):
    def wrapper(*args, **kwars):
        function(*args, **kwars)
        pyautogui.click(empty_area_position)

    return wrapper


# Main Functions
@click_empty_area_on_end
def change_student_name_in_certificate(name: str):
    pyautogui.click(*student_name_position, 2, 0.5)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(name)


@click_empty_area_on_end
def change_canva_filename(student_name: str, college_name: str):
    first_word = college_name.split()[0]
    pyautogui.click(filename_position)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(f"{student_name}-{first_word}")
    pyautogui.press('enter')


@click_empty_area_on_end
def change_college_name(previous_college_name: str, new_college_name: str):
    pyautogui.hotkey('win', '3')  # Opening pycharm to show console message
    logger.info("Move the cursor to the start first letter and press spacebar")
    college_name_position = get_position_on_spacebar()
    # Selecting the second letter of college name
    pyautogui.click(*college_name_position, 2, 0.5)
    pyautogui.press('right')
    # Erasing the college name except the first letter
    for _ in range(len(previous_college_name) - 1):
        pyautogui.press('delete')

    # Writing new college name
    pyautogui.write(new_college_name)

    # Removing te leftover letter
    for _ in range(len(new_college_name) + 2):
        pyautogui.press('left')
    pyautogui.press('delete')


@click_empty_area_on_end
def download_certificate():
    pyautogui.click(share_btn_position)
    time.sleep(0.5)
    pyautogui.click(download_menu_position)
    wait_till_image_on_screen(download_btn_image)
    time.sleep(0.2)
    pyautogui.click(download_btn_position)


def make_certificates():
    previous_college_name = next(yield_excel_row_data())[2]  # First college name in file
    pyautogui.hotkey('win', '2')  # Open chrome from taskbar

    for sno, name, college, duration in yield_excel_row_data():
        # Update if college name needs to updated in certificate
        if previous_college_name != college:
            change_college_name(previous_college_name, college)
            previous_college_name = college

        pyautogui.click(empty_area_position)
        pyautogui.click(empty_area_position)
        change_canva_filename(name, college)
        change_student_name_in_certificate(name)
        download_certificate()

        logger.info(f"[CERTIFICATE] {sno} {name}-{college.split()[0]}")
        # Sleep to avoid counting menu that opened just after the download started
        time.sleep(1)
        wait_till_image_on_screen(share_text_image)


if __name__ == "__main__":
    make_certificates()
