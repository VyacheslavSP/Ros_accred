from Bulid_poveritel_class_cortege import Build_poveritel_list
from Work_with_filters import once_filters, many_filters, mouse_pozition
import sys
import os
import shutil
import subprocess

import time

import keyboard
import pyautogui
import win32com.client
import win32con
import win32gui

import check_muse_state
sys.path.insert(0, "C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/")


def FIND_AND_RESTORE_EXCEL():
    i = 0

    def enum_callback(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
        time.sleep(0.2)
    while (i < 5):
        toplist = []
        winlist = []
        win32gui.EnumWindows(enum_callback, toplist)
        Excel = [(hwnd, title)
                 for hwnd, title in winlist if 'excel' in title.lower()]
        try:
            Excel = Excel[0]
            break
        except:
            time.sleep(0.2)
            i += 1

    win32gui.ShowWindow(Excel[0], win32con.SW_RESTORE)
    time.sleep(0.5)
    keyboard.send("Alt")
    time.sleep(0.5)
    win32gui.SetForegroundWindow(Excel[0])


def save_tmp_csv_xlx_convert(Name):

    tmp_path = 'C:/Unitess/TEMP/tmp'+Name+'.xlsx'
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    time.sleep(0.5)
    keyboard.send("Alt")
    time.sleep(0.2)
    keyboard.send("Ctrl+S")
    time.sleep(0.2)
    keyboard.send("Enter")
    time.sleep(0.2)
    keyboard.send("Ctrl+S")
    time.sleep(0.2)
    keyboard.send("Tab")
    time.sleep(0.2)
    keyboard.send("Enter")
    time.sleep(0.2)
    keyboard.write("tmp"+Name)
    time.sleep(0.5)
    keyboard.send("Enter")
    time.sleep(0.5)
    keyboard.send("Alt+F4")
    return tmp_path


def main_work_manager_new():
    Poveritel_list = Build_poveritel_list()[0]
    flag_correct_pasword = False
    count = 0
    MANAGER = subprocess.Popen('C:/Unitess/Метрология 4.0.exe')
    win = 0
    while win == 0:
        time.sleep(0.5)
        win = win32gui.FindWindow(None, 'Аутентификация')
        time.sleep(0.5)
    time.sleep(0.5)
    win32gui.ShowWindow(win, win32con.SW_RESTORE)
    time.sleep(0.5)
    keyboard.send("Alt")
    time.sleep(0.5)
    try:
        win32gui.SetForegroundWindow(win)
        pyautogui.moveTo(877, 546)
        check_muse_state.check_coursor_now(True, 10, False)
    except:
        print("ошибка фокуса окна")
    finally:
        time.sleep(2)
        keyboard.send("1")
        time.sleep(0.2)
        keyboard.send("2")
        time.sleep(0.2)
        keyboard.send("3")
        time.sleep(0.2)
        keyboard.send("Enter")
        time.sleep(2)

    check_muse_state.check_coursor_now(False, 100, True)
    mouse_pozition(664, 772,  0.5)
    once_filters()
    for poveritel in Poveritel_list:
        many_filters(poveritel.count_pozition_filters)
        pyautogui.moveTo(602, 519, 1)
        check_muse_state.check_coursor_now(False, 100, True)
        time.sleep(2)
        win = 0
        while win == 0:
            try:
                win = win32gui.FindWindow(None, 'Ошибка')
                if (win != 0):
                    break
            except:
                win = 0
        win32gui.ShowWindow(win, win32con.SW_RESTORE)
        time.sleep(2)
        mouse_pozition(969, 705,  1)
        pyautogui.moveTo(602, 519)
        pyautogui.mouseDown(button='right')
        pyautogui.mouseUp(button='right')
        time.sleep(0.5)
        for i in range(5):
            keyboard.send("Down")
        time.sleep(0.1)
        for i in range(2):
            keyboard.send("Enter")
            time.sleep(0.1)
        time.sleep(0.5)
        while (True):
            if (win32gui.FindWindow(None, 'Экспорт в Excell')) != 0:
                time.sleep(0.1)
            if ((win32gui.FindWindow(None, 'Экспорт в Excell')) == 0):
                break
        time.sleep(0.5)
        FIND_AND_RESTORE_EXCEL()
        tmp_path = save_tmp_csv_xlx_convert(poveritel.Short_name)
        time.sleep(0.5)
        shutil.move(tmp_path,
                    'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmpXLS/')
        time.sleep(0.5)
    MANAGER.kill()
    return


# main_work_manager_new()
