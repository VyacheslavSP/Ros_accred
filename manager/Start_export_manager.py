

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


def mouse_pozition(x, y, sleep):
    pyautogui.click(x, y)
    time.sleep(float(sleep))
    return


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

    tmp_path = 'C:/Unitess/TEMP/tmp.xlsx'
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
    keyboard.write("tmp")
    time.sleep(0.5)
    keyboard.send("Enter")
    time.sleep(0.5)
    keyboard.send("Alt+F4")


def main_work_manager():
    flag_correct_pasword = False
    count = 0
    MANAGER = subprocess.Popen('C:\\Unitess\\Метрология 4.0.exe')
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
        time.sleep(0.2)
        keyboard.send("1")
        time.sleep(0.2)
        keyboard.send("2")
        time.sleep(0.2)
        keyboard.send("3")
        time.sleep(0.2)
        keyboard.send("Enter")
        time.sleep(2)

    check_muse_state.check_coursor_now(False, 100, True)
    mouse_pozition(661, 775, 2)  # фильтр
    mouse_pozition(921, 373, 0.5)  # фильтр сектор
    mouse_pozition(833, 415, 0.5)  # из списка сектор-2
    mouse_pozition(789, 654,  0.5)  # по дате регистрации
    mouse_pozition(845, 735,  0.5)  # из списка-по дате завершения
    mouse_pozition(883, 676,  0.5)  # список за сегоднф
    # mouse_pozition(814, 789, 1)# из списка-за месяц
    # mouse_pozition(832, 763,  0.5)  # из списка-за неделю

    # из в диапазоне. в текущей версии выбирает с начала месяца до  текущей даты
    mouse_pozition(860, 728,  0.5)
    mouse_pozition(1078, 728,  0.5)  # список фильтров
    # mouse_pozition(906, 766, 1) # из списка- фильтр Тест (только пугачев)
    mouse_pozition(823, 791,  0.5)  # из списка- фильтр Zсектор
    mouse_pozition(797, 819,  0.5)
    mouse_pozition(602, 515,  0.5)
    pyautogui.moveTo(602, 519, 1)
    check_muse_state.check_coursor_now(False, 100, True)
    time.sleep(0.5)
    win = win32gui.FindWindow(None, 'Ошибка')
    win32gui.ShowWindow(win, win32con.SW_RESTORE)
    time.sleep(0.5)
    mouse_pozition(969, 705,  0.5)
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
    save_tmp_csv_xlx_convert('C:/Unitess/TEMP/tmp.xlsx')
    time.sleep(0.5)
    shutil.move('C:/Unitess/TEMP/tmp.xlsx',
                'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/')
    time.sleep(0.5)
    MANAGER.kill()
    return
# main_work_manager()
