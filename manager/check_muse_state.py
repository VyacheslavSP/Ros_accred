import win32gui
import time
import pyautogui


def check_coursor_now(need_wate, wate_second, need_normal):
    need_wate = bool(need_wate)
    count = 0
    need_normal = bool(need_normal)
    start = time.perf_counter()
    while (True):
        tmp_tuple = win32gui.GetCursorInfo()
        tmp = tmp_tuple[1]
        if (not need_normal):
            if (tmp != 65539):
                return
        else:
            if (count % 2 == 0):
                pyautogui.moveTo(602, 519)
            else:
                pyautogui.moveTo(600, 519)
            if (tmp == 65539):
                return
        count += 1
        now = time.perf_counter()
        if (not (need_wate)):
            start = time.perf_counter()
        if (now-start > wate_second):
            print("время истекло")
            return
