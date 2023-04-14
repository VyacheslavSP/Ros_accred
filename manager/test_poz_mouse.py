import os
import win32con
import win32gui
import time
import keyboard
while (True):
    tmp_tuple = win32gui.GetCursorInfo()
    print(tmp_tuple)
   # break


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


def save_tmp_csv_xlx_convert():
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


FIND_AND_RESTORE_EXCEL()
save_tmp_csv_xlx_convert()
