import time
import win32com.client


def build_correct_excel(path_of_excell):
    d_array = build_2d_array()
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_of_excell)
    sheet = wb.ActiveSheet
    for i in range(len(d_array)):
        sheet.Cells(i+2, 1).Value = d_array[i][0]
        sheet.Cells(i+2, 2).Value = d_array[i][1]
        sheet.Cells(i+2, 3).Value = d_array[i][2]
        try:
            sheet.Cells(i+2, 4).Value = d_array[i][3]
        except:
            sheet.Cells(i+2, 4).Value = ''
        i += 1
    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()
    time.sleep(0.5)
    return


def build_2d_array():
    d_array = []
    f = open('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/valid_value_send.txt',
             'r', encoding='utf-8')
    while (True):
        line = f.readline()
        if not line:
            break
        d_array.append(line.strip().split('*'))
    f.close
    return d_array
