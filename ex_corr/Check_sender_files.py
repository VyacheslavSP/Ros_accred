import time
import win32com.client


def get_maximum_rows(sheet):
    rows = 0
    i = 1
    while i < 10000:
        if (sheet.Cells(i, 1).value != None):
            rows += 1
            i += 1
        else:
            break
    return rows


def check_send_files(path_of_excell):
    check_array = build_arrey_for_check()
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_of_excell)
    sheet = wb.ActiveSheet
    i = 0
    while i in range(len(check_array)):
        # ДОДЕЛАТЬ!!!
        if (str(sheet.Cells(i+2, 1).value).split()[0] in check_array):
            sheet.rows(i+2).Delete()
        else:
            i += 1
    if (sheet.Cells(2, 1).value != None):  # если есть хоть одна строка
        wb.Save()
        time.sleep(0.5)
        wb.Close()
        time.sleep(0.5)
        Excel.Quit()
        return True
    else:
        wb.Save()
        time.sleep(0.5)
        wb.Close()
        time.sleep(0.5)
        Excel.Quit()
        return False


def build_arrey_for_check():
    check_array = []
    f = open('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/log_number_send.txt',
             'r', encoding='utf-8')
    while (True):
        line = f.readline()
        if not line:
            break
        check_array.append(line.strip().split('*')[0])
    f.close
    return check_array


def build_array_check_fros_stas(path_of_excell_Stas):
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_of_excell_Stas)
    sheet = wb.ActiveSheet
    maxRows = get_maximum_rows(sheet)
    i = 1
    check_array_Stas = []
    while i in range(maxRows):
        check_array_Stas.append(sheet.Cells(i+1, 1).value.strip())
        i += 1
    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()
    return check_array_Stas


def check_send_files_for_Stas_Xls(path_of_excell_Stas, path_check_excel):
    check_array_Stas = build_array_check_fros_stas(path_of_excell_Stas)
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_check_excel)
    sheet = wb.ActiveSheet
    maxRows = get_maximum_rows(sheet)
    i = 0

    while i in range(maxRows):
        if (str(sheet.Cells(i+2, 1).value).split()[0] in check_array_Stas):
            sheet.rows(i+2).Delete()
        else:
            i += 1
    if (sheet.Cells(2, 1).value != None):  # если есть хоть одна строка
        wb.Save()
        time.sleep(0.5)
        wb.Close()
        time.sleep(0.5)
        Excel.Quit()
        return True
    else:
        wb.Save()
        time.sleep(0.5)
        wb.Close()
        time.sleep(0.5)
        Excel.Quit()
        return False


# print(check_send_files(
#    'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/22_3_2023_10;54.xlsx', build_arrey_for_check()))
# проверка на пустоту списка. если нечего отправлять то и нехуй отправлять
