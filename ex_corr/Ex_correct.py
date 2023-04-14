
import win32com.client
import shutil
import os
import time
from pathlib import Path
import Check_sender_files


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


def set_text_format(sheet, row):
    i = 1
    j = 1
    while i < 7:
        while j < row+1:
            sheet.Cells(j, i).NumberFormat = "@"
            j += 1
        i += 1
        j = 1
    return


def send_id(array_id):
    f = open('C:\\Users\\VecheslavSP\\Desktop\\Python\\Ros_accred\\ex_corr\\number_send.txt',
             'w+', encoding='utf-8')
    for index in range(len(array_id)):
        if (index+1 != len(array_id)):
            f.write(str(array_id[index]).split()[0]+'\n')
        else:
            f.write(str(array_id[index]).split()[0])
    f.close
    return


def create_id_array(sheet, row):
    array_id = []
    for iter in range(row):
        array_id.append(sheet.Cells(iter+2, 1).value)
    array_id.pop(row-1)
    return array_id


def valid_date(max_row, sheet):
    j = 3
    while j < 5:  # 2 колонки с датами
        for row in range(max_row):
            object = sheet.Cells(row+2, j).value
            if object != None:
                tmp_arr = (object).split('.')
                for elemet in range(len(tmp_arr)):
                    if (len(tmp_arr[elemet]) < 2):
                        if (tmp_arr[elemet] != ' '):
                            tmp_arr[elemet] = '0'+str(tmp_arr[elemet])
                tmp_str = ""
                i = 0
                while i < len(tmp_arr):
                    tmp_str += tmp_arr[i]+"."
                    i += 1
                tmp_str = tmp_str.replace(' ', '')
                tmp_str = tmp_str[:len(tmp_str)-1]

                sheet.Cells(row+2, j).value = tmp_str
        j += 1


def valid_person(max_row, sheet):
    for row in range(max_row):
        tmp_str = ''
        tmp_str = sheet.Cells(row+2, 6).value
        if (tmp_str != None):
            tmp_str = tmp_str[:len(tmp_str)-6]
            tmp_str = tmp_str.replace(' ', '')
            sheet.Cells(row+2, 6).value = tmp_str


def upper_string(max_row, sheet):
    for row in range(max_row-1):
        sheet.Cells(row+2, 5).value = str.title(sheet.Cells(row+2, 5).value)


def rename_sheet(sheet):
    sheet.Name = "Лист1"


def Work_TMP_Excel():
    #
    status_list_check = [" Номер результирующего  документа", " Госреестр", " Дата результирующего  документа",
                         " Дата действия  результирующего документа", " Исполнитель работ"]
    status_list_correct = ["Номер результатов поверки средств измерений", "Тип поверяемого средства измерений",
                           "Дата результатов поверки", "Дата действия результатов поверки", "Фамилия лица,проводившего поверку"]

    #
    # список допущенных
    root_list = ['Пугачев', 'Белов', 'Власов', 'Трошкина', 'Маряхин',
                 'Дзюба', 'Петрунин', 'Блохин', 'Голованова',  'Добровольская', 'Максимов']
    ##
    path = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmp.xlsx'
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    sheet.columns("A:AH").Delete()
    time.sleep(0.5)
    i = 0
    while i < len(status_list_check)+1:
        if (sheet.Cells(1, i+1).value == status_list_check[i]):
            sheet.Cells(1, i+1).value = status_list_correct[i]
            i += 1
        else:
            break

    max_row = get_maximum_rows(sheet)
    delete_rows_without_acsess(sheet, max_row, root_list)  # !

    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()
    time.sleep(0.5)
    Check_sender_files.check_send_files(path)
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    max_row = get_maximum_rows(sheet)  # 1
    send_id(create_id_array(sheet, max_row))
    valid_date(max_row, sheet)
    valid_person(max_row, sheet)
    set_text_format(sheet, max_row)
    upper_string(max_row, sheet)
    time.sleep(0.5)
    rename_sheet(sheet)
    time.sleep(0.5)
    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()
    time.sleep(0.5)
    nameTostr = time.localtime()
    name = str(nameTostr[2])+"_"+str(nameTostr[1])+"_"+str(nameTostr[0]
                                                           )+"_"+str(nameTostr[3])+";"+str(nameTostr[4])+".xlsx"

    os.rename('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmp.xlsx',
              'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/'+name)
    path_of_excell = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/'+name
    return path_of_excell


def delete_rows_without_acsess(sheet, max_row, root_list):
    i = 0
    while (i < max_row):
        if (str(sheet.Cells(i+2, 6).value).split('.')[0].split()[0] in root_list):
            i += 1
        else:
            sheet.rows(i+2).Delete()
        if (sheet.Cells(i+2, 6).value == None):
            break


def write_rows_in_Stas(path_of_excel_Stas, path_of_excell_my, flag):
    if (flag):
        ExcelStas = win32com.client.Dispatch("Excel.Application")
        wbStas = ExcelStas.Workbooks.Open(path_of_excel_Stas)
        sheetStas = wbStas.ActiveSheet
        startRow = get_maximum_rows(sheetStas)
        ExcelMy = win32com.client.Dispatch("Excel.Application")
        wbMy = ExcelMy.Workbooks.Open(path_of_excell_my)
        sheetMy = wbMy.ActiveSheet
        maxMyRow = get_maximum_rows(sheetMy)
        i = 0
        while (i < maxMyRow):
            sheetStas.Cells(startRow+i+1, 1).value = sheetMy.Cells(i+2, 1)
            sheetStas.Cells(startRow+i+1, 2).value = sheetMy.Cells(i+2, 2)
            sheetStas.Cells(startRow+i+1, 3).value = sheetMy.Cells(i+2, 3)
            sheetStas.Cells(startRow+i+1, 4).value = sheetMy.Cells(i+2, 4)
            sheetStas.Cells(startRow+i+1, 5).value = sheetMy.Cells(i+2, 5)
            sheetStas.Cells(startRow+i+1, 6).value = sheetMy.Cells(i+2, 6)
            if sheetMy.Cells(i+2, 6).value == "Власов":
                sheetStas.Cells(startRow+i+1, 8).value = 2
            else:
                sheetStas.Cells(startRow+i+1, 8).value = 1
            i += 1
        wbMy.Save()
        wbStas.Save()
        time.sleep(0.5)
        wbMy.Close()
        wbStas.Close()
        time.sleep(0.5)
        ExcelStas.Quit()
        ExcelMy.Quit()
    else:
        print("ничего не надо")
