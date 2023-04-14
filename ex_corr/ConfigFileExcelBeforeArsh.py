import win32com.client
from Ex_correct import get_maximum_rows, set_text_format, delete_rows_without_acsess
from Bulid_poveritel_class_cortege import Build_poveritel_list


def configFileBeforeArsh(path):
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    max_rows = get_maximum_rows(sheet)
    sheet.columns("A:AH").Delete()
    set_text_format(sheet, max_rows)
    dictPover = Build_poveritel_list()[1]
    root_list = dictPover.keys()
    delete_rows_without_acsess(sheet, max_rows, root_list)
    wb.Save()
    wb.Close()
    Excel.Quit()


configFileBeforeArsh(
    'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmp.xlsx')
