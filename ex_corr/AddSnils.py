import win32com.client
from Bulid_poveritel_class_cortege import Build_poveritel_list
from Ex_correct import get_maximum_rows


def addSnilsAndDeleteNameSoname(path):
    disct_pover = Build_poveritel_list()[1]
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    max_rows = get_maximum_rows(sheet)
    for i in range(max_rows-1):
        tmp = sheet.Cells(i+2, 40).Value.lstrip()
        tmp = tmp.split(' ')[0]
        sheet.Cells(i+2, 41).NumberFormat = "@"
        sheet.Cells(i+2, 41).Value = str(disct_pover.get(tmp))
        sheet.Cells(i+2, 40).Value = tmp
    wb.Save()
    wb.Close()
    Excel.Quit()
