import win32com.client
from Ex_correct import get_maximum_rows


def Build_date_forever(path):
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    max_rows = get_maximum_rows(sheet)
    for i in range(max_rows):
        if (str(sheet.Cells(i+2, 5).Value).split()[0] == "Пригодно" and sheet.Cells(i+2, 4).Value == None):
            sheet.Cells(i+2, 4).Value = "31.12.9999"
    wb.Save()
    wb.Close()
    Excel.Quit()

