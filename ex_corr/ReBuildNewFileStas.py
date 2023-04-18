import win32com.client
from Ex_correct import get_maximum_rows


def ReBuild_New_stas_file(path_stas, path_tmp):
    Excel_stas = win32com.client.Dispatch("Excel.Application")
    wb_stas = Excel_stas.Workbooks.Open(path_stas)
    sheet_stas = wb_stas.ActiveSheet
    max_rows_stas = get_maximum_rows(sheet_stas)+1  # +1 строка
    Excel_tmp = win32com.client.Dispatch("Excel.Application")
    wb_tmp = Excel_tmp.Workbooks.Open(path_tmp)
    sheet_tmp = wb_tmp.ActiveSheet
    max_rows_tmp = get_maximum_rows(sheet_tmp)
    for i in range(max_rows_tmp):
        sheet_stas.Cells(
            max_rows_stas+i, 1).Value = sheet_tmp.Cells(i+2, 1).Value
        sheet_stas.Cells(
            max_rows_stas+i, 2).Value = sheet_tmp.Cells(i+2, 2).Value
        sheet_stas.Cells(
            max_rows_stas+i, 3).Value = sheet_tmp.Cells(i+2, 3).Value
        sheet_stas.Cells(
            max_rows_stas+i, 4).Value = sheet_tmp.Cells(i+2, 4).Value
        sheet_stas.Cells(
            max_rows_stas+i, 5).Value = sheet_tmp.Cells(i+2, 5).Value
        sheet_stas.Cells(
            max_rows_stas+i, 6).Value = sheet_tmp.Cells(i+2, 6).Value
        sheet_stas.Cells(
            max_rows_stas+i, 8).Value = sheet_tmp.Cells(i+2, 7).Value
    wb_tmp.Save()
    wb_tmp.Close()
    wb_stas.Save()
    wb_stas.Close()
    Excel_stas.Quit()
    Excel_tmp.Quit()
