from pathlib import Path
import pandas as pd
import win32com.client
path_of_union = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmp.xlsx'
path = Path("tmpXLS")
min_excel_file_size = 100


def union():
    df = pd.concat([pd.read_excel(f)
                    for f in path.glob("*.xlsx")
                    if f.stat().st_size >= min_excel_file_size],
                   ignore_index=True)
    df.to_excel(path_of_union)
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_of_union)
    sheet = wb.ActiveSheet
    sheet.columns("A").Delete()
    wb.Save()
    wb.Close()
    Excel.Quit()
    return
