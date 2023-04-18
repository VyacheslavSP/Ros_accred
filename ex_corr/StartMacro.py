import xlwings as xw


def start_macro(path):
    vba_book = xw.Book(path)
    vba_macro = vba_book.macro("exportXML_standalone")
    vba_macro()
