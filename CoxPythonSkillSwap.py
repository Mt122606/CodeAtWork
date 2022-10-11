import xlrd

from pathpib import PureWindowPath

file = 'bom-c1761186-cb_-_att_pre_build (1).xlsx'

fullpath = PureWindowPath(f"C:\Users\B99327\OneDrive - Cox Communications\Documents\CodeAtWork")
#iterate workbook 
with xrld.open_workbook(fullpath,formatting_info=False) as book:
    sheet_names = book.sheet_names()
    for sheet_name in book.sheet_names():
        print(sheet_name)
        sheet =book.sheet_by_name(sheet_name)
        number_of_rows =sheet.nrows
        for current_row in range(number_of_rows):
            cells=sheet.row(current_row)[0:3]
            for cell in cells:
                print(cell.value)