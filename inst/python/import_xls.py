import xlrd
def import_xls(x):
    return xlrd.open_workbook(filename=x, formatting_info=True)
