import xlrd2

path = r"C:\Users\user\Downloads\bf58dc1c6ee61d7370c3dfaed7efd98435aed215dfed58e7d90a25b195584b33.xls"
xl_workbook = xlrd2.open_workbook(path)

for sheet in xl_workbook.sheets():
    if sheet.boundsheet_type == xlrd2.biffh.XL_MACROSHEET:
        print(sheet.name)
        for cell, formula in sheet.formula_map.items():
            print("{}: {}".format(cell, formula))







