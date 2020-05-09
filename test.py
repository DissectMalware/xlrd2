import xlrd2

path = r"C:\Users\user\Downloads\bf58dc1c6ee61d7370c3dfaed7efd98435aed215dfed58e7d90a25b195584b33.xls"
xl_workbook = xlrd2.open_workbook(path)

for sheet in xl_workbook.sheets():
    if sheet.boundsheet_type == xlrd2.biffh.XL_MACROSHEET:
        print(sheet.name)
        for row in sheet.get_rows():
            for cell in row:
                if cell.formula is not None and len(cell.formula)>0:
                    print("({},{}):\t{},\t{}".format(cell.row, cell.column, cell.formula, cell.value))

        # for row in sheet.get_rows():
        #     for cell in row:
        #         if cell.formula is None or len(cell.formula)==0:
        #             print("({},{}):\t\t,\t{}".format(cell.row, cell.column, cell.value))









