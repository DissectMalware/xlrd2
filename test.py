import xlrd2
import sys

if len(sys.argv) > 1:
    path = sys.argv[1]
    xl_workbook = xlrd2.open_workbook(path)

    for sheet in xl_workbook.sheets():
        if sheet.boundsheet_type == xlrd2.biffh.XL_MACROSHEET:
            print(sheet.name)
            for row in sheet.get_rows():
                if row is not None:
                    for cell in row:
                        if cell.formula is not None and len(cell.formula)>0:
                            print("({},{}):\t{},\t{}".format(cell.row, cell.column, cell.formula, cell.value))








