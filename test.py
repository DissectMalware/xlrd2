import xlrd2
import sys

if len(sys.argv) > 1:
    path = sys.argv[1]
    xl_workbook = xlrd2.open_workbook(path)
    defined_names = xl_workbook.name_map
    for sheet in xl_workbook.sheets():
        if sheet.boundsheet_type == xlrd2.biffh.XL_MACROSHEET:
            print(sheet.name)
            for cell in sheet.get_used_cells():
                if cell.formula is not None and len(cell.formula)>0:
                    print("({},{}):\t{},\t{}".format(cell.row, cell.column, cell.formula, cell.value))








