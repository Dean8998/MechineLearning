import xlrd
import xlwt
from xlutils.copy import copy

target_workbook = xlrd.open_workbook("target.xlsx")
target_copy = copy(target_workbook)

sheet1 = target_workbook.sheet_by_name("1111")

source_dir = sheet1.cell(1, 1).value
# print(source_dir)

source_workbook = xlrd.open_workbook(source_dir)
source_sheet1 = source_workbook.sheet_by_index(0)
print(source_sheet1.cell(1, 0).value)

print(source_sheet1.ncols, source_sheet1.nrows)



for names in target_workbook.sheet_names():
    sheet = target_workbook.sheet_by_name(names)
    key = sheet.cell_value(2,1)
    if key == "sh":
        source_sheet = source_workbook.sheet_by_name("sh")
        ws = target_copy.get_sheet("sh")

        for i in range(9):
            for j in range(4, source_sheet.nrows):
                value = source_sheet.cell_value(j, i)

    elif key == "wy":
        source_sheet = source_workbook.sheet_by_name("wy")
    else:
        continue

