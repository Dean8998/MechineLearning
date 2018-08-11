from openpyxl import Workbook
from openpyxl import load_workbook
import openpyxl
import pandas as pd

target = load_workbook('target.xlsx')

target_sheet1 = target["1111"]
# target_sheet2 = target["2222"]
# print(target_sheet1["A1"].value)

source_file_dir = target_sheet1["B2"].value

source = load_workbook(filename=source_file_dir)

for names in target.sheetnames:
    ws = target[names]
    key = ws["B3"].value
    if key == "sh":
        source_sheet = source["sh"]
        print(source_sheet["A2"].value)

    elif key == "wy":
        source_sheet = source["wy"]
        print(source_sheet["A2"].value)
    else:
        continue





