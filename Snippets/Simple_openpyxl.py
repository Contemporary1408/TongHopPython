from openpyxl import Workbook, load_workbook
path = "D:/Test.xlsx"
wb = load_workbook(path)

# grab the active worksheet
ws = wb.active
ws2 = wb["Sheet2"]

# Data can be assigned directly to cells
ws2['A1'] = 42
# or:
ws2.cell(row=1,column=1) = 42

# Rows can also be appended
ws2.append([1, 2, 3])

# Python types will automatically be converted
import datetime
ws2['A2'] = datetime.datetime.now()

# Write formula to a cell:
ws['A3'] = '=SUM(A1:A2)'

# Save the file
wb.save(path)
