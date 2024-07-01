import pandas as pd
import openpyxl as xl

# Get max row index:
df = pd.read_excel(r"C:\Users\Contemporary\Desktop\exceltest.xlsx",sheet_name='Sheet1')  # Replace with your sheet name
column_name = 'Index Date'
max_row = len(df[column_name].dropna()) + 1

#Read file:
wb = xl.load_workbook(r'C:\Users\Contemporary\Desktop\exceltest.xlsx')
ws = wb.active
for x in range(2,max_row+1):
    frow = ws.cell(row = x, column=6).value
    lrow = ws.cell(row = x, column=7).value
    date = ws.cell(row = x, column=4).value.strftime("%d.%m.%Y")
    inv =  ws.cell(row = x, column=5).value
    list=()
    for y in range(frow,lrow+1):
