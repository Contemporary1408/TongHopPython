import os, os.path
import win32com.client
from xlsxwriter.workbook import Workbook
import ctypes
import csv
import pandas as pd
#Run VBA ket xuat tsv tu SAP:
def GAPIC():
    print("Chương trình tạo file GAPICS tự động v1.0 - (c) 2024 by Đỗ Đức Anh")
    ctypes.windll.user32.MessageBoxW(0, "Hãy thoát toàn bộ các file Excel đang mở trước khi tiếp tục!", "Lưu ý", 0)
    print("Vui lòng chờ trong lúc chương trình thực hiện....")
    print("Đang thoát toàn bộ cửa sổ SAP đang mở...")
    os.system("taskkill /im saplogon.exe")
    if os.path.exists(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"):
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        print("Đang mở đường dẫn đến file GAPICS: 10.118.29.7/BTMV-Data/4-ACCOUNTING/10.G-APICS/Automate GAPICS")
        wb = xl.Workbooks.Open(
            os.path.abspath(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"))
        print("THÔNG BÁO: Đang khởi tạo môi trường SAP - 20%")
        xl.Application.Run("Module4.SAP1")
        wb.Save()
        xl.Quit()
        del xl
        print('THÔNG BÁO: Kết xuất số liệu từ SAP thành công! - 30%')
        # Convert file tsv thanh xlsx
        tsv_file = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\ZGAPICS01.tsv"
        xlsx_file = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\ZGAPICS01.xlsx"
        workbook = Workbook(xlsx_file)
        worksheet = workbook.add_worksheet()
        tsv_reader = csv.reader(open(tsv_file, 'r'), delimiter='\t')
        print('THÔNG BÁO: Đang xử lý file tsv vừa kết xuất - 50%')
        for row, data in enumerate(tsv_reader):
            worksheet.write_row(row, 0, data)
        workbook.close()
    print('THÔNG BÁO: Convert file tsv sang excel thành công! - 70%')
    if os.path.exists(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"):
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        wb = xl.Workbooks.Open(
            os.path.abspath(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"))
        print('THÔNG BÁO: Đang tiến hành các bước xử lý dữ liệu cuối cùng - 80%')
        xl.Application.Run("Module1.GAPICS")
        print('THÔNG BÁO: Đang lưu dữ liệu... - 90%')
        wb.Save()
        xl.Quit()
        del xl
    print('THÔNG BÁO: Lưu dữ liệu thành công! - 95%')
    excel = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"
    tsv = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS upload.tsv"
    df = pd.read_excel(excel, sheet_name='SAP data',header=None)
    print(df.head())
    print('THÔNG BÁO: Đang tạo file tsv để upload - 99%')
    df.to_csv(tsv)
    # Read the Excel file
    df = pd.read_excel(excel, sheet_name='SAP data', header=1, usecols=lambda x: 'Unnamed' not in x)
    # Write to a text file (change the separator if needed)
    df.to_csv(tsv, sep="\t", index=False  ,header=False,float_format='%.0f')
    ctypes.windll.user32.MessageBoxW(0, "Chương trình đã thực hiện xong!", "Thông báo", 0)
    os.system("taskkill /im saplogon.exe")
GAPIC()
