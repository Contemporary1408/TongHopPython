#Dùng để xuất thành từng file excel đối với mỗi giá trị lọc của 1 cột trong bảng
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import ctypes

root = tk.Tk()
root.withdraw()
input_file = filedialog.askopenfilename(title="Chọn đường dẫn template Final price Excel:")
filedir = os.path.dirname(input_file)
new_fol = filedir + "/" + "filtered_output_python" #Directory to new folder
if not os.path.exists(new_fol): #Check folder exist then create new folder
    os.mkdir(new_fol)
# Load your Excel file into a DataFrame
#input_file = r"D:\Labs\excelgen\Shipment List1.xlsx"  # Sửa lại tên file
df = pd.read_excel(input_file)

# Specify the column to filter on
filter_column = 'STOCK POINT'  # Sửa lại tên cột

# Get all unique values in the specified column
unique_values = df[filter_column].unique()

# Create a new Excel file for each unique value
for value in unique_values:
    filtered_df = df[df[filter_column] == value]
    
    # Define output file name based on the unique value
    output_file = new_fol + "/" + f'filtered_output_{value}.xlsx'  # Chỗ này cần khai báo đường dẫn cụ thể thì mới lưu file được
    filtered_df.to_excel(output_file, index=False)

ctypes.windll.user32.MessageBoxW(0, "Done, 520! Xem file vừa tạo mới trong folder 'filtered_output_python'", "Thông báo", 0)
