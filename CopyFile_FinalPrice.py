import tkinter as tk
from tkinter import filedialog
import openpyxl as xl
import ctypes
import os
import shutil

def get_value(dictionary, string):
    for key, value in dictionary.items():
        if string in key:
            return value
    return None
root = tk.Tk()
root.withdraw()
temp_path = filedialog.askopenfilename(title="Chọn đường dẫn template Final price Excel:")
wb = xl.load_workbook(temp_path,data_only=True)
ws = wb.active
lrow = ws.max_row
tempdir = os.path.dirname(temp_path)
for i in range(2,lrow+1):
    foldname = str(ws.cell(row=i,column=4).value.date().strftime("%d-%b"))
    new_fol = tempdir + "/" + foldname
    if not os.path.exists(new_fol):
        os.mkdir(new_fol)
dict1={}
for i in range(2,lrow+1):
    foldname = str(ws.cell(row=i,column=4).value.date().strftime("%d-%b"))
    filename = "Inv " + str(ws.cell(row=i,column=1).value) + " (final price).xlsx"
    dict1[filename] = foldname
for path, dirs, files in os.walk(tempdir):
    for file in files:
        if file in dict1:
            src_path = os.path.join(path, file).replace("\\","/")
            dest_path = os.path.join(tempdir, get_value(dict1,file),file).replace("\\","/")
            try:
                shutil.copyfile(src_path, dest_path)

            # If source and destination are same
            except shutil.SameFileError:
                print("Source and destination represents the same file.")
ctypes.windll.user32.MessageBoxW(0, "Done, 520!", "Thông báo", 0)
