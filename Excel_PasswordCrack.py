from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
from winotify import Notification
import os
import shutil
import ctypes
#Dành tặng Sứa.
home_directory = os.path.expanduser('~').replace("\\","/")
downloads_path = os.path.join(home_directory, 'Downloads').replace("\\","/")
desktop_path = os.path.join(home_directory, 'Desktop').replace("\\","/")
src_path = r"\\10.118.29.7\BTMV-Data\1-ALL\05. 1611\asset\sua.ico"
ico_path = os.path.join(downloads_path, "sua.ico").replace("\\","/")
shutil.copyfile(src_path, ico_path)
toast = Notification(app_id="Tool remove pw Excel v1.0 by Do Duc Anh",
                     title="Thông báo",
                     msg="Đã phá pw thành công, xem file Cracked ngoài Desktop!",
                     icon=ico_path)
def crack_excel():
    root = tk.Tk()
    root.withdraw()
    excel_path = filedialog.askopenfilename(title="Chọn file Excel:")
    try:
        sample = load_workbook(filename=excel_path)
        for sheet in sample: 
            sheet.protection.disable()
        cracked_name = os.path.basename(excel_path)
        crack_name1 = cracked_name[:-5] + "-CRACKED.xlsx"
        cracked_path = os.path.join(desktop_path,crack_name1).replace("\\","/")
        sample.save(filename=cracked_path)
        sample.close()
        toast.show()
    except:
        ctypes.windll.user32.MessageBoxW(0, "Bạn chưa chọn file hoặc \nFile Excel bị khóa Workbook! \n**Chương trình chỉ hỗ trợ phá pw cho Worksheet.", "Tool remove pw Excel v1.0 by Do Duc Anh", 0)
                 
crack_excel()
