import ctypes
ctypes.windll.user32.MessageBoxW(0, "Content here", "Title here", 0)


import os
home_directory = os.path.expanduser('~').replace("\\","/")
downloads_path = os.path.join(home_directory, 'Downloads').replace("\\","/")
desktop_path = os.path.join(home_directory, 'Desktop').replace("\\","/")
src_path = r"\\10.118.29.7\BTMV-Data\1-ALL\05. 1611\asset\sua.ico"
ico_path = os.path.join(downloads_path, "sua.ico").replace("\\","/")
shutil.copyfile(src_path, ico_path)
