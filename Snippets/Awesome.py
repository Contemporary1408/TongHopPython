# Msgbox like VBA
import ctypes
ctypes.windll.user32.MessageBoxW(0, "Content here", "Title here", 0)

# Get path of any user
import os
home_directory = os.path.expanduser('~').replace("\\","/")
downloads_path = os.path.join(home_directory, 'Downloads').replace("\\","/")
desktop_path = os.path.join(home_directory, 'Desktop').replace("\\","/")
src_path = r"\\10.118.29.7\BTMV-Data\1-ALL\05. 1611\asset\sua.ico"
ico_path = os.path.join(downloads_path, "sua.ico").replace("\\","/")
shutil.copyfile(src_path, ico_path)

# Check today is Monday?
from datetime import date
if date.today().isoweekday() == 5:
    print("Yes, today is Monday")
else:
    print("Nope...")
