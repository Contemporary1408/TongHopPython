if __name__ == '__main__': # to directly run script
    
# Common tag for HTML:
    <a>: Creates hyperlinks (anchors).
    <ul>: Defines an unordered list (bulleted list).
    <ol>: Defines an ordered list (numbered list).
    <li>: Represents a list item within <ul> or <ol>.
    <img>: Embeds images.
    <br>: Inserts a line break (empty element).
    <strong> or <b>: Makes text bold.
    <em> or <i>: Makes text italic.
    <p>This is a paragraph of text.</p>
    <h1>This is the main heading</h1>
    <h2>This is a subheading</h2>
    
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
if date.today().isoweekday() == 1:
    print("Yes, today is Monday")
else:
    print("Nope...")

#Formate date https://www.tutorialspoint.com/How-to-get-formatted-date-and-time-in-Python
from datetime import date
t = date(2024,6,17) #assign 17/6/2024 to var t
a = date.today().strftime("%d-%m-%y") # Output 21-06-24
a = date.today().strftime("%d-%b-%y") # Output 21-Jun-24
a = date.today().strftime("%a-%m-%y") # Output Fri-06-24

#Nested list: https://www.learnbyexample.org/python-nested-list/
L = ['a', 'b', ['cc', 'dd', ['eee', 'fff']], 'g', 'h']
print(L[2])         # Output: ['cc', 'dd', ['eee', 'fff']]
print(L[2][2])      # Output: ['eee', 'fff']
print(L[2][2][0])   # Output: eee

#Get max row index of a column:
import pandas as pd
df = pd.read_excel(r"C:\Users\Contemporary\Desktop\exceltest.xlsx",sheet_name='Sheet1')  # Replace with your sheet name
column_name = 'Index Date'
max_row = len(df[column_name].dropna()) + 1
