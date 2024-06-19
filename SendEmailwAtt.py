import win32com.client as win32
import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox
from winotify import Notification
import shutil
home_directory = os.path.expanduser('~').replace("\\","/")
downloads_path = os.path.join(home_directory, 'Downloads').replace("\\","/")
src_path = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\27.Budget report\Act vs budget report\Automate\old\sua.ico"
ico_path = os.path.join(downloads_path, "sua.ico").replace("\\","/")
shutil.copyfile(src_path, ico_path)
toast = Notification(app_id="App gửi mail tự động by Do Duc Anh",
                     title="Thông báo",
                     msg="Đã xong!",
                     icon=ico_path)
flag = False
def ask_yes_no():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    return messagebox.askyesno("Thông báo", "Xem lại và gửi bằng tay?")
# Function to create personalized email body with kwargs
def create_email_body(**kwargs):
    body = f"Dear {kwargs.get('name')},\n\n"
    body += kwargs.get('custom_message')
    body += f"\nBest regards,\n{kwargs.get('sender')} - Accounting & Finance Dept. \nDect: 51624"
    #body += f"\nThis automated email is powered by Python"
    return body
def send_email():
    # Load the Excel file with recipient details
    mail_list = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\27.Budget report\Act vs budget report\Automate\Mail list.xlsx"  # Update with the path to your Excel file
    recipients_df = pd.read_excel(mail_list)
    # Loop through the DataFrame and send an email to each recipient
    for index, row in recipients_df.iterrows():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = row['EmailTo']  # Column name in Excel file containing email addresses
        mail.Cc = row['EmailCc']
        mail.Subject = row['Subject']

        # Create a personalized email body using kwargs
        mail.Body = create_email_body(
            name=row['Name'],  # Assuming 'Name' is a column in your Excel file
            custom_message=f"{row['Body']}",  # Example of using another column
            sender = row['Sender']
        )

        # Add attachment
        src_folder = os.path.dirname(mail_list)
        for path, dirs, files in os.walk(src_folder):
            for file in files:
                if str(row['Code']) in file:
                    attachment_path = os.path.join(path, file).replace("\\","/")
                    mail.Attachments.Add(Source=attachment_path)
        if flag == False:                    
            mail.Display()
        else:
            mail.Send()
if ask_yes_no():
    flag = False
else:
    flag = True    
send_email()
toast.show()
