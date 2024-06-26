# With simple non-HTML outlook email
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


# *******************************************************************************************************************************
# With HTML support changing format, add signature,...
import win32com.client as win32

def add_signature_to_email(message, subject, recipient):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject

    # Get the existing email signature
    mail.GetInspector()
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    existing_signature = mail.HTMLbody[:index + 1]

    # Your custom message
    custom_message = "Hello there! This is a test email."
    #custom_message = "<p>This is a <i>italic</i> message.</p>" #Italic
    #custom_message = "<p>This is a <b>bold</b> message.</p>" #Bold
    #custom_message = "<p>This is a <u>underlined</u> message.</p>" #Underlined
    #custom_message = "<p>This is a message.<br> </p>" #Add <br> at end of line to break line
    body += f"<p style='font-size: 10px;'>This automated email is powered by Python.</p>" #Set font size
  
    # Combine the existing signature and your custom message
    mail.HTMLbody = existing_signature + custom_message + mail.HTMLbody[index + 1:]

    # Display the email (you can uncomment 'mail.send' to actually send it)
    mail.Display(True)

# Example usage:
email_subject = "Important Update"
recipient_email = "recipient@example.com"
add_signature_to_email("Your custom message here.", email_subject, recipient_email)
