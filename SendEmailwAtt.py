import win32com.client as win32
import pandas as pd
import os
# Function to create personalized email body with kwargs
def create_email_body(**kwargs):
    body = f"Dear {kwargs.get('name')},\n\n"
    body += kwargs.get('custom_message')
    body += f"\nBest regards,\n{kwargs.get('sender')} - Accounting & Finance Dept. \nDect: 51624"
    return body
def send_email():
    # Load the Excel file with recipient details
    mail_list = r"C:\Users\anh.doduc\Desktop\Mails\Mail list.xlsx"  # Update with the path to your Excel file
    recipients_df = pd.read_excel(mail_list)
    # Loop through the DataFrame and send an email to each recipient
    for index, row in recipients_df.iterrows():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = row['EmailTo']  # Column name in Excel file containing email addresses
        mail.Cc = row['EmailCc']
        mail.Subject = row['Subject']
        #mail_dict[row[]]
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
    
        mail.Display()
        #mail.Send()
