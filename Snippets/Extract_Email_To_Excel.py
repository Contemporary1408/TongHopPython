import pandas as pd

# Example list of emails separated by semicolons
email_list = "john@example.com;alice@example.com;bob@example.com"

# Split the email list into separate rows
emails = email_list.split(";")

# Create a DataFrame with the emails
df = pd.DataFrame(emails, columns=["Email"])

# Save the DataFrame to an Excel file
df.to_excel("D:\SSSemails.xlsx", index=False)

# Print a success message
print("The emails are successfully saved to emails.xlsx file.")

############################################################################################

import re
setemail = "Nguyen Thi Thu Huong <Huong.Nguyen1@bridgestone.com>; Nguyen Thi Dung 2 <dung.nguyen2@bridgestone.com>; Dong Thi Ninh <dongthininh@bridgestone.com>; Gabriel Goh <gabriel.goh@bridgestone.com>; Le Tien Hung <letienhung@bridgestone.com>; Kelly Qi <kelly.qi@bridgestone.com>; Jia Min Shum <jiamin.shum@bridgestone.com>; Pham Thi Phuong Mai <mai.pham1@bridgestone.com>; Nguyen Thi Ngoc Quynh <nguyenthingocquynh@bridgestone.com>; Nguyen Thi Quynh Trang <trang.nguyen3@bridgestone.com>; Usa Kijsawangsak <usa.kijsawangsak@bridgestone.com>; Do Duc Anh <do.anh@bridgestone.com>"
matches = re.findall(r'<(.*?)>', setemail)
listemail=""
for match in matches:
    listemail = listemail + match + ";"
print(listemail)
