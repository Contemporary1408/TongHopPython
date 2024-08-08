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
