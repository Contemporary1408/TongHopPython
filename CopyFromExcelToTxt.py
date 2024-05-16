import pandas as pd
excel = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS Template.xlsm"
csv = r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\10.G-APICS\Automate GAPICS\GAPICS upload.tsv"
# Read the Excel file (From A3 to last row, exclude Unnamed column)
df = pd.read_excel(excel, sheet_name='SAP data', header=1, usecols=lambda x: 'Unnamed' not in x)
# Write to a text file (change the separator if needed)
df.to_csv(csv, sep='\t', index=False,header=False)
