import datetime
#now = datetime.datetime.now()
now = datetime.datetime(2025,1,15)
last_month = now.month-1 if now.month > 1 else 12
year = str(now.year-1) if now.month==1 else str(now.year)
month = datetime.date(1900, last_month, 1).strftime('%B')
path = "//10.118.29.7/BTMV-Data/4-ACCOUNTING/10.G-APICS/Automate GAPICS/GAPICS "
path1 = path + month + year + ".xlsx"
print(path1)
