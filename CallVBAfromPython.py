import os, os.path
import win32com.client

if os.path.exists("D:/python2.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    wb = xl.Workbooks.Open(os.path.abspath("D:/python2.xlsm"))
    xl.Application.Run("python2.xlsm!Module1.prelim")
    wb.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Quit() # Comment this out if your excel script closes
    del xl
