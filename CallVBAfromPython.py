import os, os.path
import win32com.client

if os.path.exists("D:/Vendor list-basic.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(os.path.abspath("D:/Vendor list-basic.xlsm"))
    #xl.VBE.VBProjects(1).VBComponents.Import('Vendor list-basic.xlsm/UserForm1.frm')
    #xl.VBE.VBProjects(1).VBComponents.Import('Vendor list-basic.xlsm/Module2.bas')
    #xl.Run('ShowUserForm')
    xl.Application.Run("Module2.ShowUserForm")
    #wb.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    #xl.Quit() # Comment this out if your excel script closes
    del xl
