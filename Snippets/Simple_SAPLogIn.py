import os
import win32com.client
import time
path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
os.system('start "" "' + path + '"')
time.sleep(1)
SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine
connection = application.OpenConnection("BTMV Production Server", True)
session = connection.Children(0)
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "VHACS12"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "btmv@140894"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").maximize()
