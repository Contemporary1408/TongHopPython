#Basic concept:
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

# ############################################################################################
#Full script to check whether SAP is opened or not, if opened then open new window:
# Name this script below as GLoginSAP.py:
import os
import win32com.client
import time
import psutil
#tcode = "ZGM07" #Thay đổi tcode
def GSAP(tcode,id,pw,con):
    flag = True
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'] == 'saplogon.exe':
            flag = False
    #Nếu đã đăng nhập SAP:
    if flag == False:      
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + tcode
        session.findById("wnd[0]").sendVKey(0)
        return session
    #Nếu chưa đăng nhập SAP:    
    else:    
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        os.system('start "" "' + path + '"')
        time.sleep(1.5)
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(con, True)
        session = connection.Children(0)
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = id
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pw
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"+ tcode
        session.findById("wnd[0]").sendVKey(0)
        return session
        
# In other script, to reuse it:
from GLoginSAP import GSAP
session = GSAP(tcode="ZGM07",id="VHACS12",pw="btmv@140894",con="BTMV Production Server")
#then run code from here:
session.findById("wnd[0]/usr/ctxtP_BUKRS").Text = "1000"
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "1000"
session.findById("wnd[0]/usr/ctxtS_PO_DAT-LOW").Text = "01.01.2020"
session.findById("wnd[0]/usr/ctxtS_PO_DAT-HIGH").Text = "30.08.2023"


# For SAP Business Client ##################################################################
# Importing the Libraries
import win32com.client
import subprocess
import time
# This function will Login to SAP from the SAP Logon window
def saplogin():
        path = r"C:\Program Files\SAP\NWBC770\NWBC.exe"
        subprocess.Popen(path)
        time.sleep(13)
        SapGuiAuto = win32com.client.GetObject("SAPGUISERVER")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "VHACS12"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "btmv@140894"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize()
saplogin()
