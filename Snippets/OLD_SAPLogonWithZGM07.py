# Importing the Libraries
import win32com.client
import sys
import subprocess
import time


# This function will Login to SAP from the SAP Logon window

def saplogin():

    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(1)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection("BTMV Production Server", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "VHACS12"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "ducanh@16119"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "zgm07"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtP_BUKRS").text = "1000"
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1000"
        session.findById("wnd[0]/usr/radP_POBASE").setFocus()
        session.findById("wnd[0]/usr/radP_POBASE").select()
        session.findById("wnd[0]/usr/ctxtS_PO_DAT-LOW").text = "010424"
        session.findById("wnd[0]/usr/ctxtS_PO_DAT-HIGH").text = "300424"
        session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").text = ""
        session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").setFocus()
        session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").caretPosition = 0
        session.findById("wnd[0]").sendVKey(8)
    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


saplogin()
