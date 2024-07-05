#File json trong folder assets
import FreeSimpleGUI as sg
import json
from datetime import date, timedelta
import win32com.client as win32
from suacopy import suacopy
ico_path =suacopy()
t = date.today().strftime('%d %b')
t1 = date.today()+ timedelta(days=1)
t1 = t1.strftime('%d %b')
# ##################################################################################
# Tạo combo box:
choices = ['Mẫu 1 - Vanning', 'Mẫu 2 - Tờ khai', 'Mẫu 3 - Kế hoạch']
# Dành tặng Sứa và mẹ bé Sứa
# Tạo layout cửa sổ:
layout = [
    [sg.Text('Chọn mẫu Email:')],
    [sg.Combo(choices, key='combo', size=(50, 1), enable_events=True)],
    [sg.Button('OK')]
]
# Create the window
window = sg.Window('Tool tạo mail tự động by Do Duc Anh', layout, finalize=True,icon=ico_path)
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
       break
    elif event == 'OK':
        chosen_value = values['combo']
        # ##################################################################################
        with open("mail_loop_adv.json",encoding="utf-8") as file:
            nested_data = json.load(file)
        # ##################################################################################
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.GetInspector()
        index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
        existing_signature = mail.HTMLbody[:index + 1]
        if chosen_value == 'Mẫu 1 - Vanning':
            mail.To = nested_data['Mau1']['EmailTo']
            mail.Cc = nested_data['Mau1']['EmailCc']
            mail.Subject = f"{nested_data['Mau1']['Subject']} {t}"
            mail.HTMLBody = existing_signature +f"{nested_data['Mau1']['Dear']} "+ mail.HTMLbody[index + 1:]
            mail.Display()
        elif chosen_value == 'Mẫu 2 - Tờ khai':
            mail.To = nested_data['Mau2']['EmailTo']
            mail.Cc = nested_data['Mau2']['EmailCc']
            mail.Subject = f"{nested_data['Mau2']['Subject']} {t}"
            mail.HTMLBody = existing_signature +f"{nested_data['Mau2']['Dear']} {t} "+ mail.HTMLbody[index + 1:]
            mail.Display()
        else: 
            mail.To = nested_data['Mau3']['EmailTo']
            mail.Cc = nested_data['Mau3']['EmailCc']
            mail.Subject = f"{nested_data['Mau3']['Subject']} {t1}"
            mail.HTMLBody = existing_signature +f"{nested_data['Mau3']['Dear']} {t1} "+ mail.HTMLbody[index + 1:]
            mail.Display()
window.close()
