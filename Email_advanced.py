#File json trong folder assets
import FreeSimpleGUI as sg
import json
from datetime import date, timedelta
import win32com.client as win32
t = date.today().strftime('%d %b')
t1 = date.today()+ timedelta(days=1)
t1 = t1.strftime('%d %b')
# ##################################################################################
# Define your list of choices for the combo box
choices = ['Mẫu 1 - Vanning', 'Mẫu 2 - Hàng xuất', 'Mẫu 3 - Đóng hàng']
# Create the layout with the combo box and OK button
layout = [
    [sg.Text('Chọn mẫu Email:')],
    [sg.Combo(choices, key='combo', size=(50, 1), enable_events=True)],
    [sg.Button('OK')]
]
# Create the window
window = sg.Window('Tool tạo mail tự động by Do Duc Anh', layout, finalize=True,icon="sua.ico")
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    elif event == 'OK':
        chosen_value = values['combo']
        break
window.close()
# ##################################################################################
with open("mail_loop_adv.json",encoding='utf-8') as file: #Thêm encoding để nhận ký tự unicode
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
    mail.Subject = f"VANNING // ĐÓNG HÀNG FGWH {t}"
    mail.HTMLBody = existing_signature +f'Dear Jupiter team,<br><br>Em gửi thông tin cont seal lô Jupiter chạy '+ mail.HTMLbody[index + 1:]
    mail.Display()
elif chosen_value == 'Mẫu 2 - Hàng xuất':
    mail.To = nested_data['Mau2']['EmailTo']
    mail.Cc = nested_data['Mau2']['EmailCc']
    mail.Subject = f"HÀNG XUẤT THÀNH PHẨM .... LOADING {t}"
    mail.HTMLBody = existing_signature +f'Dear Jupiter team,<br><br>Em gửi lô hàng 4705S đóng hàng {t} '+ mail.HTMLbody[index + 1:]
    mail.Display()
else: 
    mail.To = nested_data['Mau3']['EmailTo']
    mail.Cc = nested_data['Mau3']['EmailCc']
    mail.Subject = f"KẾ HOẠCH ĐÓNG HÀNG NGÀY {t1}"
    mail.HTMLBody = existing_signature +f'Dear Jupiter team,<br><br>Em gửi kế hoạch đóng hàng ngày {t1} '+ mail.HTMLbody[index + 1:]
    mail.Display()
