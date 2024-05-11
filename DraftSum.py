import flet as ft

def main(page):
    def btn_click(e):
        if not txt_name1.value:
            txt_name1.error_text = "Nhập đúng định dạng số!!"
            page.update()
        else:
             name1 = txt_name1.value
             name2 = txt_name2.value
             c = int(name1) + int(name2)
             page.clean()
             page.add(ft.Text(f"Tổng cần tính là: {c}",color="Brown"))

    txt_name1 = ft.TextField(label="Số thứ 1",color="Red",width=300)
    txt_name2 = ft.TextField(label="Số thứ 2",color="Cyan",width=300)

    page.add(txt_name1)
    page.add(txt_name2, ft.ElevatedButton("Tính tổng!", on_click=btn_click))
ft.app(target=main)
