#Giao diện tính tổng đơn giản bằng Flet
import flet as ft

def main(page):
    page.title = "Flet counter example"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    txt_name1 = ft.TextField(label="Số thứ 1", color="Red", width=300)
    txt_name2 = ft.TextField(label="Số thứ 2", color="Cyan", width=300)
    def btn_click(e):
        if not txt_name1.value:
            txt_name1.error_text = "Nhập đúng định dạng số!!"
            page.update()
        else:
             name1 = txt_name1.value
             name2 = txt_name2.value
             c = int(name1) + int(name2)
             page.clean()
             page.add(ft.Row([ft.Text(f"Tổng cần tính là: {c}",color="Brown")],alignment=ft.MainAxisAlignment.CENTER))

    page.add(ft.Row([txt_name1],alignment=ft.MainAxisAlignment.CENTER))
    page.add(ft.Row([txt_name2],alignment=ft.MainAxisAlignment.CENTER))
    page.add(ft.Row([ft.ElevatedButton("Tính tổng!", on_click=btn_click)],alignment=ft.MainAxisAlignment.CENTER))
ft.app(target=main)
