import tkinter as tk
from PyPDFForm import PdfWrapper
from tkinter.filedialog import askopenfilename
window = tk.Tk()
window.title("A simple window")
window.geometry("400x400")


def save_textbox_input():
    tk.Tk().withdraw()
    filename = askopenfilename()
    file1 = filename[:-4]+'-done'+".pdf"
    sender = namebox.get("1.0",'end-1c')
    shipper = namebox1.get("1.0",'end-1c')

    filled = PdfWrapper(filename).fill(
        {
            "Given Name Text Box": str(sender),
            "Family Name Text Box": str(shipper),
        },
    )
    with open(file1, "wb+") as output:
        output.write(filled.read())
    text_box.config(text="Done,520")


button = tk.Button(text="Chọn file và nhập", command=save_textbox_input)
namebox = tk.Text(width="30", height="2")
namebox1 = tk.Text(width="30", height="2")
text_box = tk.Label(width="30", height="2")

namebox.pack()
namebox1.pack()
button.pack()
text_box.pack()
tk.mainloop()
