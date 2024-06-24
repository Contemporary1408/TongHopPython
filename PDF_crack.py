import pikepdf
with pikepdf.open(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\99.Personal files\Duc Anh\V-RE DR 025.pdf") as pdf:
    num_pages = len(pdf.pages)
    del pdf.pages[-1]
    pdf.save(r"\\10.118.29.7\BTMV-Data\4-ACCOUNTING\99.Personal files\Duc Anh\V-RE DR 025aa.pdf")
