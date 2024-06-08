#Ref https://chinapandaman.github.io/PyPDFForm/fill/
# #################################################
# Firstly run this code to get json structure of PDF:
import json
from PyPDFForm import PdfWrapper
pdf_form_schema = PdfWrapper("sample_template.pdf").schema
print(json.dumps(pdf_form_schema, indent=4, sort_keys=True))
# After that, copy output to a notepad as reference for field names:
# Run this code below to set input value for each field:
from PyPDFForm import PdfWrapper
filled = PdfWrapper(r"C:\Users\anh.doduc\Desktop\OoPdfFormExample.pdf").fill(
    {
        "Given Name Text Box": "GVname",
        "Family Name Text Box": "Fname",
    },
)
with open(r"C:\Users\anh.doduc\Desktop\OoPdfFormExample1.pdf", "wb+") as output:
    output.write(filled.read())
