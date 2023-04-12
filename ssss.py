from docxtpl import DocxTemplate
from docx2pdf import convert
import os  
import openpyxl as op
import ReadExcel as rxls


names = rxls.reader("values.xlsx")
print()

for i in range(len(names)):
    name = names[i]
    present_name = "Alan Perez"
    date = "12/2/2023"
    print(name)
    doc = DocxTemplate("plantilla.docx")
    context = { 'Name' : name}
    doc.render(context)
    doc.save(f"{i}.docx")
    convert(f"{i}.docx")
    os.system(f"del {i}.docx")

  