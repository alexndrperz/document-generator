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
    
    doc = DocxTemplate("value.docx")
    context = { 'Name' : name, 'Presentator' : present_name, 'Date' : date }
    doc.render(context)
    doc.save(f"{name}_{i}.docx")
    convert(f"{name}_{i}.docx")
    os.system(f"del {name}_{i}.docx")

    os.system(f"{name}_{i}.pdf")

  