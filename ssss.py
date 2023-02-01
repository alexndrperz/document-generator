from docxtpl import DocxTemplate
from docx2pdf import convert
import os  
import openpyxl as op
import ReadExcel as rxls


m = rxls.reader("values.xlsx")
print()

fil = op.load_workbook("values.xlsx")
sheet = fil.get_active_sheet()
for cellObj in sheet.columns[0]:
    print(cellObj.value)


values = pandas.read_excel('values.xlsx', index_col=0)

print(values['Juan'])
des = 1 
while True:
    
    print("Deposite los datos solicitados")
    print()

    name = input("Nombre: ")
    present_name = input("Nombre Presentador: ")
    date = input("Fecha: ")
    
    doc = DocxTemplate("value.docx")
    context = { 'Name' : name, 'Presentator' : present_name, 'Date' : date }
    doc.render(context)
    doc.save(f"{name}.docx")
    convert(f"{name}.docx")
    
    os.system(f"\"{name}.pdf\"")

    des = int(input("""¿Desea proseguir?
    1) Si
    2) No
    """))

    if des == 2:
        context = {}
        break