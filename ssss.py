from docxtpl import DocxTemplate
from docx2pdf import convert
import os   


# doc = DocxTemplate("value.docx")
# context = { 'Name' : "World company" }
# doc.render(context)
# doc.save("generated_doc.docx")

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