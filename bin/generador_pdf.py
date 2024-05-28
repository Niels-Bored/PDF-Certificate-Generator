import os
import io
import textwrap
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime, timedelta

current_folder = os.path.dirname (__file__)
parent_folder = os.path.dirname (current_folder)
files_folder = os.path.join (parent_folder, "files")
data = os.path.join (files_folder, f"Data.xlsx")
original_pdf = os.path.join (current_folder, f"Certificado.pdf")

def generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('times','times.ttf'))
    pdfmetrics.registerFont(TTFont('timesbd', 'timesbd.ttf'))
    pdfmetrics.registerFont(TTFont('arial', 'arial.ttf'))
    pdfmetrics.registerFont(TTFont('arialbd', 'arial_bold.ttf'))

    c = canvas.Canvas(packet, letter)

    #Página 1

    #Header
    c.setFont('arialbd', 14)
    c.drawString(254, 579, str(nombre) + " " + str(apellidos))
    c.drawString(363, 561, str(dni))

    #Middle
    c.setFont('arial', 14)
    text = "de acuerdo con los requisitos establecidos en la Instrucción Técnica IT07 Revisión 4: “Competencia técnica de los instaladores en Baja Tensión”, conforme a los contenidos detallados en el Apéndice II de la ITCBT-03 del Reglamento Electrotécnico para Baja Tensión aprobado por el Real Decreto 842/2002, 18 de septiembre modificado por el Real Decreto 298/2021, de 27 de abril."
    
    c.drawString(140, 470, "Para la categoría ")
    
    c.setFont('arialbd', 13)
    if len(categoria) > 20:
        c.drawString(250, 470, str(categoria[0:44]))
        c.drawString(140, 453, str(categoria[44:len(categoria)].strip()))
    else:
        c.drawString(250, 470, str(categoria))
    c.setFont('arial', 13)

    if categoria == 'BÁSICA (IBTB)':
        c.drawString(342, 470, " de acuerdo con los requisitos estable-")
        c.drawString(140, 453, "cidos en la Instrucción Técnica IT07 Revisión 4: “Competencia técnica de")
        c.drawString(140, 436, "los instaladores en Baja Tensión”, conforme a los contenidos detallados")
        c.drawString(140, 419, "en el Apéndice II de la ITCBT-03 del Reglamento Electrotécnico para")
        c.drawString(140, 402, "Baja Tensión aprobado por el Real Decreto 842/2002, 18 de septiembre")
        c.drawString(140, 385, "modificado por el Real Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Sistemas de Automatización" (IBTE)':
        c.drawString(182, 453, "de acuerdo con los requisitos establecidos en la Instrucción Téc-")
        c.drawString(140, 436, "nica IT07 Revisión 4: “Competencia técnica de los instaladores en Baja")
        c.drawString(140, 419, "Tensión”, conforme a los contenidos detallados en el Apéndice II de la")
        c.drawString(140, 402, "ITCBT-03 del Reglamento Electrotécnico para BajaTensión aprobado")
        c.drawString(140, 385, "por el Real Decreto 842/2002, 18 de septiembre modificado por el Real")
        c.drawString(140, 368, "Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Líneas de Distribución BT"  (IBTE)':
        c.drawString(182, 453, "de acuerdo con los requisitos establecidos en la Instrucción Téc-")
        c.drawString(140, 436, "nica IT07 Revisión 4: “Competencia técnica de los instaladores en Baja")
        c.drawString(140, 419, "Tensión”, conforme a los contenidos detallados en el Apéndice II de la")
        c.drawString(140, 402, "ITCBT-03 del Reglamento Electrotécnico para BajaTensión aprobado")
        c.drawString(140, 385, "por el Real Decreto 842/2002, 18 de septiembre modificado por el Real")
        c.drawString(140, 368, "Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Instalaciones en locales con riesgo de incendio y explosión" (IBTE)':
        c.drawString(382, 453, "de acuerdo con los requisitos")
        c.drawString(140, 436, "establecidos en la Instrucción Técnica IT07 Revisión 4: “Competencia")
        c.drawString(140, 419, "técnica de los instaladores en Baja Tensión”, conforme a los contenidos")
        c.drawString(140, 402, "detallados en el Apéndice II de la ITCBT-03 del Reglamento Electrotéc-")
        c.drawString(140, 385, "nico para Baja Tensión aprobado por el Real Decreto 842/2002, 18 de")
        c.drawString(140, 368, "septiembre modificado por el Real Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Instalaciones en quirófanos y salas de intervención" (IBTE)':
        c.drawString(335, 453, "de acuerdo con los requisitos esta-")
        c.drawString(140, 436, "blecidos en la Instrucción Técnica IT07 Revisión 4: “Competencia téc-")
        c.drawString(140, 419, "nica de los instaladores en Baja Tensión”, conforme a los contenidos")
        c.drawString(140, 402, "detallados en el Apéndice II de la ITCBT-03 del Reglamento Electrotéc-")
        c.drawString(140, 385, "nico para Baja Tensión aprobado por el Real Decreto 842/2002, 18 de")
        c.drawString(140, 368, "septiembre modificado por el Real Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Instalaciones de lámparas de descarga en alta tensión y rótulos luminosos" (IBTE)':
        c.drawString(470, 453, "de acuerdo con")
        c.drawString(140, 436, "los requisitos establecidos en la Instrucción Técnica IT07 Revisión 4: “Com-")
        c.drawString(140, 419, "petencia técnica de los instaladores en Baja Tensión”, conforme a los con-")
        c.drawString(140, 402, "tenidos detallados en el Apéndice II de la ITCBT-03 del Reglamento Electro-")
        c.drawString(140, 385, "técnico para Baja Tensión aprobado por el Real Decreto 842/2002, 18 de")
        c.drawString(140, 368, "septiembre modificado por el Real Decreto 298/2021, de 27 de abril.")
    elif categoria == 'ESPECIALISTA - "Instalaciones generadoras de baja tensión de potencia superior o igual a 10 Kw" (IBTE)':
        c.drawString(498, 453, "de acuer-")
        c.drawString(140, 436, "do con los requisitos establecidos en la Instrucción Técnica IT07 Revi-")
        c.drawString(140, 419, "sión 4: “Competencia técnica de los instaladores en Baja Tensión”, con-")
        c.drawString(140, 402, "forme a los contenidos detallados en el Apéndice II de la ITCBT-03 del")
        c.drawString(140, 385, "Reglamento Electrotécnico para Baja Tensión aprobado por el Real De-")
        c.drawString(140, 368, "creto 842/2002, 18 de septiembre modificado por el Real Decreto")
        c.drawString(140, 351, "298/2021, de 27 de abril.")
    c.setFont('arialbd', 12)
    c.drawString(288, 306, str(fecha_vigor))

    #Footer
    c.setFont('arial', 11)
    c.drawString(50, 27, str(referencia))
    c.drawString(340, 27, str(certificado))
    c.drawString(508, 27, str(fecha_caducidad))
    c.setFont('arialbd', 9)
    c.drawString(500, 10, str(revision))

    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)
    
    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()
    
    #Creación página
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    new_pdf = os.path.join (files_folder, f"c{int(expediente)}.pdf")
    output_stream = open(new_pdf, "wb")
    output.write(output_stream)
    output_stream.close()

wb = xlrd.open_workbook(data) 

hoja = wb.sheet_by_index(0) 
for i in range (1, hoja.nrows):
    for j in range(10):      
        print(hoja.cell_value(i, j))
    nombre = hoja.cell_value(i, 0)
    apellidos = hoja.cell_value(i, 1)
    dni = hoja.cell_value(i, 2)
    categoria = hoja.cell_value(i, 3)
    fecha_vigor = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 4))
    fecha_vigor = str(fecha_vigor).split(" ")[0]
    fecha_vigor = fecha_vigor.split("-")[2] + "/" + fecha_vigor.split("-")[1] + "/" + fecha_vigor.split("-")[0].replace("20", "")
    referencia = hoja.cell_value(i, 5)
    certificado = hoja.cell_value(i, 6)
    fecha_caducidad = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 7))
    fecha_caducidad = str(fecha_caducidad).split(" ")[0]
    fecha_caducidad = fecha_caducidad.split("-")[2] + "/" + fecha_caducidad.split("-")[1] + "/" + fecha_caducidad.split("-")[0].replace("20", "")
    revision = hoja.cell_value(i, 8)
    expediente = hoja.cell_value(i, 9)
    print(fecha_vigor)
    print(fecha_caducidad)
    print("_______________________________")
    generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente)
print("Documentos generados correctamente")    
input()