import os
import io
import textwrap
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from datetime import datetime, timedelta
import textwrap

current_folder = os.path.dirname (__file__)
parent_folder = os.path.dirname (current_folder)
files_folder = os.path.join (parent_folder, "files")
data = os.path.join (files_folder, f"Data.xlsx")
original_pdf = os.path.join (current_folder, f"Certificado.pdf")
arial = os.path.join (current_folder, f"arial.ttf")
arial_bold = os.path.join (current_folder, f"arial_bold.ttf")


def generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('arial', arial))
    pdfmetrics.registerFont(TTFont('arialbd', arial_bold))

    c = canvas.Canvas(packet, letter)

    width, height = letter

    #Página 1

    text_width = c.stringWidth(nombre, 'arialbd', 14)
    x_position = (width - text_width) / 2
    #Header
    c.setFont('arialbd', 14)
    c.drawString(x_position, 590, str(nombre) + " " + str(apellidos))
    c.drawString(377, 572, str(dni))

    #Middle
    c.setFont('arial', 14)

    c.setFont('arialbd', 12)
    c.drawString(402, 334.5, str(fecha_vigor))

    #Footer
    c.setFont('arial', 11)
    c.drawString(67, 38, str(referencia))
    c.drawString(367, 38, str(certificado))
    c.drawString(532, 38, str(fecha_caducidad))

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
    try:    
        fecha_vigor = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 4))
        fecha_vigor = str(fecha_vigor).split(" ")[0]
        fecha_vigor = fecha_vigor.split("-")[2] + "/" + fecha_vigor.split("-")[1] + "/" + fecha_vigor.split("-")[0].replace("20", "")
    except:
        fecha_vigor = hoja.cell_value(i, 4)
    referencia = hoja.cell_value(i, 5)
    certificado = hoja.cell_value(i, 6)
    try:
        fecha_caducidad = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 7))
        fecha_caducidad = str(fecha_caducidad).split(" ")[0]
        fecha_caducidad = fecha_caducidad.split("-")[2] + "/" + fecha_caducidad.split("-")[1] + "/" + fecha_caducidad.split("-")[0].replace("20", "")
    except:
        fecha_caducidad = hoja.cell_value(i, 7)
    revision = hoja.cell_value(i, 8)
    expediente = hoja.cell_value(i, 9)
    print(fecha_vigor)
    print(fecha_caducidad)
    print("_______________________________")
    generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente)
print("Documentos generados correctamente")    
input()