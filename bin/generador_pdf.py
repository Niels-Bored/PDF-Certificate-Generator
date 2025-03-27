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


# Lista de frases que deben ir en negrita
bold_phrases = [
    "BÁSICA (IBTB)",
    "ESPECIALISTA - SISTEMAS DE AUTOMATIZACIÓN (IBTE)",
    "ESPECIALISTA - LÍNEAS DE DISTRIBUCIÓN  (IBTE)",
    "ESPECIALISTA - INSTALACIONES EN LOCALES CON RIESGO DE INCENDIO Y EXPLOSIÓN (IBTE)",
    "ESPECIALISTA - INSTALACIONES EN QUIRÓFANOS Y SALAS DE INTERVENCIÓN (IBTE)",
    "ESPECIALISTA - INSTALACIONES DE LÁMPARAS DE DESCARGA EN ALTA TENSIÓN Y RÓTULOS LUMINOSOS (IBTE)",
    "ESPECIALISTA - INSTALACIONES GENERADORAS DE BAJA TENSIÓN DE POTENCIA SUPERIOR O IGUAL A 10 KW (IBTE)"
]


# Función para justificar el texto y resaltar en negritas las frases clave
def justify_text(c, text, bold_phrases, x, y, width=440, font="Helvetica", font_bold="Helvetica-Bold", font_size=12):
    c.setFont(font, font_size)
    
    words = text.split(" ")  
    line = []
    line_width = 0
    space_width = c.stringWidth(" ", font, font_size)
    
    lines = []  # Almacena las líneas ya formadas
    word_positions = []  # Guarda la posición de cada palabra en la línea
    
    for word in words:
        word_width = c.stringWidth(word, font, font_size)
        
        if line_width + word_width <= width:
            line.append(word)
            line_width += word_width + space_width
        else:
            lines.append(line)
            line = [word]
            line_width = word_width + space_width
    
    if line:
        lines.append(line)
    
    for line in lines:
        draw_justified_line(c, line, x, y, width, font, font_bold, font_size, bold_phrases)
        y -= font_size + 4

# Función para imprimir una línea con justificación y negritas
def draw_justified_line(c, words, x, y, width, font, font_bold, font_size, bold_phrases):
    total_spaces = len(words) - 1
    text_width = sum(c.stringWidth(word, font, font_size) for word in words)
    
    if total_spaces > 0:
        extra_space = (width - text_width) / total_spaces
    else:
        extra_space = 0
    
    current_x = x
    for word in words:
        word_font = font_bold if any(word.replace(",", "").replace("\"","") in phrase and word.isupper() for phrase in bold_phrases) else font
        c.setFont(word_font, font_size)
        """ word_font = font_bold if word.isupper() else font
        c.setFont(word_font, font_size) """
        
        c.drawString(current_x, y, word)
        current_x += c.stringWidth(word, word_font, font_size) + extra_space

def justify_text2(text, max_width, c, x, y, font_name="Helvetica", font_size=12):
    """
    Dibuja texto justificado en un canvas de ReportLab.
    
    - text: Texto a imprimir
    - max_width: Ancho máximo de la línea en puntos
    - c: Objeto canvas de ReportLab
    - x, y: Coordenadas iniciales
    """
    words = text.split()  # Dividir en palabras
    lines = []
    current_line = []
    current_width = 0

    c.setFont(font_name, font_size)

    # Medir y dividir en líneas según el ancho permitido
    for word in words:
        word_width = c.stringWidth(word, font_name, font_size)
        space_width = c.stringWidth(" ", font_name, font_size)

        if current_width + word_width + (space_width if current_line else 0) > max_width:
            lines.append(current_line)
            current_line = [word]
            current_width = word_width
        else:
            current_line.append(word)
            current_width += word_width + (space_width if current_line else 0)

    if current_line:
        lines.append(current_line)

    # Dibujar cada línea justificada
    for line in lines:
        if len(line) == 1:  # Si solo hay una palabra, alinear a la izquierda
            c.drawString(x, y, line[0])
        else:
            total_word_width = sum(c.stringWidth(word, font_name, font_size) for word in line)
            total_spaces = len(line) - 1
            extra_space = (max_width - total_word_width) / total_spaces  # Espacio adicional entre palabras

            current_x = x
            for i, word in enumerate(line):
                c.drawString(current_x, y, word)
                current_x += c.stringWidth(word, font_name, font_size) + (extra_space if i < total_spaces else 0)

        y -= font_size + 4  # Ajustar altura para la siguiente línea


def generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente, text):
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


    
    #justify_text(text, max_width=440, c=c, x=152, y=470)
    justify_text(c, text, bold_phrases, font = "arial", font_bold="arialbd", x=152, y=470)

    
    c.setFont('arialbd', 12)
    c.drawString(402, 339.5, str(fecha_vigor))

    #Footer
    c.setFont('arial', 11)
    c.drawString(72, 38, str(referencia))
    c.drawString(370, 38, str(certificado))
    c.drawString(535, 38, str(fecha_caducidad))
    c.setFont('arialbd', 9)
    c.drawString(510, 23.5, str(revision))

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
    text = hoja.cell_value(i, 10)
    print(fecha_vigor)
    print(fecha_caducidad)
    print(text)
    print("_______________________________")
    generatePDF(nombre, apellidos, dni, categoria, fecha_vigor, referencia, certificado, fecha_caducidad, revision, expediente,text)
print("Documentos generados correctamente")    
input()