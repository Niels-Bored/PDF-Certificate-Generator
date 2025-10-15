import os
import io
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime, timedelta

current_folder = os.path.dirname(__file__)
parent_folder = os.path.dirname(current_folder)
files_folder = os.path.join(parent_folder, "files")
data = os.path.join(files_folder, "Data.xlsx")
original_pdf = os.path.join(current_folder, "Certificado.pdf")
arial = os.path.join(current_folder, "arial.ttf")
arial_bold = os.path.join(current_folder, "arial_bold.ttf")


# Lista de frases que deben ir en negrita
bold_phrases = [
    "BÁSICA (IBTB)",
    "ESPECIALISTA - SISTEMAS DE AUTOMATIZACIÓN (IBTE)",
    "ESPECIALISTA - LÍNEAS DE DISTRIBUCIÓN  (IBTE)",
    "ESPECIALISTA - INSTALACIONES EN LOCALES CON RIESGO DE INCENDIO Y EXPLOSIÓN (IBTE)",
    "ESPECIALISTA - INSTALACIONES EN QUIRÓFANOS Y SALAS DE INTERVENCIÓN (IBTE)",
    "ESPECIALISTA - INSTALACIONES DE LÁMPARAS DE DESCARGA EN ALTA TENSIÓN Y RÓTULOS LUMINOSOS (IBTE)",
    "ESPECIALISTA - INSTALACIONES GENERADORAS DE BAJA TENSIÓN DE POTENCIA SUPERIOR O IGUAL A 10 KW (IBTE)",
]


def justify_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    bold_phrases: str,
    width: float = 440,
    font: str = "arial",
    font_bold="arialbd",
    font_size: float = 12,
):
    """Justify text on PDF file

    Args:
        c (canvas.Canvas): PDF Canvas representation
        text (str): text to justify
        x (float): x coordinate
        y (float): y coordinate
        width (float): PDF page width
        font (str): font name
        font_size (int): font size
    """
    c.setFont(font, font_size)

    words = text.split(" ")
    line = []
    line_width = 0
    space_width = c.stringWidth(" ", font, font_size)

    lines = []  # Store formated lines

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

    for i, line in enumerate(lines):
        final = i == len(lines) - 1
        draw_justified_line(
            c, line, x, y, width, font, font_bold, font_size, bold_phrases, final
        )
        y -= font_size + 4


def draw_justified_line(
    c, words, x, y, width, font, font_bold, font_size, bold_phrases, final
):
    """Draw a line with justification

    Args:
        c (canvas.Canvas): PDF Canvas representation
        words (list): list of words
        x (float): x coordinate
        y (float): y coordinate
        width (float): PDF page width
        font (str): font name
        font_size (int): font size
        final (bool): if it's the last line
    """
    total_spaces = len(words) - 1
    text_width = sum(c.stringWidth(word, font, font_size) for word in words)

    if total_spaces > 0:
        extra_space = (width - text_width) / total_spaces
    else:
        extra_space = 0

    if final:
        extra_space = 4

    current_x = x
    for word in words:
        word_font = (
            font_bold
            if any(
                word.replace(",", "").replace('"', "") in phrase and word.isupper()
                for phrase in bold_phrases
            )
            else font
        )
        c.setFont(word_font, font_size)
        c.drawString(current_x, y, word)
        current_x += c.stringWidth(word, font, font_size) + extra_space


def generatePDF(
    nombre,
    apellidos,
    dni,
    categoria,
    fecha_vigor,
    referencia,
    certificado,
    fecha_caducidad,
    revision,
    expediente,
    text,
    footer
):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont("arial", arial))
    pdfmetrics.registerFont(TTFont("arialbd", arial_bold))

    c = canvas.Canvas(packet, letter)

    width, height = letter

    # Página 1

    text_width = c.stringWidth(nombre + " " + apellidos, "arialbd", 14)
    x_position = (width - text_width) / 2

    print(f"width: {width}")
    print(f"text width: {text_width}")
    print(f"x position: {x_position}")
    x_position += 70
    print(f"x position adjustment: {x_position}")
    # Header
    c.setFont("arialbd", 14)
    c.drawString(x_position, 590, str(nombre) + " " + str(apellidos))
    c.drawString(377, 572, str(dni))

    # Middle
    c.setFont("arial", 14)

    justify_text(
        c=c,
        text=text,
        x=152,
        y=470,
        bold_phrases=bold_phrases,
        font="arial",
        font_bold="arialbd",
    )

    c.setFont("arialbd", 12)
    c.drawString(402, 261.5, str(fecha_vigor))

    # Footer
    justify_text(
        c=c,
        text=footer,
        x=152,
        y=100,
        bold_phrases=bold_phrases,
        font="arial",
        font_bold="arialbd",
        font_size=11
    )

    c.setFont("arial", 11)
    c.drawString(72, 38, str(referencia))
    c.drawString(370, 38, str(certificado))
    c.drawString(535, 38, str(fecha_caducidad))
    c.setFont("arialbd", 9)
    c.drawString(510, 23.5, str(revision))

    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)

    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()

    # Creación página
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    new_pdf = os.path.join(files_folder, f"c{int(expediente)}.pdf")
    output_stream = open(new_pdf, "wb")
    output.write(output_stream)
    output_stream.close()


wb = xlrd.open_workbook(data)

hoja = wb.sheet_by_index(0)
for i in range(1, hoja.nrows):
    for j in range(10):
        print(hoja.cell_value(i, j))
    nombre = hoja.cell_value(i, 0)
    apellidos = hoja.cell_value(i, 1)
    dni = hoja.cell_value(i, 2)
    categoria = hoja.cell_value(i, 3)
    try:
        fecha_vigor = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 4))
        fecha_vigor = str(fecha_vigor).split(" ")[0]
        fecha_vigor = (
            fecha_vigor.split("-")[2]
            + "/"
            + fecha_vigor.split("-")[1]
            + "/"
            + fecha_vigor.split("-")[0].replace("20", "")
        )
    except:
        fecha_vigor = hoja.cell_value(i, 4)
    referencia = hoja.cell_value(i, 5)
    certificado = hoja.cell_value(i, 6)
    try:
        fecha_caducidad = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 7))
        fecha_caducidad = str(fecha_caducidad).split(" ")[0]
        fecha_caducidad = (
            fecha_caducidad.split("-")[2]
            + "/"
            + fecha_caducidad.split("-")[1]
            + "/"
            + fecha_caducidad.split("-")[0].replace("20", "")
        )
    except:
        fecha_caducidad = hoja.cell_value(i, 7)
    revision = hoja.cell_value(i, 8)
    expediente = hoja.cell_value(i, 9)
    text = hoja.cell_value(i, 10)
    footer = hoja.cell_value(i, 11)
    print(fecha_vigor)
    print(fecha_caducidad)
    print(text)
    print(footer)
    print("_______________________________")
    generatePDF(
        nombre,
        apellidos,
        dni,
        categoria,
        fecha_vigor,
        referencia,
        certificado,
        fecha_caducidad,
        revision,
        expediente,
        text,
        footer
    )
print("Documentos generados correctamente")
input()
