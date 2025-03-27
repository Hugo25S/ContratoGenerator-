import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from num2words import num2words
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def convert_date_to_spanish(date):
    months = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    day, month, year = map(int, date.split('-'))
    return f"{day} de {months[month - 1]} del {year}"

def convert_num_to_words(number):
    integer_part, decimal_part = f"{number:.2f}".split('.')
    integer_words = num2words(int(integer_part), lang='es')
    decimal_words = f"{decimal_part}/100"
    return f"{integer_words} y {decimal_words} soles"


def set_font(paragraph, font_name="Arial", font_size=9):
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        
        # Este bloque asegura que el cambio de fuente se aplique también en los documentos de formato Word XML.
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def autocomplete_word_template(template_path, output_path, excel_path):
    df = pd.read_excel(excel_path)
    
    for index, row in df.iterrows():
        doc = Document(template_path)

        for paragraph in doc.paragraphs:
            if "NOMBREp" in paragraph.text:
                paragraph.text = paragraph.text.replace("NOMBREp", row["NOMBRE_P"])
            if "DNIp" in paragraph.text:
                paragraph.text = paragraph.text.replace("DNIp", str(row["DNI_P"]))
            if "DIRECCIONp" in paragraph.text:
                paragraph.text = paragraph.text.replace("DIRECCIONp", row["DIRECCION_P"])
            if "CAUSA_OBJETIVAp" in paragraph.text:
                paragraph.text = paragraph.text.replace("CAUSA_OBJETIVAp", row["CAUSA OBJETIVA_P"])
                # Establecer la alineación del párrafo a justificado
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if "AREAp" in paragraph.text:
                paragraph.text = paragraph.text.replace("AREAp", row["AREA_P"])
            if "CONDICIONp" in paragraph.text:
                paragraph.text = paragraph.text.replace("CONDICIONp", row["CONDICION_P"])
            if "CARGOp" in paragraph.text:
                paragraph.text = paragraph.text.replace("CARGOp", row["CARGO_P"])
            if "SEDEp" in paragraph.text:
                paragraph.text = paragraph.text.replace("SEDEp", row["SEDE_P"])
            if "FECHA_INICIOp" in paragraph.text:
                paragraph.text = paragraph.text.replace("FECHA_INICIOp", row["FECHA_INICIO_P"].strftime('%d-%m-%Y'))
            if "FECHA_FINp" in paragraph.text:
                paragraph.text = paragraph.text.replace("FECHA_FINp", row["FECHA_FIN_P"].strftime('%d-%m-%Y'))
            if "FECHA_FIN_LETRAp" in paragraph.text:
                fecha_fin_letra = convert_date_to_spanish(row["FECHA_FIN_P"].strftime('%d-%m-%Y'))
                paragraph.text = paragraph.text.replace("FECHA_FIN_LETRAp", fecha_fin_letra)
            if "REMUNERACIONp" in paragraph.text:
                paragraph.text = paragraph.text.replace("REMUNERACIONp", f"S/{row['REMUNERACION_P']:.2f}")
            if "NUMERO_LETRAp" in paragraph.text:
                numero_letra = convert_num_to_words(row["REMUNERACION_P"])
                paragraph.text = paragraph.text.replace("NUMERO_LETRAp", numero_letra)
            if "PERIODO_PRUEBAp" in paragraph.text:
                paragraph.text = paragraph.text.replace("PERIODO_PRUEBAp", str(row["PERIODO_PRUEBA_P"]))
            if "FONOp" in paragraph.text:
                paragraph.text = paragraph.text.replace("FONOp", str(row["FONO_P"]))
            if "CORREO_P" in paragraph.text:
                paragraph.text = paragraph.text.replace("CORREO_P", row["CORREO_P"])
            
            # Aplicar formato a cada párrafo
            set_font(paragraph, "Arial", 9)


        output_file = f"{output_path}/Incremento_actividades_{row['NOMBRE_P']}.docx"
        doc.save(output_file)

# Ejemplo de uso
template_path = "Incremento_actividades.docx"
output_path = "output"
excel_path = "datos.xlsx"

autocomplete_word_template(template_path, output_path, excel_path)
