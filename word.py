from io import BytesIO
from collections import OrderedDict
from typing import List, Optional, Tuple

# Librerías para manejo de documentos Word (python-docx)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL


def crear_documento_a4() -> Document:
    """
    Crea un documento Word con tamaño de página A4 y márgenes de 1 pulgada.
    """
    doc = Document()
    section = doc.sections[0]
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    return doc


def guardar_documento(doc: Document, ruta_salida: str) -> None:
    doc.save(ruta_salida)


def agregar_titulo(doc: Document, texto: str, nivel: int) -> None:
    # Paleta de colores corporativos sobrios
    COLOR_TITULO = RGBColor(0x2E, 0x3F, 0x5F)  # Azul marino oscuro
    COLOR_SUBTITULO = RGBColor(0x4F, 0x4F, 0x4F)  # Gris oscuro

    if nivel == 1:
        # Título principal - Nivel 1
        titulo = doc.add_heading(level=1)
        run = titulo.add_run(texto.upper())
        run.font.name = 'Lora'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = COLOR_TITULO
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        titulo.paragraph_format.space_before = Pt(18)
        titulo.paragraph_format.space_after = Pt(12)

        # Agregar línea decorativa inferior
        p = titulo._element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '8')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '2E3F5F')
        pBdr.append(bottom)

    elif nivel == 2:
        # Subtítulo importante - Nivel 2
        titulo = doc.add_heading(level=2)
        run = titulo.add_run(texto)
        run.font.name = 'Lora'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_SUBTITULO
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        titulo.paragraph_format.space_before = Pt(14)
        titulo.paragraph_format.space_after = Pt(8)

        # Subrayado decorativo
        p = titulo._element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'D3D3D3')
        pBdr.append(bottom)

    elif nivel == 3:
        # Subtítulo secundario - Nivel 3
        titulo = doc.add_heading(level=3)
        run = titulo.add_run(texto)
        run.font.name = 'Lora'
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_SUBTITULO
        run.font.italic = True
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        titulo.paragraph_format.space_before = Pt(10)
        titulo.paragraph_format.space_after = Pt(4)

    else:
        # Para niveles inferiores
        parrafo = doc.add_paragraph(texto)
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        run = parrafo.runs[0]
        run.font.name = 'Segoe UI Light'
        run.font.size = Pt(8)
        run.font.underline = True
        run.font.bold = True


def agregar_parrafo(doc: Document, texto: str) -> None:
    parrafo = doc.add_paragraph(texto)
    parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    run = parrafo.runs[0]
    run.font.name = 'Segoe UI Light'
    run.font.size = Pt(8)


def insertar_figura(doc: Document, figura, titulo: Optional[str] = None, pie: Optional[str] = None) -> None:
    if titulo:
        agregar_titulo(doc, titulo, 3)
    imagen_stream = BytesIO()
    figura.savefig(imagen_stream, format='png', bbox_inches='tight')
    imagen_stream.seek(0)

    p = doc.add_paragraph()
    run = p.add_run()
    run.add_picture(imagen_stream, width=Inches(5.5))
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    imagen_stream.close()

    if pie:
        pie_p = doc.add_paragraph(pie)
        pie_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = pie_p.runs[0]
        run.font.name = 'Segoe UI Light'
        run.font.size = Pt(6)
        run.font.bold = True
        run.font.italic = True


def set_cell_width(cell, width_inches: float) -> None:
    """
    Establece el ancho de una celda en pulgadas.
    """
    width_twips = int(width_inches * 1440)
    cell.width = Inches(width_inches)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Eliminar cualquier w:tcW anterior
    for child in tcPr.findall(qn('w:tcW')):
        tcPr.remove(child)

    # Crear nuevo elemento de ancho
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_twips))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


def insertar_tabla(doc: Document, df, titulo: Optional[str] = None):
    if titulo:
        agregar_titulo(doc, titulo, 3)

    tabla = doc.add_table(rows=1, cols=len(df.columns))
    tabla.style = 'Table Grid'
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

    ancho_total = 6.0
    ancho_columna = ancho_total / len(df.columns)

    # Encabezados
    hdr_cells = tabla.rows[0].cells
    for i, col_name in enumerate(df.columns):
        cell = hdr_cells[i]
        cell.text = str(col_name)
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(6.5)
        run.font.name = 'Segoe UI Light'
        set_cell_width(cell, ancho_columna)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Filas de datos
    for _, row in df.iterrows():
        row_cells = tabla.add_row().cells
        for i, value in enumerate(row):
            cell = row_cells[i]
            cell.text = str(value)
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(7)
            set_cell_width(cell, ancho_columna)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    return tabla


def insertar_tabla_con_merge(doc: Document, df, titulo: Optional[str] = None, group_cols: Optional[List[str]] = None):
    if titulo:
        agregar_titulo(doc, titulo, 3)

    tabla = doc.add_table(rows=1, cols=len(df.columns))
    tabla.style = 'Table Grid'
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

    ancho_total = 6.0
    ancho_columna = ancho_total / len(df.columns)

    # Encabezados
    hdr_cells = tabla.rows[0].cells
    for i, col in enumerate(df.columns):
        cell = hdr_cells[i]
        cell.text = str(col)
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(6.5)
        run.font.name = 'Segoe UI Light'
        set_cell_width(cell, ancho_columna)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Filas de datos
    for _, row in df.iterrows():
        row_cells = tabla.add_row().cells
        for i, val in enumerate(row):
            cell = row_cells[i]
            cell.text = str(val)
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(7)
            set_cell_width(cell, ancho_columna)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if group_cols:
        col2idx = {col: idx for idx, col in enumerate(df.columns)}
        sizes = OrderedDict()
        prev_key = None
        for key_vals in df[group_cols].itertuples(index=False, name=None):
            if key_vals == prev_key:
                sizes[key_vals] += 1
            else:
                sizes[key_vals] = 1
                prev_key = key_vals

        current_row = 1
        for key_vals, size in sizes.items():
            if size > 1:
                for col in group_cols:
                    c_idx = col2idx[col]
                    start = tabla.cell(current_row, c_idx)
                    end = tabla.cell(current_row + size - 1, c_idx)
                    for r in range(current_row + 1, current_row + size):
                        tabla.cell(r, c_idx).text = ''
                    start.merge(end)
                    start.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    start.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            current_row += size

    return tabla


def insertar_salto_pagina(doc: Document) -> None:
    doc.add_page_break()


def agregar_viñetas(
    doc: Document,
    items: List[str],
    nivel: int = 1,
    espacio_antes: Pt = Pt(4),
    espacio_despues: Pt = Pt(4),
) -> None:
    """
    Inserta una lista usando guiones '-' como viñetas.
    """
    indent_por_nivel = Pt(12)

    for texto in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = espacio_antes
        p.paragraph_format.space_after = espacio_despues
        p.paragraph_format.left_indent = indent_por_nivel * (nivel - 1)

        run = p.add_run(f"- {texto}")
        run.font.name = 'Segoe UI Light'
        run.font.size = Pt(8)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def insertar_en_posicion(
    doc: Document,
    funcion_contenido,
    *args,
    posicion: str = 'final',
    **kwargs,
) -> None:
    """
    Inserta contenido generado por una función en una posición específica del documento.
    posicion: 'inicio', 'final' o 'index:<n>'
    """
    doc_temp = Document()
    funcion_contenido(doc_temp, *args, **kwargs)

    elementos_temp = list(doc_temp.element.body)
    body = doc.element.body

    if posicion == 'inicio':
        for elem in reversed(elementos_temp):
            body.insert(0, elem)
    elif posicion == 'final':
        for elem in elementos_temp:
            body.append(elem)
    elif posicion.startswith('index:'):
        idx = int(posicion.split(':')[1])
        for i, elem in enumerate(elementos_temp):
            body.insert(idx + i, elem)
    else:
        raise ValueError("La posición debe ser 'inicio', 'final' o 'index:<n>'")


def insertar_indice(doc: Document, titulo: str = "Índice") -> None:
    agregar_titulo(doc, titulo, 1)
    p = doc.add_paragraph()
    run = p.add_run()

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = r'TOC \o "1-3" \h \z \u'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_after = Pt(6)


def agregar_advertencia_actualizacion(doc: Document) -> None:
    p = doc.add_paragraph()
    run = p.add_run("⚠️ Al abrir este documento, recuerde actualizar los campos (índice, referencias cruzadas, etc.).")
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x00, 0x00)
    p.paragraph_format.space_before = Pt(12)


def mostrar_contenido(doc: Document) -> None:
    print("Índice | Tipo   | Contenido resumido")
    print("--------------------------------------")

    idx_parrafo = 0
    idx_tabla = 0

    for i, elem in enumerate(doc.element.body):
        tag = elem.tag.split('}')[-1]

        if tag == 'p':
            parrafo = doc.paragraphs[idx_parrafo]
            texto = parrafo.text.strip().replace('\n', ' ')
            print(f"{i:<6} | Párrafo | '{texto[:60]}'")
            idx_parrafo += 1

        elif tag == 'tbl':
            print(f"{i:<6} | Tabla   | [Tabla con {len(doc.tables[idx_tabla].rows)} filas]")
            idx_tabla += 1

        else:
            print(f"{i:<6} | Otro    | Etiqueta: {tag}")


def mostrar_contenido_posicional(doc: Document, buscar: Optional[str] = None) -> List[int]:
    idx_parrafo = 0
    posiciones_encontradas: List[int] = []

    for i, elem in enumerate(doc.element.body):
        tag = elem.tag.split('}')[-1]

        if tag == 'p':
            parrafo = doc.paragraphs[idx_parrafo]
            texto = parrafo.text.strip().replace('\n', ' ')
            if buscar and buscar.lower() in texto.lower():
                posiciones_encontradas.append(i)
            idx_parrafo += 1

    return posiciones_encontradas


def reemplazar_parrafo(original: Paragraph, nuevo: Paragraph) -> None:
    original._element.getparent().replace(original._element, nuevo._element)


def numerar_titulos_existentes(doc: Document) -> None:
    contador = {1: 0, 2: 0, 3: 0}
    reemplazos: List[Tuple[Paragraph, Paragraph]] = []

    for _, parrafo in enumerate(doc.paragraphs):
        estilo = parrafo.style.name.strip()
        if estilo.startswith("Heading"):
            try:
                nivel = int(estilo.split()[-1])
            except (ValueError, IndexError):
                continue

            if nivel in contador:
                contador[nivel] += 1
                for deeper in range(nivel + 1, 4):
                    contador[deeper] = 0

                if nivel == 1:
                    numeracion = f"{contador[1]}."
                elif nivel == 2:
                    numeracion = f"{contador[1]}.{contador[2]}"
                else:
                    numeracion = f"{contador[1]}.{contador[2]}.{contador[3]}"

                texto = parrafo.text.strip()
                if not texto.startswith(numeracion):
                    doc_temp = Document()
                    agregar_titulo(doc_temp, f"{numeracion} {texto}", nivel)
                    nuevo_parrafo = doc_temp.paragraphs[0]
                    reemplazos.append((parrafo, nuevo_parrafo))

    for original, nuevo in reemplazos:
        reemplazar_parrafo(original, nuevo)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generador de documento Word (A4) con utilidades de formato.")
    parser.add_argument("salida", help="Ruta del archivo .docx de salida, p.ej.: informe.docx")
    args = parser.parse_args()

    doc = crear_documento_a4()
    insertar_indice(doc)
    agregar_advertencia_actualizacion(doc)
    agregar_titulo(doc, "Resumen de Proyecto", 1)
    agregar_parrafo(doc, "Este es un ejemplo mínimo de documento generado desde funcion.py.")
    agregar_titulo(doc, "Sección", 2)
    agregar_viñetas(doc, ["Ítem 1", "Ítem 2", "Ítem 3"]) 

    guardar_documento(doc, args.salida)
    print(f"Documento guardado en: {args.salida}")


