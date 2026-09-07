from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[2]
PROJECT = ROOT / "sistema-reposicion-mini-market"
PROPOSAL = ROOT / "Propuesta"
REFERENCE = PROPOSAL / "TFE_Tercera_Entrega_Orlando_Delgado_Pinzon_revisado.docx"
OUTPUT = PROPOSAL / "TFE_Final_Jurado_Orlando_Delgado_Pinzon.docx"
FIGURES = PROJECT / "reports" / "figures"
ARCH = PROPOSAL / "figuras_generadas" / "figura_arquitectura.png"
PIPELINE = PROPOSAL / "figuras_generadas" / "figura_pipeline.png"


# These page maps are filled after rendering the final document. Keeping them
# explicit makes the delivered indexes independent from Word field updates.
INDEX_PAGE_MAP: dict[str, int] = {
    "Introducción": 8,
    "Motivación": 8,
    "Planteamiento del trabajo": 9,
    "Estructura del trabajo": 10,
    "Contexto y estado del arte": 11,
    "Contexto del problema": 11,
    "Relevancia de la predicción de demanda en retail": 12,
    "Estado del arte": 12,
    "Marco teórico": 13,
    "Series temporales y predicción de demanda": 13,
    "Aprendizaje automático supervisado aplicado al forecasting": 14,
    "Métricas de error y validación temporal": 14,
    "Reposición de inventario": 15,
    "Automatización con n8n y concepto de agente inteligente": 15,
    "Interfaz, despliegue local y reproducibilidad": 15,
    "Soluciones existentes y brecha identificada": 16,
    "Conclusiones del estado del arte": 16,
    "Objetivos concretos y metodología de trabajo": 17,
    "Objetivo general": 17,
    "Objetivos específicos y resultados verificables": 17,
    "Diseño metodológico": 18,
    "Protocolo experimental y prevención de fuga temporal": 19,
    "Reproducibilidad y gestión de evidencias": 20,
    "Definición de requisitos": 21,
    "Alcance y contexto de uso": 21,
    "Priorización del alcance mínimo": 21,
    "Actores del sistema": 22,
    "Casos de uso": 22,
    "Requisitos funcionales": 24,
    "Requisitos no funcionales": 26,
    "Requisitos de datos y criterios de aceptación": 28,
    "Trazabilidad y validación de requisitos": 28,
    "Diseño y desarrollo de la solución": 30,
    "Arquitectura implementada": 30,
    "Conjunto de datos utilizado": 31,
    "Perfil de productos y control de calidad": 32,
    "Validación, limpieza y transformación": 33,
    "Modelos implementados y configuración real": 35,
    "Selección y serialización del modelo": 36,
    "Motor explicable de reposición": 36,
    "Interfaz Streamlit": 37,
    "API y persistencia local": 38,
    "Automatización con n8n y despliegue con Docker": 39,
    "Pruebas automatizadas y manejo de errores": 40,
    "Estructura del repositorio y responsabilidades": 41,
    "Secuencia operativa y trazabilidad de una ejecución": 42,
    "Evaluación de la solución": 44,
    "Diseño de la evaluación": 44,
    "Métricas de error": 45,
    "Evaluación auxiliar de la ventana final": 45,
    "Resultados por horizonte": 46,
    "Selección del modelo final": 47,
    "Resultados por producto": 48,
    "Análisis del pronóstico operativo": 50,
    "Resultados de reposición": 51,
    "Análisis de sensibilidad de inventario": 52,
    "Verificación funcional y reproducibilidad": 53,
    "Matriz de verificación de requisitos": 54,
    "Discusión de validez": 55,
    "Comparación con el estado del arte": 56,
    "Interpretación operativa de los errores": 56,
    "Conclusiones y trabajo futuro": 59,
    "Conclusiones generales": 59,
    "Cumplimiento de los objetivos": 59,
    "Contribuciones del trabajo": 60,
    "Limitaciones": 61,
    "Trabajo futuro": 61,
    "Reflexión final": 62,
    "Referencias bibliográficas": 63,
    "Anexo A. Código fuente y datos analizados": 65,
    "Contenido del repositorio": 65,
    "Comandos de reproducción": 65,
    "Endpoints de la API": 66,
    "Diccionario de artefactos de evaluación": 66,
    "Protocolo pendiente de validación con usuarios": 67,
    "Lista de evidencias antes de entrega": 68,
}
FIGURE_PAGE_MAP: dict[str, int] = {
    "Figura 1. Arquitectura implementada del sistema de predicción y reposición": 30,
    "Figura 2. Volumen mensual de ventas sintéticas por categoría": 32,
    "Figura 3. Pipeline reproducible desde la ingestión hasta la recomendación": 35,
    "Figura 4. Documentación interactiva de la API FastAPI en ejecución": 39,
    "Figura 5. Ejecución satisfactoria de la suite automatizada": 41,
    "Figura 6. Comparación de MAE y SMAPE por modelo y horizonte": 47,
    "Figura 7. SMAPE de Random Forest por producto y horizonte": 50,
    "Figura 8. Pronóstico operativo de catorce días para un producto de ejemplo": 51,
    "Figura 9. Cantidad de reposición sugerida y prioridad por producto": 52,
}
TABLE_PAGE_MAP: dict[str, int] = {
    "Tabla 1. Trazabilidad de objetivos y resultados": 17,
    "Tabla 2. Actores identificados para el prototipo": 22,
    "Tabla 3. Casos de uso principales": 22,
    "Tabla 4. Requisitos funcionales": 24,
    "Tabla 5. Requisitos no funcionales": 26,
    "Tabla 6. Criterios de aceptación de datos y trazabilidad": 28,
    "Tabla 7. Características de los archivos de entrada": 31,
    "Tabla 8. Productos incluidos en la experimentación": 32,
    "Tabla 9. Estado inicial y parámetros fijos de inventario": 32,
    "Tabla 10. Variables finales del modelo": 34,
    "Tabla 11. Configuración ejecutada de los modelos": 36,
    "Tabla 12. Cálculo de la recomendación de reposición": 37,
    "Tabla 13. Servicios comprobados mediante Docker Compose": 40,
    "Tabla 14. Módulos principales y responsabilidad": 41,
    "Tabla 15. Comprobación de cardinalidad por horizonte y modelo": 44,
    "Tabla 16. Métricas auxiliares del endpoint de entrenamiento": 45,
    "Tabla 17. Resultados predictivos por horizonte": 46,
    "Tabla 18. Ranking promedio multihorizonte": 47,
    "Tabla 19. Rendimiento medio de Random Forest por producto": 48,
    "Tabla 20. Recomendaciones generadas para el horizonte operativo": 51,
    "Tabla 21. Sensibilidad a días de seguridad y revisión": 52,
    "Tabla 22. Resumen de pruebas automatizadas": 53,
    "Tabla 23. Matriz de verificación con evidencias": 54,
    "Tabla 24. Contrato resumido de la API": 66,
    "Tabla 25. Archivos de resultados reproducibles": 66,
    "Tabla 26. Instrumento SUS propuesto": 67,
}

# Verified against the rendered 63-page jury version after the final content
# and table-layout pass.
INDEX_PAGE_MAP.update({
    "Introducción": 8,
    "Motivación": 8,
    "Planteamiento del trabajo": 9,
    "Estructura del trabajo": 10,
    "Contexto y estado del arte": 11,
    "Contexto del problema": 11,
    "Relevancia de la predicción de demanda en retail": 12,
    "Estado del arte": 12,
    "Marco teórico": 13,
    "Series temporales y predicción de demanda": 14,
    "Aprendizaje automático supervisado aplicado al forecasting": 14,
    "Métricas de error y validación temporal": 14,
    "Reposición de inventario": 15,
    "Automatización con n8n y concepto de agente inteligente": 15,
    "Interfaz, despliegue local y reproducibilidad": 16,
    "Soluciones existentes y brecha identificada": 16,
    "Conclusiones del estado del arte": 16,
    "Objetivos concretos y metodología de trabajo": 18,
    "Objetivo general": 18,
    "Objetivos específicos y resultados verificables": 18,
    "Diseño metodológico": 18,
    "Protocolo experimental y prevención de fuga temporal": 20,
    "Reproducibilidad y gestión de evidencias": 21,
    "Definición de requisitos": 22,
    "Alcance y contexto de uso": 22,
    "Priorización del alcance mínimo": 22,
    "Actores del sistema": 23,
    "Casos de uso": 23,
    "Requisitos funcionales": 24,
    "Requisitos no funcionales": 25,
    "Requisitos de datos y criterios de aceptación": 25,
    "Trazabilidad y validación de requisitos": 26,
    "Diseño y desarrollo de la solución": 27,
    "Arquitectura implementada": 27,
    "Conjunto de datos utilizado": 28,
    "Perfil de productos y control de calidad": 29,
    "Validación, limpieza y transformación": 30,
    "Modelos implementados y configuración real": 31,
    "Selección y serialización del modelo": 32,
    "Motor explicable de reposición": 32,
    "Interfaz Streamlit": 33,
    "API y persistencia local": 34,
    "Automatización con n8n y despliegue con Docker": 35,
    "Pruebas automatizadas y manejo de errores": 36,
    "Estructura del repositorio y responsabilidades": 37,
    "Contratos de ejecución y control de fallos": 38,
    "Secuencia operativa y trazabilidad de una ejecución": 39,
    "Evaluación de la solución": 41,
    "Diseño de la evaluación": 41,
    "Métricas de error": 42,
    "Evaluación auxiliar de la ventana final": 42,
    "Resultados por horizonte": 43,
    "Selección del modelo final": 44,
    "Resultados por producto": 45,
    "Resultados por categoría y sesgo": 46,
    "Análisis del pronóstico operativo": 47,
    "Resultados de reposición": 48,
    "Análisis de sensibilidad de inventario": 49,
    "Verificación funcional y reproducibilidad": 50,
    "Matriz de verificación de requisitos": 51,
    "Discusión de validez": 52,
    "Comparación con el estado del arte": 52,
    "Interpretación operativa de los errores": 53,
    "Conclusiones y trabajo futuro": 55,
    "Conclusiones generales": 55,
    "Cumplimiento de los objetivos": 55,
    "Contribuciones del trabajo": 56,
    "Limitaciones": 57,
    "Trabajo futuro": 57,
    "Reflexión final": 58,
    "Referencias bibliográficas": 59,
    "Anexo A. Código fuente y datos analizados": 61,
    "Contenido del repositorio": 61,
    "Comandos de reproducción": 61,
    "Endpoints de la API": 62,
    "Diccionario de artefactos de evaluación": 62,
    "Protocolo pendiente de validación con usuarios": 62,
    "Lista de evidencias antes de entrega": 63,
})
FIGURE_PAGE_MAP.clear()
FIGURE_PAGE_MAP.update({
    "Figura 1. Arquitectura implementada del sistema de predicción y reposición": 27,
    "Figura 2. Volumen mensual de ventas sintéticas por categoría": 29,
    "Figura 3. Pipeline reproducible desde la ingestión hasta la recomendación": 31,
    "Figura 4. Documentación interactiva de la API FastAPI en ejecución": 35,
    "Figura 5. Ejecución satisfactoria de la suite automatizada": 37,
    "Figura 6. Comparación de MAE y SMAPE por modelo y horizonte": 44,
    "Figura 7. SMAPE de Random Forest por producto y horizonte": 46,
    "Figura 8. Pronóstico operativo de catorce días para un producto de ejemplo": 48,
    "Figura 9. Cantidad de reposición sugerida y prioridad por producto": 49,
})
TABLE_PAGE_MAP.clear()
TABLE_PAGE_MAP.update({
    "Tabla 1. Trazabilidad de objetivos y resultados": 18,
    "Tabla 2. Actores identificados para el prototipo": 23,
    "Tabla 3. Casos de uso principales": 23,
    "Tabla 4. Requisitos funcionales": 24,
    "Tabla 5. Requisitos no funcionales": 25,
    "Tabla 6. Criterios de aceptación de datos y trazabilidad": 25,
    "Tabla 7. Características de los archivos de entrada": 28,
    "Tabla 8. Productos incluidos en la experimentación": 29,
    "Tabla 9. Estado inicial y parámetros fijos de inventario": 29,
    "Tabla 10. Variables finales del modelo": 31,
    "Tabla 11. Configuración ejecutada de los modelos": 32,
    "Tabla 12. Cálculo de la recomendación de reposición": 33,
    "Tabla 13. Servicios comprobados mediante Docker Compose": 36,
    "Tabla 14. Módulos principales y responsabilidad": 37,
    "Tabla 15. Precondiciones, salidas y fallos controlados": 38,
    "Tabla 16. Comprobación de cardinalidad por horizonte y modelo": 41,
    "Tabla 17. Métricas auxiliares del endpoint de entrenamiento": 42,
    "Tabla 18. Resultados predictivos por horizonte": 43,
    "Tabla 19. Ranking promedio multihorizonte": 44,
    "Tabla 20. Rendimiento medio de Random Forest por producto": 45,
    "Tabla 21. Rendimiento de Random Forest por categoría": 46,
    "Tabla 22. Recomendaciones generadas para el horizonte operativo": 48,
    "Tabla 23. Sensibilidad a días de seguridad y revisión": 49,
    "Tabla 24. Resumen de pruebas automatizadas": 50,
    "Tabla 25. Tiempos de ejecución en el entorno local": 50,
    "Tabla 26. Matriz de verificación con evidencias": 51,
    "Tabla 27. Contrato resumido de la API": 62,
    "Tabla 28. Archivos de resultados reproducibles": 62,
    "Tabla 29. Instrumento SUS propuesto": 63,
})


def paragraph_from_element(doc: Document, element):
    from docx.text.paragraph import Paragraph

    return Paragraph(element, doc._body)


def find_heading(doc: Document, text: str, style: str = "Heading 1"):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == style and text.lower() in paragraph.text.lower():
            return paragraph
    raise ValueError(f"Heading not found: {text}")


def clear_section(doc: Document, heading_text: str):
    heading = find_heading(doc, heading_text)
    current = heading._p.getnext()
    while current is not None:
        if current.tag == qn("w:sectPr"):
            break
        if current.tag == qn("w:p"):
            paragraph = paragraph_from_element(doc, current)
            if paragraph.style.name == "Heading 1" and paragraph.text.strip():
                break
        nxt = current.getnext()
        current.getparent().remove(current)
        current = nxt
    return heading._p


def insert_paragraph(doc: Document, after, text: str = "", style: str | None = None):
    element = OxmlElement("w:p")
    after.addnext(element)
    paragraph = paragraph_from_element(doc, element)
    if style:
        paragraph.style = style
    if text:
        paragraph.add_run(text)
    return element


def add_heading(doc: Document, cursor, text: str, level: int = 2):
    return insert_paragraph(doc, cursor, text, f"Heading {level}")


def add_body(doc: Document, cursor, text: str):
    return insert_paragraph(doc, cursor, text, "Normal")


def add_list_item(doc: Document, cursor, text: str):
    element = insert_paragraph(doc, cursor, text, "List Paragraph")
    paragraph = paragraph_from_element(doc, element)
    p_pr = paragraph._p.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")
        num_id = OxmlElement("w:numId")
        num_id.set(qn("w:val"), "1")
        num_pr.append(ilvl)
        num_pr.append(num_id)
        p_pr.append(num_pr)
    return element


def add_code(doc: Document, cursor, text: str):
    element = insert_paragraph(doc, cursor, "", "No Spacing")
    paragraph = paragraph_from_element(doc, element)
    run = paragraph.add_run(text)
    run.font.name = "Consolas"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Consolas")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Consolas")
    run.font.size = Pt(9)
    p_pr = paragraph._p.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "F2F2F2")
    p_pr.append(shading)
    paragraph.paragraph_format.left_indent = Inches(0.25)
    paragraph.paragraph_format.right_indent = Inches(0.25)
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(4)
    return element


def add_action_note(doc: Document, cursor, title: str, instructions: str):
    element = insert_paragraph(doc, cursor, "", "Normal")
    paragraph = paragraph_from_element(doc, element)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    lead = paragraph.add_run(f"PENDIENTE DE ORLANDO - {title}: ")
    lead.bold = True
    lead.font.color.rgb = RGBColor(192, 0, 0)
    body = paragraph.add_run(instructions)
    body.bold = True
    body.font.color.rgb = RGBColor(156, 87, 0)
    p_pr = paragraph._p.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "FFF2CC")
    p_pr.append(shading)
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:color"), "F4B183")
    border.append(bottom)
    p_pr.append(border)
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(8)
    return element


def set_cell(cell, text: str, bold: bool = False, center: bool = False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(8.5)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def mark_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def add_table(doc: Document, cursor, title: str, headers: list[str], rows: list[list], widths: list[float] | None = None):
    caption = insert_paragraph(doc, cursor, f"Tabla 0. {title}", "Caption")
    paragraph_from_element(doc, caption).alignment = WD_ALIGN_PARAGRAPH.LEFT
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = False
    mark_header(table.rows[0])
    for index, header in enumerate(headers):
        set_cell(table.rows[0].cells[index], header, bold=True, center=True)
        shade_cell(table.rows[0].cells[index], "D9EAF7")
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            set_cell(cells[index], value, center=isinstance(value, (int, float)))
    if widths:
        width_twips = [round(width * 1440) for width in widths]
        if sum(width_twips) > 9000:
            factor = 9000 / sum(width_twips)
            width_twips = [max(240, round(width * factor)) for width in width_twips]
            width_twips[-1] += 9000 - sum(width_twips)
        for grid_col, width in zip(table._tbl.tblGrid.gridCol_lst, width_twips):
            grid_col.set(qn("w:w"), str(width))
        for row in table.rows:
            for index, width in enumerate(width_twips):
                tc_w = row.cells[index]._tc.get_or_add_tcPr().find(qn("w:tcW"))
                if tc_w is None:
                    tc_w = OxmlElement("w:tcW")
                    row.cells[index]._tc.get_or_add_tcPr().append(tc_w)
                tc_w.set(qn("w:type"), "dxa")
                tc_w.set(qn("w:w"), str(width))
    table_element = table._tbl
    table_element.getparent().remove(table_element)
    caption.addnext(table_element)
    source = insert_paragraph(doc, table_element, "Fuente: elaboración propia.", "Normal")
    paragraph_from_element(doc, source).runs[0].italic = True
    paragraph_from_element(doc, source).runs[0].font.size = Pt(9)
    return source


def set_picture_alt(run, title: str):
    drawing = run._element.find(qn("w:drawing"))
    if drawing is None:
        return
    for doc_pr in drawing.iter(qn("wp:docPr")):
        doc_pr.set("title", title)
        doc_pr.set("descr", title)


def add_figure(doc: Document, cursor, path: Path, title: str, width: float = 6.15):
    image_element = insert_paragraph(doc, cursor, "", "Normal")
    paragraph = paragraph_from_element(doc, image_element)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Inches(width))
    set_picture_alt(run, title)
    caption = insert_paragraph(doc, image_element, f"Figura 0. {title}", "Caption")
    paragraph_from_element(doc, caption).alignment = WD_ALIGN_PARAGRAPH.CENTER
    source = insert_paragraph(doc, caption, "Fuente: elaboración propia.", "Normal")
    paragraph_from_element(doc, source).alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_from_element(doc, source).runs[0].italic = True
    paragraph_from_element(doc, source).runs[0].font.size = Pt(9)
    return source


def add_field(paragraph, instruction: str, result: str = ""):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = result
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])


def convert_caption_to_seq(paragraph, label: str, number: int, title: str):
    paragraph.clear()
    paragraph.add_run(f"{label} ")
    add_field(paragraph, f"SEQ {label} \\* ARABIC", str(number))
    paragraph.add_run(f". {title}")


def renumber_captions(doc: Document):
    counts = {"Tabla": 0, "Figura": 0}
    for paragraph in doc.paragraphs:
        match = re.match(r"^(Tabla|Figura)\s+\d+\.\s*(.+)$", paragraph.text.strip())
        if not match:
            continue
        label, title = match.groups()
        counts[label] += 1
        convert_caption_to_seq(paragraph, label, counts[label], title)
    return counts


def set_update_fields(doc: Document):
    settings = doc.settings._element
    existing = settings.find(qn("w:updateFields"))
    if existing is None:
        existing = OxmlElement("w:updateFields")
        settings.append(existing)
    existing.set(qn("w:val"), "true")


def add_paragraphs(doc: Document, cursor, paragraphs: list[str]):
    for text in paragraphs:
        cursor = add_body(doc, cursor, text)
    return cursor


def revise_existing_sections(doc: Document):
    """Actualiza los apartados heredados que todavía describían un diseño futuro."""
    replacements = {
        "Este capítulo amplía el análisis del dominio de aplicación": (
            "Este capítulo amplía el análisis del dominio de aplicación y recoge los conceptos técnicos necesarios para comprender la solución implementada. En primer lugar, se describe el problema de inventario en pequeños comercios y su relación con la predicción de demanda. Posteriormente, se revisan aproximaciones existentes en forecasting minorista, demanda intermitente y evaluación de modelos. Finalmente, se incorpora un marco teórico que define las técnicas, métricas y herramientas utilizadas en el TFE."
        ),
        "Este capítulo presenta la problemática que motiva el TFE": (
            "Este capítulo presenta la problemática que motiva el TFE, delimita la intervención realizada y explica la organización general de la memoria. El propósito es que el lector comprenda desde el inicio por qué la gestión de inventario en pequeños comercios constituye un problema relevante para la inteligencia artificial aplicada, qué solución se implementó y cómo se estructuran los capítulos posteriores."
        ),
        "El capítulo 7 recoge las conclusiones derivadas": (
            "El capítulo 7 recoge las conclusiones derivadas del prototipo implementado y evaluado, las vincula con los objetivos y expone limitaciones y líneas de trabajo futuro. Finalmente, se incluyen las referencias bibliográficas utilizadas y un anexo técnico con la estructura del repositorio, los comandos de reproducción, los contratos de la API y el protocolo de usabilidad todavía pendiente."
        ),
        "A partir de las observaciones recibidas, el desarrollo específico de la contribución": (
            "A partir de las observaciones recibidas, la contribución se documenta en tres capítulos. El capítulo 4 define los requisitos, actores, casos de uso, criterios de aceptación y datos necesarios. El capítulo 5 describe la arquitectura implementada, el conjunto de datos, el pipeline ejecutado, los modelos comparados, la lógica de reposición, las interfaces y la automatización con n8n. El capítulo 6 presenta la evaluación realizada mediante métricas por horizonte, producto y categoría, pruebas funcionales, análisis de sensibilidad y verificación de requisitos. La validación de usabilidad con participantes no se ejecutó y se mantiene identificada como limitación."
        ),
        "Docker se considera para facilitar la ejecución local": (
            "Docker se utilizó para ejecutar localmente el panel Streamlit, la API y n8n con dependencias y puertos explícitos. La contenedorización redujo diferencias entre entornos y dejó una configuración reproducible mediante Dockerfile y Docker Compose. El alcance fue demostrativo y local: no se implementaron alta disponibilidad, balanceo de carga ni exposición segura a Internet (Docker, s. f.)."
        ),
        "La competición M5 Forecasting Accuracy evidenció": (
            "La competición M5 Forecasting Accuracy evidenció la complejidad de predecir ventas minoristas diarias a nivel de producto cuando se combinan miles de series, jerarquías, calendario, precios y eventos (Makridakis et al., 2022). Aunque sus datos corresponden a gran distribución y no a un único mini-market, M5 se empleó como referencia metodológica para estructurar series de producto y evaluar horizontes múltiples. No se utilizó el dataset M5 en la experimentación: el prototipo se ejecutó con el conjunto sintético descrito en el capítulo 5."
        ),
        "La evaluación de modelos de forecasting también ha sido objeto": (
            "La evaluación de modelos de forecasting también ha sido objeto de discusión. Hyndman y Koehler (2006) advierten que métricas como MAPE pueden distorsionarse cuando existen valores reales iguales o cercanos a cero. Por esta razón, se priorizaron MAE y SMAPE, complementadas con RMSE, MAPE y análisis por horizonte y producto. La evaluación cuantitativa y funcional se completó; la validación de usabilidad prevista según criterios de efectividad, eficiencia y satisfacción quedó pendiente y se reconoce como limitación (Brooke, 1996; International Organization for Standardization, 2018)."
        ),
        "La demanda intermitente constituye otro reto relevante": (
            "La demanda intermitente constituye otro reto relevante. Syntetos y Boylan (2005) mostraron que las series con muchos ceros y ventas ocasionales requieren atención específica, ya que algunos métodos y métricas pueden producir conclusiones poco fiables. En pequeños comercios es habitual que convivan productos básicos de alta rotación con referencias de consumo esporádico. Por ello, la evaluación del prototipo no se apoyó exclusivamente en porcentajes sensibles a valores cero: combinó MAE, RMSE, MAPE y SMAPE y conservó resultados por producto."
        ),
        "La principal ventaja de este enfoque es su flexibilidad": (
            "La principal ventaja de este enfoque es su flexibilidad para incorporar información exógena. No obstante, también exige preparación cuidadosa de datos, control de fugas temporales y evaluación rigurosa. En el prototipo se priorizaron modelos disponibles en scikit-learn por su madurez, documentación y facilidad de reproducción (Pedregosa et al., 2011). La selección final se justificó con resultados de backtesting, estabilidad entre horizontes, simplicidad y coste de mantenimiento."
        ),
        "La reposición de inventario busca determinar cuándo": (
            "La reposición de inventario busca determinar cuándo realizar un pedido y qué cantidad solicitar. En el prototipo, la decisión se modeló con demanda prevista, existencias, pedidos agregados en tránsito, tiempo de reposición, días de seguridad, periodo de revisión, tamaño de empaque y stock máximo. La demanda prevista cubre el tiempo de reposición más el periodo de revisión; el stock de seguridad añade un margen configurable; las existencias y unidades en tránsito reducen la necesidad; y el resultado se limita por capacidad y se redondea al empaque. La vida útil se almacenó como atributo, pero todavía no interviene en la fórmula."
        ),
        "n8n se considera en este TFE como un orquestador visual": (
            "n8n se utilizó como orquestador visual de un flujo ligero. El workflow implementado programa una ejecución diaria, solicita el reentrenamiento mediante POST /train y genera el pronóstico y las recomendaciones mediante GET /predict. Este alcance demuestra automatización reproducible sin pretender una plataforma MLOps empresarial: no incluye aprobación automática del modelo, monitorización de deriva ni notificaciones externas."
        ),
        "Este capítulo reemplaza el antiguo apartado de alcance funcional": (
            "Este capítulo reemplaza el antiguo apartado de alcance funcional y desarrolla la identificación de requisitos solicitada para un TFE de desarrollo de software. Se describen el contexto de uso, los actores, los casos de uso, los requisitos funcionales, los requisitos no funcionales y los criterios de aceptación que guiaron la implementación y su verificación."
        ),
        "El alcance imprescindible incluye preparación de datos": (
            "El alcance imprescindible quedó formado por preparación de datos, entrenamiento de un baseline y dos modelos candidatos, generación de pronósticos, cálculo de reposición, visualización en dashboard, API, almacenamiento local y un workflow n8n. Estos componentes se implementaron. Quedaron fuera la optimización multiobjetivo, la predicción probabilística, la integración en tiempo real con proveedores, la gestión contable, la compra automática y una arquitectura MLOps completa."
        ),
        "Los casos de uso se han definido a partir del flujo esperado": (
            "Los casos de uso se definieron a partir del flujo operativo del prototipo. Cada caso representa una interacción relevante entre el usuario o un componente automatizado y el sistema. Esta identificación conectó las necesidades del mini-market con los requisitos funcionales que debía cumplir la implementación."
        ),
        "Los requisitos imprescindibles aparecen implementados": (
            "Los requisitos imprescindibles aparecen implementados y verificados mediante evidencias de ejecución. Los requisitos deseables se incorporaron cuando no comprometían el alcance principal y, en caso contrario, se justificaron como limitaciones o trabajo futuro. Los requisitos opcionales no condicionaron la contribución principal. Esta clasificación permite que el tribunal distinga qué parte del sistema está demostrada y qué elementos siguen pendientes."
        ),
        "Los requisitos de datos aseguran que el modelo": (
            "Los requisitos de datos aseguraron que el modelo y la recomendación se construyeran sobre información mínima suficiente. La experimentación utilizó un conjunto sintético diario de diez productos con fecha, identificador, nombre, categoría, unidades, precio y promoción, unido a un inventario con existencias, pedidos en tránsito, tiempo de reposición, empaque, vida útil y stock máximo. El mismo contrato admite sustituir los archivos por datos anonimizados de una tienda real, siempre que se conserven el nivel producto-fecha, los tipos de datos y las validaciones indicadas en la Tabla 6."
        ),
        "El trabajo plantea el desarrollo de un prototipo software": (
            "El trabajo desarrolló un prototipo software capaz de predecir la demanda diaria por producto y convertirla en recomendaciones concretas de reposición para un mini-market. La entrada utilizada fue un histórico sintético y reproducible de ventas con calendario, precio y promociones, complementado por un inventario operativo con existencias, pedidos en tránsito, tiempo de reposición, tamaño de empaque y stock máximo. Como salida, el sistema generó pronósticos para horizontes de 7, 14 y 30 días y cantidades sugeridas de pedido calculadas a partir del stock disponible, la demanda prevista, el stock de seguridad y el periodo de revisión."
        ),
        "Desde el punto de vista técnico, el trabajo comparará": (
            "Desde el punto de vista técnico, se comparó un baseline estacional con Ridge Regression y Random Forest. El propósito no fue demostrar que el modelo más complejo fuera siempre superior, sino seleccionar una alternativa reproducible y suficientemente precisa para apoyar la decisión de inventario. La selección se realizó mediante separación temporal y métricas MAE, RMSE, MAPE y SMAPE. Como recomiendan Hyndman y Koehler (2006), las métricas porcentuales se interpretaron junto con errores absolutos y resultados desagregados."
        ),
        "El alcance se acota a un prototipo local orientado": (
            "El alcance se acotó a un prototipo local de apoyo a la decisión. El sistema no emite órdenes de compra ni reemplaza el criterio del responsable del negocio: permite revisar el pronóstico, modificar días de seguridad y periodo de revisión, y descargar la recomendación. Se utilizó un conjunto sintético de diez productos porque no se dispuso de un histórico anonimizado de una tienda real con la calidad necesaria. Esta decisión permitió ejecutar y documentar el flujo completo, pero se mantuvo como la principal limitación de validez externa."
        ),
        "Para horizontes de 7, 14 y 30 días, el sistema deberá": (
            "El sistema generó predicciones para horizontes de 7, 14 y 30 días, que representan una revisión semanal, una planificación quincenal y una aproximación mensual. La evaluación respetó el orden cronológico: para cada horizonte se ocultó el tramo final correspondiente y se entrenó exclusivamente con fechas anteriores. Este diseño evitó mezclar aleatoriamente pasado y futuro y redujo el riesgo de fuga de información."
        ),
        "El MAE mide el promedio de los errores absolutos": (
            "El MAE mide el promedio de los errores absolutos y se expresa en unidades del producto. El RMSE penaliza con mayor intensidad los errores grandes. El SMAPE expresa un error relativo simétrico y facilita comparar series con escalas distintas, aunque también requiere cautela cuando los valores son muy bajos. El MAPE se calculó como indicador complementario e ignoró observaciones con demanda real igual a cero para evitar divisiones indefinidas."
        ),
        "La validación temporal se realizará": (
            "La validación temporal se ejecutó mediante un corte cronológico específico para cada horizonte. Se entrenó con el bloque histórico anterior al origen y se pronosticó recursivamente el periodo oculto, de modo que el sistema solo dispuso del pasado al generar cada predicción. Los tres modelos se evaluaron sobre las mismas fechas y productos, y sus resultados se compararon con un baseline estacional de siete días."
        ),
        "La lógica de reposición se mantendrá explicable": (
            "La lógica de reposición se implementó como una regla explicable. Para cada producto se sumó la demanda prevista durante el tiempo de reposición más el periodo de revisión, se añadió un stock de seguridad expresado en días y se restaron las existencias y los pedidos en tránsito. La necesidad positiva se limitó por el stock máximo y se redondeó al empaque del proveedor. La salida fue una sugerencia revisable, no una orden automática."
        ),
        "Streamlit se utilizará como tecnología de interfaz": (
            "Streamlit se utilizó como tecnología de interfaz por su integración directa con Python y su adecuación para prototipos de datos (Streamlit, s. f.). El panel permite consultar el histórico y el pronóstico por producto, escoger el horizonte, modificar parámetros de reposición, interpretar indicadores mediante ayudas contextuales y descargar resultados. Los rótulos se presentaron en español y los modelos conservaron sus nombres técnicos reconocibles."
        ),
        "El prototipo estará orientado a un mini-market": (
            "El prototipo se orientó a un mini-market que desea apoyar la reposición con datos históricos. Se ejecutó localmente y no dependió de una integración con ERP o punto de venta. El usuario pudo consultar predicciones, revisar recomendaciones y modificar parámetros operativos; el sistema no realizó pagos ni generó órdenes automáticas. Su función fue informar y conservar al responsable como decisor final."
        ),
        "En este TFE, el alcance mínimo validable será": (
            "El alcance mínimo validable se concretó en un flujo de extremo a extremo: dataset procesado, comparación frente a baseline, modelo serializado, panel, API, persistencia local, recomendación de reposición y automatización con n8n. La evaluación cuantitativa y las pruebas funcionales se completaron. La validación con usuarios no se ejecutó y, por tanto, se registró como objetivo parcialmente cumplido y como trabajo pendiente, sin sustituirla por resultados simulados."
        ),
        "La validación de requisitos se realizará": (
            "La validación de requisitos se realizó mediante una matriz de trazabilidad que conectó cada requisito con un componente y una evidencia concreta. Se utilizaron scripts ejecutados, archivos CSV, tablas de métricas, pruebas automatizadas, respuestas HTTP, una base SQLite y el workflow exportado de n8n. Las capturas interactivas que todavía debe incorporar el autor permanecen señaladas en amarillo."
        ),
    }
    for paragraph in doc.paragraphs:
        current = paragraph.text.strip()
        for prefix, replacement in replacements.items():
            if current.startswith(prefix):
                paragraph.text = replacement
                paragraph.style = "Normal"
                break

    table_replacements = {
        "Existe un flujo exportado y documentado que ejecuta una tarea programada.": (
            "El workflow versionado contiene un trigger diario y llamadas a /train y /predict mediante la red interna de Docker; ambos endpoints respondieron correctamente."
        ),
        "Realizar tareas guiadas; la encuesta será opcional": (
            "Realizar tareas guiadas y responder la escala SUS."
        ),
        "El usuario completa tareas básicas de carga, consulta e interpretación": (
            "Tres o más participantes completan tareas de consulta e interpretación y responden la escala SUS; este criterio permanece pendiente de validación."
        ),
        "Modificar stock actual, stock de seguridad, lead time, lote mínimo y umbral.": (
            "Modificar los días globales de stock de seguridad y el periodo de revisión desde el panel."
        ),
        "Parámetros guardados y aplicados al cálculo.": (
            "Las recomendaciones se recalculan con los parámetros seleccionados durante la sesión."
        ),
        "Ejecutar flujo programado para generar reporte o alerta de productos críticos.": (
            "Ejecutar el flujo programado que reentrena el modelo y regenera pronósticos y recomendaciones mediante la API."
        ),
        "Reporte disponible o alerta emitida.": (
            "Pronóstico y recomendaciones actualizados por la ejecución automatizada."
        ),
        "El sistema debe permitir modificar parámetros de reposición por producto.": (
            "El sistema debe permitir modificar los días de seguridad y el periodo de revisión usados en la reposición."
        ),
        "Los parámetros modificados actualizan la recomendación.": (
            "Al cambiar cualquiera de los dos parámetros, el panel recalcula las cantidades y prioridades."
        ),
        "El sistema debe calcular una cantidad sugerida de pedido con base en forecast": (
            "El sistema debe calcular una cantidad sugerida con forecast, existencias, pedidos en tránsito, tiempo de reposición, seguridad, revisión, empaque y stock máximo."
        ),
        "La recomendación se muestra con explicación de sus componentes.": (
            "La salida muestra demanda prevista, existencias, cobertura, cantidad redondeada y prioridad."
        ),
        "Calcula cantidades sugeridas usando stock, lead time, stock de seguridad y lote mínimo.": (
            "Calcula la necesidad con forecast, existencias, pedidos en tránsito, cobertura, empaque y stock máximo."
        ),
        "Genera un reporte básico de reposición a partir de resultados ya calculados.": (
            "Programa el reentrenamiento y la regeneración de pronósticos y recomendaciones mediante la API."
        ),
        "Variables como día de semana, mes, festivo o evento.": (
            "Variables derivadas de fecha: día, semana, mes, fin de mes, quincena y componentes cíclicos."
        ),
        "Derivable desde la fecha o disponible en dataset.": (
            "Todas las variables de calendario se derivan de la fecha sin consultar información futura."
        ),
        "Variables exógenas opcionales para mejorar predicción.": (
            "Precio unitario obligatorio y promoción opcional como variables exógenas."
        ),
        "Se usan solo si están disponibles de forma consistente.": (
            "El precio se valida y la promoción ausente se completa con cero."
        ),
        "Datos necesarios para transformar forecast en recomendación.": (
            "Existencias, pedidos en tránsito, lead time y empaque; stock máximo y vida útil son opcionales."
        ),
        "Pueden ser cargados manualmente o definidos como parámetros configurables.": (
            "El archivo de inventario cumple el esquema y los días de seguridad/revisión se suministran al ejecutar."
        ),
        "Los procesos principales se ejecutan en tiempos aceptables para uso académico/local.": (
            "Entrenamiento, evaluación multihorizonte y predicción de 14 días finalizan individualmente en menos de 60 segundos en el equipo local del autor."
        ),
    }
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                current = cell.text.strip()
                for prefix, replacement in table_replacements.items():
                    if current.startswith(prefix):
                        cell.text = replacement
                        break


def repair_spelling(doc: Document):
    """Corrige errores ortográficos heredados sin alterar campos de Word."""
    replacements = {
        "medíante": "mediante",
        "díaria": "diaria",
        "díarias": "diarias",
        "díario": "diario",
        "cotidíana": "cotidiana",
        "inmedíata": "inmediata",
        "medíana": "mediana",
        "medías": "medias",
        "predicciónes": "predicciones",
        "relaciónadas": "relacionadas",
        "relaciónes": "relaciones",
        "Inteligencia artificial": "Inteligencia Artificial",
    }

    def repair_runs(paragraph):
        for run in paragraph.runs:
            original = run.text
            repaired = original
            for wrong, correct in replacements.items():
                repaired = repaired.replace(wrong, correct)
            if repaired != original:
                run.text = repaired

    for paragraph in doc.paragraphs:
        repair_runs(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    repair_runs(paragraph)


def italicize_technical_terms(doc: Document):
    """Aplica cursiva a anglicismos metodológicos en el cuerpo académico."""
    terms = re.compile(
        r"\b(Random Forest|Ridge Regression|Seasonal Naive|forecasting|forecast|baseline|backtesting|dataset|pipeline|workflow|lead time|snapshot|stockouts|bias|rollback)\b",
        re.IGNORECASE,
    )
    in_references = False
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == "Referencias bibliográficas":
            in_references = True
        if in_references or paragraph.style.name != "Normal":
            continue
        if "PENDIENTE DE ORLANDO" in paragraph.text or len(paragraph.runs) != 1:
            continue
        text = paragraph.text
        matches = list(terms.finditer(text))
        if not matches:
            continue
        paragraph.clear()
        position = 0
        for match in matches:
            if match.start() > position:
                paragraph.add_run(text[position:match.start()])
            emphasized = paragraph.add_run(match.group(0))
            emphasized.italic = True
            position = match.end()
        if position < len(text):
            paragraph.add_run(text[position:])


def italicize_reference_titles(doc: Document):
    """Aplica las cursivas bibliográficas principales exigidas por APA 7."""
    targets = {
        "Aamer, A. M.": "Operations and Supply Chain Management, 14",
        "Brooke, J.": "Usability evaluation in industry",
        "Carbonneau, R.": "European Journal of Operational Research, 184",
        "Docker. (s. f.)": "Docker: Accelerated container application development",
        "Hyndman, R. J., & Athanasopoulos": "Forecasting: Principles and practice",
        "Hyndman, R. J., & Koehler": "International Journal of Forecasting, 22",
        "IBM. (s. f.)": "Conceptos básicos de ayuda de CRISP-DM",
        "International Organization for Standardization": "ISO 9241-11:2018. Ergonomics of human-system interaction - Part 11: Usability: Definitions and concepts",
        "Makridakis, S.": "International Journal of Forecasting, 38",
        "n8n. (s. f.)": "AI workflow automation platform - n8n",
        "Pedregosa, F.": "Journal of Machine Learning Research, 12",
        "Russell, S.": "Artificial intelligence: A modern approach",
        "Shearer, C.": "Journal of Data Warehousing, 5",
        "Silver, E. A.": "Inventory and production management in supply chains",
        "Streamlit. (s. f.)": "Streamlit documentation",
        "Syntetos, A. A.": "International Journal of Forecasting, 21",
    }
    in_references = False
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if text.strip() == "Referencias bibliográficas":
            in_references = True
            continue
        if in_references and paragraph.style.name == "Heading 1":
            break
        if not in_references:
            continue
        target = next((title for prefix, title in targets.items() if text.strip().startswith(prefix)), None)
        if not target or target not in text:
            continue
        before, after = text.split(target, 1)
        paragraph.clear()
        paragraph.add_run(before)
        emphasized = paragraph.add_run(target)
        emphasized.italic = True
        paragraph.add_run(after)


def build_methodology(doc: Document, sales: pd.DataFrame):
    cursor = clear_section(doc, "Objetivos concretos y metodología de trabajo")
    cursor = add_heading(doc, cursor, "Objetivo general")
    cursor = add_body(
        doc,
        cursor,
        "El objetivo general fue desarrollar y validar una herramienta software reproducible que predijera la demanda diaria por producto y generara recomendaciones explicables de reposición para un mini-market. La solución debía integrar preparación de datos, comparación de modelos, consulta mediante interfaz y API, persistencia local y automatización del ciclo de ejecución. El alcance se acotó a diez productos y a horizontes de 7, 14 y 30 días para mantener una experimentación controlada, verificable y compatible con recursos informáticos domésticos.",
    )
    cursor = add_heading(doc, cursor, "Objetivos específicos y resultados verificables")
    objectives = [
        ["OE1", "Preparar un conjunto de ventas y un pipeline reproducible", "7.310 observaciones diarias, diez productos, validación de esquema y generación determinista con semilla fija", "Cumplido"],
        ["OE2", "Comparar un baseline y dos modelos de aprendizaje automático", "Backtesting de baseline estacional, Ridge y Random Forest en 7, 14 y 30 días", "Cumplido"],
        ["OE3", "Implementar una regla explicable de reposición", "Cálculo de demanda durante cobertura, stock de seguridad, cantidad redondeada y prioridad", "Cumplido"],
        ["OE4", "Construir una interfaz y una API", "Panel Streamlit, endpoints FastAPI y snapshot de reportes en SQLite", "Cumplido técnicamente"],
        ["OE5", "Evaluar precisión, funcionamiento y utilidad", "Métricas MAE, RMSE, MAPE y SMAPE, seis pruebas automatizadas; validación con usuarios no ejecutada", "Parcial"],
        ["OE6", "Automatizar el flujo operativo", "Workflow n8n 2.30.5 que invoca entrenamiento y predicción mediante la API, versionado junto con Docker Compose", "Cumplido técnicamente; captura pendiente"],
    ]
    cursor = add_table(doc, cursor, "Trazabilidad de objetivos y resultados", ["Id.", "Objetivo", "Evidencia obtenida", "Estado"], objectives, [0.55, 2.15, 3.25, 1.05])
    cursor = add_heading(doc, cursor, "Diseño metodológico")
    cursor = add_paragraphs(doc, cursor, [
        "El trabajo se organizó con CRISP-DM como marco principal porque permite relacionar la comprensión del negocio, los datos, el modelado, la evaluación y el despliegue en un proceso iterativo. No se utilizó como una secuencia rígida: los hallazgos del backtesting obligaron a revisar la selección del modelo y a separar la métrica de ajuste a un paso de la evaluación multihorizonte. Esta adaptación es coherente con la naturaleza iterativa descrita por Shearer (2000) y con la guía de IBM (s. f.) para proyectos de minería de datos.",
        "En la comprensión del negocio se definió la decisión operativa que debía apoyar el sistema: determinar cuánto pedir de cada producto para cubrir el tiempo de reposición y un periodo de revisión, descontando las existencias disponibles. Esta formulación evitó tratar la predicción como un fin aislado. La salida de interés no fue únicamente una curva de demanda, sino una recomendación trazable hasta el pronóstico, los parámetros de inventario y el stock actual.",
        "En la comprensión de los datos se inspeccionaron cobertura temporal, granularidad, duplicados, valores faltantes, rangos y consistencia entre ventas e inventario. El dataset quedó compuesto por " + f"{len(sales):,}".replace(",", ".") + f" filas, que corresponden a 731 días para cada uno de diez productos. La fecha mínima fue el {sales['date'].min().strftime('%d/%m/%Y')} y la máxima el {sales['date'].max().strftime('%d/%m/%Y')}. La unidad de análisis fue producto-día; esta definición se mantuvo en la preparación, el entrenamiento, la evaluación y la generación de recomendaciones.",
        "La preparación transformó las ventas al nivel diario y preservó precio y promoción. Los días sin ventas se completaron con cero dentro del intervalo de cada producto, mientras que el precio se propagó desde el último valor conocido. Para evitar fuga de información, los retardos y estadísticas móviles se desplazaron un día antes de calcularse. Así, una fila fechada en t solo utilizó datos disponibles hasta t-1. Esta decisión fue comprobada mediante pruebas sobre la cardinalidad y el pipeline; las variables lag_1, lag_2, lag_7, lag_14 y lag_28 se generan siempre con valores anteriores a la fecha objetivo.",
        "La fase de modelado comparó tres familias deliberadamente distintas. El baseline estacional replicó la demanda de siete días atrás y estableció una referencia mínima interpretable. Ridge aportó un modelo lineal regularizado, útil para medir cuánto explicaban las variables construidas sin relaciones no lineales. Random Forest capturó interacciones y umbrales entre calendario, precio, promociones, retardos y ventanas móviles sin requerir una red neuronal. La comparación respondió al alcance del máster y al criterio de mantener el sistema ejecutable en local.",
        "La evaluación se diseñó como backtesting con un origen temporal específico por horizonte. Para cada horizonte se reservó el tramo final correspondiente y se entrenó con todas las observaciones anteriores. Los pronósticos fueron recursivos: cada predicción se incorporó al historial temporal necesario para construir los retardos del día siguiente. Se calcularon MAE, RMSE, MAPE y SMAPE sobre exactamente el mismo conjunto de fechas, productos y observaciones para los tres modelos, evitando comparaciones con muestras distintas.",
        "La fase de despliegue se materializó mediante módulos Python, una aplicación Streamlit, una API FastAPI, SQLite, Docker Compose y un workflow n8n. El control de versiones y los archivos de dependencias permiten reconstruir el entorno. El despliegue se consideró local y demostrativo: no se realizaron pruebas de alta disponibilidad, concurrencia elevada ni operación continua porque esas capacidades no formaban parte del alcance aprobado.",
    ])
    cursor = add_heading(doc, cursor, "Protocolo experimental y prevención de fuga temporal")
    cursor = add_paragraphs(doc, cursor, [
        "El protocolo se fijó antes de interpretar los resultados. La variable objetivo fue units_sold. Las variables predictoras fueron identificadores codificados, componentes de calendario, precio, promoción, retardos de 1, 2, 7, 14 y 28 días y medias y desviaciones móviles de 7, 14 y 28 días. Los registros sin historia suficiente se eliminaron únicamente después de construir las variables; se aplicó el mismo criterio a todos los modelos de aprendizaje automático.",
        "Los horizontes de 7, 14 y 30 días se evaluaron por separado porque representan decisiones distintas: revisión semanal, planificación quincenal y compra mensual. Cada horizonte produjo h × 10 predicciones por modelo: 70, 140 y 300 registros respectivamente. Esta regla de cardinalidad fue añadida a las pruebas automatizadas, lo que resuelve explícitamente la discrepancia de la versión anterior, donde se informó de 70 registros para diez productos y catorce días.",
        "El criterio principal de selección fue el SMAPE medio de los tres horizontes. Se eligió SMAPE porque mantiene una escala porcentual y reduce algunos problemas de MAPE cuando la demanda real se acerca a cero, aunque no elimina por completo la inestabilidad en series de bajo volumen (Hyndman y Koehler, 2006). MAE se utilizó como lectura complementaria en unidades físicas y RMSE para detectar la influencia de errores grandes.",
        "La selección no se basó en la métrica de entrenamiento ni en una única ventana. Ridge obtuvo el menor SMAPE en 7 y 14 días, pero Random Forest fue más estable en 30 días y alcanzó el mejor promedio multihorizonte. Por esta razón se serializó Random Forest como modelo operativo. La decisión y los valores completos quedaron almacenados en reports/model_ranking.csv, no solo descritos en la memoria.",
    ])
    cursor = add_heading(doc, cursor, "Reproducibilidad y gestión de evidencias")
    cursor = add_paragraphs(doc, cursor, [
        "La reproducibilidad se abordó en cuatro niveles. Primero, la generación sintética utiliza una semilla fija y produce archivos CSV con esquema documentado. Segundo, requirements.txt fija las dependencias principales. Tercero, los scripts separan entrenamiento, evaluación, predicción y generación de figuras. Cuarto, los reportes intermedios se guardan como CSV, lo que permite recalcular tablas sin transcribir valores manualmente.",
        "Cada afirmación cuantitativa de los capítulos 5 y 6 se vinculó con un artefacto: métricas por horizonte, métricas por producto, ranking, predicciones de backtesting, recomendaciones y análisis de sensibilidad. Las capturas de API y pruebas fueron obtenidas de ejecuciones reales. En cambio, las capturas de Streamlit, n8n y Docker se marcaron como acciones pendientes del autor porque requieren abrir los servicios interactivos en su entorno y mostrar fecha, datos y estado de ejecución. Esta distinción evita presentar maquetas como evidencia.",
        "El diseño no incluyó ajuste exhaustivo de hiperparámetros sobre el segmento de prueba. Los parámetros se mantuvieron constantes y moderados para limitar el riesgo de sobreajuste y el coste de cómputo. Una optimización futura debería incorporar validación temporal anidada; realizarla sobre los mismos horizontes usados para reportar resultados produciría estimaciones optimistas.",
    ])
    return cursor


def build_development(doc: Document, sales: pd.DataFrame, inventory: pd.DataFrame):
    cursor = clear_section(doc, "Diseño y desarrollo de la solución")
    cursor = add_heading(doc, cursor, "Arquitectura implementada")
    cursor = add_paragraphs(doc, cursor, [
        "La solución se implementó como un repositorio monolítico modular. La capa de datos ingiere CSV y puede persistir el snapshot actual de resultados en SQLite; la capa de aprendizaje prepara variables, entrena, evalúa y serializa modelos; la capa de negocio transforma pronósticos en recomendaciones; y la capa de presentación expone los resultados mediante Streamlit y FastAPI. n8n invoca la API y Docker Compose describe los servicios. La separación permite sustituir una interfaz o un modelo sin reescribir la regla de inventario.",
        "El flujo principal comienza con sales_sample.csv e inventory_sample.csv. train.py valida y transforma las ventas, ejecuta la comparación multihorizonte y guarda el modelo seleccionado. predict.py carga el artefacto, genera una trayectoria diaria por producto, agrega la demanda de cobertura, calcula la cantidad sugerida y escribe CSV y SQLite. Streamlit y FastAPI reutilizan las mismas funciones de dominio para calcular sus respuestas; no duplican la lógica de modelado ni de reposición.",
        "La arquitectura se diseñó para ejecución local y sin servicios de pago. Esta decisión reduce barreras de reproducción y protege datos potencialmente sensibles del comercio. También introduce límites: un único archivo SQLite y procesos locales no son apropiados para alta concurrencia ni múltiples tiendas escribiendo simultáneamente. Para el prototipo, esa compensación fue aceptable.",
    ])
    cursor = add_figure(doc, cursor, ARCH, "Arquitectura implementada del sistema de predicción y reposición", 5.9)
    cursor = add_heading(doc, cursor, "Conjunto de datos utilizado")
    cursor = add_paragraphs(doc, cursor, [
        "Se utilizó un conjunto sintético generado específicamente para representar un mini-market. No corresponde a una tienda real ni contiene datos personales. La elección respondió a la imposibilidad de obtener a tiempo un histórico anonimizado con granularidad diaria, precio, promoción e inventario. Para que la experimentación siguiera siendo auditable, el generador quedó incluido en el repositorio y utiliza parámetros explícitos y una semilla determinista.",
        "El archivo de ventas contiene diez referencias ficticias agrupadas en siete categorías: aseo, básicos, bebidas, despensa, frutas, lácteos y panadería. Cada producto cuenta con una observación diaria desde el 1 de enero de 2024 hasta el 31 de diciembre de 2025. La demanda combina nivel base por producto, efecto de fin de semana y quincena, estacionalidad por categoría, promociones aleatorias, tendencia suave, ruido gaussiano y muestreo Poisson. El propósito no fue reproducir toda la complejidad comercial, sino disponer de señales conocidas para probar el pipeline de extremo a extremo.",
        "El archivo de inventario contiene product_id, product_name, category, stock_on_hand, on_order, lead_time_days, pack_size, perishability_days y max_stock. Los días de seguridad y el periodo de revisión no forman parte del CSV: son parámetros que el usuario suministra al ejecutar el pronóstico. La regla sí descuenta los pedidos en tránsito mediante on_order y limita la necesidad por max_stock; perishability_days se conserva como atributo descriptivo, pero todavía no modifica la cantidad sugerida. No se incluyeron costes, presupuesto ni restricciones por proveedor.",
    ])
    dataset_rows = [
        ["Ventas", "sales_sample.csv", f"{len(sales):,}".replace(",", "."), f"{sales['date'].min().date()} a {sales['date'].max().date()}", "producto-día"],
        ["Inventario", "inventory_sample.csv", str(len(inventory)), "corte operativo", "producto"],
    ]
    cursor = add_table(doc, cursor, "Características de los archivos de entrada", ["Entidad", "Archivo", "Filas", "Cobertura", "Granularidad"], dataset_rows, [1.0, 1.7, 0.65, 1.65, 1.15])
    cursor = add_figure(doc, cursor, FIGURES / "ventas_mensuales_categoria.png", "Volumen mensual de ventas sintéticas por categoría")
    cursor = add_heading(doc, cursor, "Perfil de productos y control de calidad")
    product_profile = (
        sales.groupby(["product_id", "product_name", "category"], as_index=False)
        .agg(
            dias=("date", "nunique"),
            demanda_media=("units_sold", "mean"),
            porcentaje_promocion=("promo", lambda values: 100 * values.mean()),
        )
        .sort_values("product_id")
    )
    product_rows = [
        [
            row.product_id,
            row.product_name,
            row.category,
            int(row.dias),
            f"{row.demanda_media:.2f}",
            f"{row.porcentaje_promocion:.2f} %",
        ]
        for row in product_profile.itertuples(index=False)
    ]
    cursor = add_table(
        doc,
        cursor,
        "Productos incluidos en la experimentación",
        ["Id.", "Producto", "Categoría", "Días", "Demanda media", "Días en promoción"],
        product_rows,
        [0.55, 1.75, 1.05, 0.55, 1.1, 1.25],
    )
    inventory_rows = [
        [
            row.product_id,
            int(row.stock_on_hand),
            int(row.on_order),
            int(row.lead_time_days),
            int(row.pack_size),
            int(row.perishability_days),
            int(row.max_stock),
        ]
        for row in inventory.sort_values("product_id").itertuples(index=False)
    ]
    cursor = add_table(
        doc,
        cursor,
        "Estado inicial y parámetros fijos de inventario",
        ["Id.", "Stock", "En tránsito", "Lead time", "Empaque", "Vida útil", "Stock máximo"],
        inventory_rows,
        [0.55, 0.75, 0.85, 0.85, 0.8, 0.85, 1.0],
    )
    duplicate_rows = int(sales.duplicated(["date", "product_id"]).sum())
    missing_values = int(sales.isna().sum().sum())
    promo_rows = int(sales["promo"].sum())
    sales_rows_text = f"{len(sales):,}".replace(",", ".")
    total_units_text = f"{int(sales['units_sold'].sum()):,}".replace(",", ".")
    mean_units_text = f"{sales['units_sold'].mean():.2f}".replace(".", ",")
    promo_pct_text = f"{100 * sales['promo'].mean():.2f}".replace(".", ",")
    cursor = add_paragraphs(doc, cursor, [
        f"La auditoría del archivo de ventas confirmó {sales_rows_text} filas, {sales['date'].nunique()} fechas y {sales['product_id'].nunique()} productos, con {missing_values} valores ausentes y {duplicate_rows} duplicados producto-fecha. El volumen total fue de {total_units_text} unidades; la demanda media alcanzó {mean_units_text} unidades por producto-día, con mínimo de {int(sales['units_sold'].min())} y máximo de {int(sales['units_sold'].max())}.",
        f"Se registraron {promo_rows} observaciones promocionales, equivalentes al {promo_pct_text} % de las filas. Estos controles no demuestran realismo externo, pero sí establecen que las métricas posteriores no provienen de columnas vacías, duplicados o coberturas desiguales. La semilla 42 permite regenerar exactamente el mismo escenario mediante scripts/generate_sample_data.py.",
    ])
    cursor = add_heading(doc, cursor, "Validación, limpieza y transformación")
    cursor = add_paragraphs(doc, cursor, [
        "La función load_sales_csv verifica las columnas obligatorias, convierte date a fecha, normaliza product_id como texto y transforma units_sold y unit_price a valores numéricos; los valores no convertibles se sustituyen por cero. La función daily_product_series agrega duplicados producto-fecha: suma las unidades, calcula el precio medio y conserva el máximo del indicador promocional. El dataset generado no contenía duplicados ni valores faltantes, pero estas transformaciones permiten recibir archivos con más de un movimiento diario.",
        "Después de la agregación se construye un calendario completo por producto. Los días ausentes reciben cero unidades, lo que distingue ausencia de venta de ausencia de registro dentro del conjunto sintético. El precio se completa hacia delante y hacia atrás dentro de cada producto. La promoción se conserva durante la agregación y los valores faltantes se sustituyen por cero; esta corrección fue importante porque una versión preliminar descartaba accidentalmente el indicador al reagrupar.",
        "El pipeline calcula día de la semana, día del mes, semana ISO, mes, fin de semana, quincena y componentes seno-coseno para los ciclos semanal y mensual. product_id y category se codifican con pd.get_dummies; la lista de columnas resultante se guarda dentro del ModelBundle y se reutiliza en inferencia. Después de descartar los primeros 28 días de cada producto por falta de historia suficiente, quedaron 7.030 filas supervisadas y 40 variables predictoras. El estimador se implementó con scikit-learn (Pedregosa et al., 2011).",
        "Las variables autorregresivas se construyen por producto. lag_1 y lag_2 capturan persistencia inmediata; lag_7 representa el patrón semanal; lag_14 y lag_28 aportan memoria de dos y cuatro semanas. Las medias y desviaciones móviles resumen nivel y variabilidad. Todas las ventanas se desplazan antes de agregarse, de modo que una observación nunca incorpora su propio valor objetivo.",
    ])
    feature_rows = [
        ["Identificación", "product_id, category", "Segmentación de patrones por referencia y categoría"],
        ["Calendario", "day_of_week, day_of_month, week_of_year, month, is_weekend, is_payday, seno/coseno", "Ciclos semanales y mensuales"],
        ["Comerciales", "unit_price, promo", "Precio y promociones observadas"],
        ["Retardos", "lag_1, lag_2, lag_7, lag_14, lag_28", "Historia disponible antes de la fecha pronosticada"],
        ["Ventanas", "rolling_mean y rolling_std de 7, 14 y 28 días", "Nivel y dispersión recientes"],
    ]
    cursor = add_table(doc, cursor, "Variables finales del modelo", ["Grupo", "Variables", "Función"], feature_rows, [1.1, 3.0, 2.7])
    cursor = add_figure(doc, cursor, PIPELINE, "Pipeline reproducible desde la ingestión hasta la recomendación")
    cursor = add_heading(doc, cursor, "Modelos implementados y configuración real")
    cursor = add_paragraphs(doc, cursor, [
        "El baseline seasonal_naive_7d predice para cada producto el valor observado siete días antes. Durante la evaluación recursiva, una predicción puede convertirse en el valor de referencia de una fecha posterior. Este comportamiento reproduce la disponibilidad real de información y evita utilizar observaciones futuras. El baseline no requiere ajuste y ofrece una referencia clara para juzgar si los modelos añaden valor.",
        "Ridge se configuró con alpha igual a 1,0. Las variables numéricas se estandarizaron y las categóricas se codificaron dentro del pipeline. La regularización L2 estabiliza coeficientes cuando existen retardos y ventanas correlacionadas. No se ejecutó una búsqueda automática de alpha, por lo que el resultado representa una configuración razonable y reproducible, no el máximo rendimiento alcanzable.",
        "Random Forest se configuró con 160 árboles, profundidad máxima de 12, mínimo de tres muestras por hoja, random_state 42 y ejecución paralela. La profundidad acotada y el tamaño mínimo de hoja reducen la memorización de ruido. El número de árboles ofreció estabilidad suficiente con un tiempo de entrenamiento inferior a un minuto en el equipo de desarrollo.",
        "Los modelos de aprendizaje automático pronostican de forma recursiva. Para cada día futuro se crea una fila por producto, se calculan variables con el historial disponible, se ejecuta predict y se añade el resultado no negativo al historial. La promoción futura se fija en cero porque no se proporcionó un calendario comercial. Esta suposición es conservadora y se declara explícitamente: si una promoción futura es conocida, la API o el dataset deberían permitir suministrarla.",
    ])
    config_rows = [
        ["Estacional ingenuo", "valor t-7", "Sin hiperparámetros", "Referencia mínima"],
        ["Ridge", "Pipeline lineal regularizado", "alpha=1,0", "Interpretación y estabilidad"],
        ["Random Forest", "Bosque de regresión", "160 árboles; max_depth=12; min_samples_leaf=3; random_state=42", "Relaciones no lineales"],
    ]
    cursor = add_table(doc, cursor, "Configuración ejecutada de los modelos", ["Modelo", "Implementación", "Configuración", "Propósito"], config_rows, [1.15, 1.65, 2.7, 1.3])
    cursor = add_heading(doc, cursor, "Selección y serialización del modelo")
    cursor = add_paragraphs(doc, cursor, [
        "train.py llama a la evaluación multihorizonte y ordena los candidatos por SMAPE medio. Random Forest ocupó el primer lugar con 28,33 %, seguido por Ridge con 31,22 % y el baseline con 35,42 %. El modelo ganador se volvió a ajustar con todas las filas utilizables y se guardó en models/demand_model.joblib junto con metadatos, lista de variables, fecha de entrenamiento y parámetros.",
        "La serialización con joblib evita entrenar al iniciar la interfaz o la API. El archivo solo debe cargarse desde una fuente confiable, dado que los formatos de objetos Python no son seguros frente a artefactos manipulados. En un despliegue real se añadirían firma, checksum y un registro de modelos con control de acceso.",
    ])
    cursor = add_heading(doc, cursor, "Motor explicable de reposición")
    cursor = add_paragraphs(doc, cursor, [
        "La cantidad sugerida se calcula a partir de la demanda pronosticada durante el periodo de cobertura, definido como lead_time_days + review_period_days. El stock de seguridad se aproxima como la demanda diaria media de todo el horizonte multiplicada por safety_stock_days. El stock disponible suma stock_on_hand y on_order. La necesidad neta es demanda de cobertura más stock de seguridad menos stock disponible; los valores negativos se truncan a cero. La lógica adapta conceptos básicos de gestión de inventario descritos por Silver et al. (2016).",
        "Cuando existe max_stock, la necesidad se limita al espacio disponible antes de redondear. Después se redondea hacia arriba al múltiplo de pack_size para producir una cantidad compatible con el empaque. La prioridad es alta cuando la cobertura del stock disponible es menor que el tiempo de reposición, media cuando se recomienda comprar sin riesgo inmediato y baja cuando no se necesita pedido. Esta regla permite auditar cada cantidad desde sus componentes.",
        "La formulación no optimiza costes de mantenimiento, faltantes o caducidad, aunque conserva la vida útil como dato. Tampoco modela fechas de llegada para cada pedido en tránsito ni múltiples proveedores. Se eligió porque el objetivo fue demostrar la integración entre predicción y decisión con parámetros básicos. En una fase posterior podría reemplazarse por una política (s, S) o por optimización estocástica sin alterar los módulos de ingestión y presentación.",
    ])
    formula_rows = [
        ["D_c", "Suma del forecast durante lead_time + review_period", "Unidades"],
        ["SS", "Demanda diaria media × safety_stock_days", "Unidades"],
        ["S_disp", "stock_on_hand + on_order", "Unidades disponibles o en tránsito"],
        ["Q_neta", "min(max(0, D_c + SS - S_disp), max(0, max_stock - S_disp))", "Necesidad limitada por capacidad"],
        ["Q_sugerida", "ceil(Q_neta / pack_size) × pack_size", "Unidades por empaque"],
    ]
    cursor = add_table(doc, cursor, "Cálculo de la recomendación de reposición", ["Símbolo", "Cálculo", "Interpretación"], formula_rows, [1.0, 3.6, 2.1])
    cursor = add_heading(doc, cursor, "Interfaz Streamlit")
    cursor = add_paragraphs(doc, cursor, [
        "La aplicación app/streamlit_app.py presenta controles para horizonte, días de seguridad, periodo de revisión y ventana auxiliar de evaluación. La pestaña Predicción permite seleccionar productos y consultar su trayectoria futura; Reposición muestra stock, pedidos en camino, cobertura, cantidad y prioridad; Evaluación presenta las métricas; y Datos permite inspeccionar ventas e inventario. El usuario puede descargar las recomendaciones en CSV.",
        "La interfaz se mantuvo deliberadamente operativa y compacta. Los encabezados y ayudas contextuales están en español, mientras los algoritmos conservan los nombres técnicos Random Forest, Ridge Regression y Seasonal Naive (7 días). No se añadieron edición de ventas ni autenticación, pues el caso de uso es una demostración local. Los cambios de parámetros de inventario recalculan la recomendación sin reentrenar el modelo.",
    ])
    cursor = add_action_note(doc, cursor, "DOS CAPTURAS DEL PANEL STREAMLIT", "Ejecuta `streamlit run app/streamlit_app.py` y selecciona 14 días. Inserta una captura de la pestaña `Predicción` con P003 seleccionado y el gráfico visible, y otra de `Reposición` con la tabla de stock, cobertura, cantidad y prioridad. Usa los pies `Figura X. Predicción de demanda en el panel Streamlit` y `Figura X. Recomendaciones de reposición en el panel Streamlit`; conserva `Fuente: elaboración propia`. No incluyas pantallas en carga ni rutas personales.")
    cursor = add_heading(doc, cursor, "API y persistencia local")
    cursor = add_paragraphs(doc, cursor, [
        "FastAPI expone GET /health, POST /sample-data, POST /train y GET /predict. El endpoint de salud confirma que el proceso responde; sample-data regenera los CSV sintéticos; train acepta la ventana de evaluación; y predict recibe horizonte, días de seguridad y periodo de revisión y devuelve forecast y recomendaciones. Los tipos y límites de los parámetros se validan mediante FastAPI. La documentación OpenAPI se genera automáticamente en /docs.",
        "El script predict.py utiliza el módulo storage.py y sqlite3 para escribir tres tablas: forecast, recommendations y model_metrics. Cada ejecución reemplaza el contenido anterior, por lo que la base representa el snapshot más reciente y no un historial de corridas. En el prototipo no existe migración de esquema, marca temporal ni gestión de usuarios; estas capacidades serían necesarias antes de un uso comercial.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "captura_api.png", "Documentación interactiva de la API FastAPI en ejecución")
    cursor = add_heading(doc, cursor, "Automatización con n8n y despliegue con Docker")
    cursor = add_paragraphs(doc, cursor, [
        "El archivo n8n/workflows/retrain_and_report.json define tres nodos: Daily schedule, que programa la ejecución diaria a las 07:00; Train model via API, que envía POST a /train; y Generate recommendations via API, que consulta GET /predict para catorce días, tres días de seguridad y un día de revisión. n8n se comunica con api-prediccion mediante la red interna de Docker Compose. El workflow se importó correctamente en n8n 2.30.5 y quedó versionado como JSON inspeccionable (n8n, s. f.).",
        "Docker Compose agrupa los servicios bajo tfe-mini-market-prediccion-reposicion. La API y el panel reutilizan la imagen Python del proyecto; n8n utiliza la imagen oficial n8nio/n8n:2.30.5. Los volúmenes montan data, models y reports en los servicios Python y conservan la configuración de n8n. Los puertos se publican únicamente en el equipo local, de acuerdo con el carácter demostrativo del despliegue (Docker, s. f.).",
        "La composición se inició y los tres servicios respondieron correctamente: Streamlit en /_stcore/health, FastAPI en /health y n8n en /healthz. Además, POST /train?test_days=30 devolvió las métricas de los tres candidatos y GET /predict?horizon_days=14&safety_stock_days=3&review_period_days=1 devolvió 140 predicciones y diez recomendaciones. La captura del workflow ejecutado desde el botón Test workflow sigue pendiente porque debe mostrar la sesión interactiva del autor.",
    ])
    service_rows = [
        ["tfe_mini_market_panel_streamlit", "Panel Streamlit", "8501", "/_stcore/health", "HTTP correcto"],
        ["tfe_mini_market_api_prediccion", "API FastAPI", "8000", "/health", "HTTP correcto"],
        ["tfe_mini_market_automatizacion_n8n", "Orquestación n8n 2.30.5", "5678", "/healthz", "HTTP correcto"],
    ]
    cursor = add_table(doc, cursor, "Servicios comprobados mediante Docker Compose", ["Contenedor", "Componente", "Puerto", "Comprobación", "Resultado"], service_rows, [2.35, 1.5, 0.55, 1.25, 1.1])
    cursor = add_action_note(doc, cursor, "CAPTURA DEL WORKFLOW N8N", "Abre http://localhost:5678, entra al workflow ya importado y pulsa `Test workflow`. Espera a que finalicen los tres nodos. Captura el lienzo completo con Daily schedule, Train model via API y Generate recommendations via API marcados como correctos y, si cabe, el panel de ejecución. Sustituye este recuadro por la imagen y usa el pie `Figura X. Ejecución satisfactoria del workflow de reentrenamiento y recomendación en n8n`.")
    cursor = add_action_note(doc, cursor, "CAPTURA DE DOCKER", "Ejecuta `docker compose up --build -d`, abre Docker Desktop y despliega el grupo `tfe-mini-market-prediccion-reposicion`. Captura los contenedores `tfe_mini_market_panel_streamlit`, `tfe_mini_market_api_prediccion` y `tfe_mini_market_automatizacion_n8n` en ejecución. Oculta rutas, correos o datos personales. Sustituye este recuadro por la imagen y usa el pie `Figura X. Servicios del prototipo ejecutados mediante Docker Compose`.")
    cursor = add_heading(doc, cursor, "Pruebas automatizadas y manejo de errores")
    cursor = add_paragraphs(doc, cursor, [
        "La suite incluye pruebas de métricas, tratamiento de demanda real igual a cero, redondeo por empaque, generación de una recomendación, ejecución integral del pipeline y cardinalidad multihorizonte. Las seis pruebas se ejecutaron con pytest y finalizaron correctamente. La prueba multihorizonte verifica la regla productos × modelos × días; la ejecución completa de evaluación generó 70, 140 y 300 predicciones por modelo para 7, 14 y 30 días, respectivamente.",
        "El manejo de errores se concentra en los límites del sistema: columnas ausentes, fechas inválidas, modelos no entrenados, horizontes fuera de rango y archivos inexistentes. Los scripts devuelven códigos de salida no nulos y mensajes identificables por n8n. Esta estrategia simplifica el diagnóstico y evita que una ejecución incompleta sea registrada como satisfactoria.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "pruebas_automatizadas.png", "Ejecución satisfactoria de la suite automatizada")
    cursor = add_heading(doc, cursor, "Estructura del repositorio y responsabilidades")
    module_rows = [
        ["src/demand_forecasting/data.py", "Carga, validación, calendario y variables"],
        ["src/demand_forecasting/models.py", "Pipelines Ridge y Random Forest"],
        ["src/demand_forecasting/evaluation.py", "Backtesting multihorizonte y métricas"],
        ["src/demand_forecasting/replenishment.py", "Regla de cantidad y prioridad"],
        ["src/demand_forecasting/pipeline.py", "Entrenamiento y predicción operativa"],
        ["app/streamlit_app.py", "Panel de consulta"],
        ["app/api.py", "API REST y documentación OpenAPI"],
        ["scripts/*.py", "Entrenamiento, evaluación, predicción y evidencias"],
        ["tests/", "Pruebas unitarias y de integración ligera"],
    ]
    cursor = add_table(doc, cursor, "Módulos principales y responsabilidad", ["Ruta", "Responsabilidad"], module_rows, [3.3, 3.4])
    cursor = add_paragraphs(doc, cursor, [
        "La organización evita concentrar toda la lógica en la interfaz. Streamlit y FastAPI consumen funciones de dominio; los scripts actúan como puntos de entrada para personas, Docker y n8n. Esta separación hizo posible probar la preparación de datos y la reposición sin iniciar servidores.",
        "La implementación utiliza pandas y NumPy para manipulación, scikit-learn para modelado, joblib para serialización, FastAPI y Streamlit para presentación, sqlite3 de la biblioteca estándar para persistencia y pytest para pruebas. Todas son herramientas abiertas y ampliamente documentadas. Docker encapsula el entorno y n8n aporta orquestación visual, en línea con el objetivo de un MLOps ligero.",
    ])
    cursor = add_heading(doc, cursor, "Contratos de ejecución y control de fallos")
    contract_rows = [
        ["Carga de ventas", "CSV con seis columnas obligatorias", "Serie diaria validada", "ValueError y proceso detenido"],
        ["Carga de inventario", "Identidad, stock, tránsito, lead time y empaque", "Inventario tipado", "ValueError antes de recomendar"],
        ["POST /train", "test_days entre 7 y 120", "Modelo, ranking y métricas", "HTTP 422 si el parámetro no es válido"],
        ["GET /predict", "horizonte 1-60; seguridad 0-30; revisión 1-14", "Forecast y recomendaciones", "HTTP 422 para valores fuera de rango"],
        ["Modelo ausente", "Datos válidos disponibles", "Entrenamiento previo automático", "El error se propaga si no puede entrenarse"],
        ["Persistencia", "Forecast y recomendaciones completos", "CSV y snapshot SQLite", "No se registra una ejecución como satisfactoria"],
    ]
    cursor = add_table(
        doc,
        cursor,
        "Precondiciones, salidas y fallos controlados",
        ["Operación", "Precondición", "Salida correcta", "Respuesta ante fallo"],
        contract_rows,
        [1.2, 2.15, 1.8, 2.0],
    )
    cursor = add_paragraphs(doc, cursor, [
        "Los contratos hacen explícito dónde termina cada responsabilidad. data.py valida esquemas y tipos antes de que features.py construya retardos; FastAPI limita los parámetros numéricos antes de invocar el pipeline; pipeline.py coordina los artefactos; y storage.py escribe únicamente resultados ya calculados. Esta secuencia evita que la interfaz replique reglas de dominio o que n8n necesite conocer la estructura interna del modelo.",
        "La API adopta dos recuperaciones automáticas propias de una demostración. Si faltan los CSV, genera el conjunto sintético; si falta el modelo, ejecuta entrenamiento antes de predecir. Estas decisiones facilitan una exposición reproducible, pero serían peligrosas en producción porque podrían ocultar una ruta mal configurada o consumir recursos sin aprobación. En un piloto real deberían deshabilitarse, devolver un error explícito y requerir una acción autorizada.",
        "No se implementaron autenticación, autorización, cifrado de secretos ni limitación de peticiones. El servicio se publica solo en localhost y la red de Docker se utiliza para la comunicación interna. Esta frontera es suficiente para la demostración académica, pero no para exponer la API a Internet ni para procesar datos identificables. La restricción queda vinculada con RNF-04 y RNF-06 y se retoma en las limitaciones.",
        "El control de fallos se verificó principalmente en los límites de datos, parámetros y artefactos. No se probaron fallos de disco, pérdida de energía ni concurrencia de dos ejecuciones. SQLite reemplaza las tablas del snapshot y no conserva transacciones de negocio entre corridas; por ello, la recuperación operativa depende de los CSV generados y de volver a ejecutar el pipeline. Esta es una decisión consciente de alcance y no una propiedad de resiliencia empresarial.",
    ])
    cursor = add_heading(doc, cursor, "Secuencia operativa y trazabilidad de una ejecución")
    cursor = add_paragraphs(doc, cursor, [
        "Una ejecución completa comienza con la validación de columnas y tipos de los archivos de entrada. Si el esquema no coincide o no quedan filas suficientes después de construir las variables, el proceso se detiene antes de guardar un nuevo modelo. Esta comprobación es relevante en un flujo automatizado: un archivo vacío o una columna renombrada no debe promover silenciosamente un artefacto inválido.",
        "Superada la validación, train.py construye las variables, calcula el ranking multihorizonte en memoria, ajusta el ganador con todas las filas utilizables y guarda demand_model.joblib, metrics.csv y model_ranking.csv. evaluate.py ejecuta después la evaluación detallada y escribe metrics_by_horizon.csv, metrics_by_product.csv y backtest_predictions.csv. Esta separación distingue el artefacto operativo de las evidencias de análisis; el archivo serializado no se utiliza por sí solo como prueba de superioridad.",
        "La predicción futura carga el modelo y sus metadatos, verifica la compatibilidad de variables y construye las fechas a partir del último día observado. Para cada iteración se conservan product_id, fecha y valor pronosticado. La salida diaria se guarda antes de agregar inventario. Si la recomendación pareciera incorrecta, el analista puede comprobar el forecast de ese producto, los días cubiertos, el stock inicial y el tamaño de empaque en vez de tratar la cantidad final como una caja negra.",
        "La persistencia conserva en SQLite el último forecast, las últimas recomendaciones y las métricas de ajuste. Las tablas se reemplazan en cada ejecución de predict.py; por ello permiten consultar un snapshot estructurado, pero no distinguir dos corridas ni auditar evolución histórica. Una ampliación deberá añadir execution_id, fecha, parámetros y escritura acumulativa. No se implementó promoción automática basada en un umbral porque una mejora métrica pequeña no garantiza que el nuevo modelo sea apropiado para operación.",
        "Los errores se manejan según su etapa. Un error de datos bloquea entrenamiento; un error al cargar el modelo bloquea forecast; un fallo al persistir conserva los CSV y devuelve un estado no satisfactorio; y un fallo de interfaz no modifica artefactos. Esta separación reduce el riesgo de resultados parciales. En un despliegue posterior deberían añadirse transacciones para escritura conjunta, logs estructurados y una tabla de estados de ejecución.",
        "La trazabilidad también abarca la memoria. Las tablas de resultados se construyeron a partir de los CSV finales y las figuras utilizan esos mismos archivos. No se redondearon datos antes de calcular el ranking; el redondeo a dos decimales se aplica únicamente para presentación. Las cantidades de reposición sí se redondean en el dominio, porque deben respetar el tamaño de empaque. Esta distinción evita confundir formato editorial con una regla de negocio.",
    ])
    return cursor


def build_evaluation(doc: Document, sales: pd.DataFrame, metrics_simple: pd.DataFrame, metrics_h: pd.DataFrame, metrics_p: pd.DataFrame, ranking: pd.DataFrame, recommendations: pd.DataFrame, sensitivity: pd.DataFrame):
    cursor = clear_section(doc, "Evaluación de la solución")
    display = {"ridge": "Ridge Regression", "random_forest": "Random Forest", "seasonal_naive_7d": "Seasonal Naive (7 días)"}
    cursor = add_heading(doc, cursor, "Diseño de la evaluación")
    cursor = add_paragraphs(doc, cursor, [
        "La evaluación respondió a tres preguntas: si los modelos de aprendizaje automático superaban una referencia estacional, si el rendimiento se mantenía al ampliar el horizonte y si la salida podía transformarse en recomendaciones coherentes. Para contestarlas se combinaron evaluación predictiva, verificación funcional y análisis de sensibilidad de inventario. No se mezclaron estas dimensiones: una métrica de forecast no demuestra usabilidad, y una interfaz funcional no demuestra precisión.",
        "El backtesting utilizó el final de la serie como periodo de prueba. Para el horizonte de siete días se entrenó con todas las fechas anteriores a los últimos siete días y se pronosticó ese bloque; el mismo procedimiento se repitió con catorce y treinta días. En todos los casos participaron diez productos. Los modelos recibieron exactamente los mismos cortes para garantizar comparabilidad.",
        "La predicción fue diaria y recursiva. Por tanto, el número de registros esperado por modelo fue el producto del horizonte y el número de referencias. La ejecución generó 70 observaciones para 7 días, 140 para 14 días y 300 para 30 días. El archivo reports/backtest_predictions.csv contiene 1.530 filas en total: 510 por cada uno de los tres modelos. Esta comprobación corrige la inconsistencia señalada en el predepósito.",
        "La evaluación operativa separó el backtesting de la predicción futura. Después de seleccionar Random Forest, predict.py volvió a ajustar el modelo con la historia disponible y generó catorce fechas futuras para cada producto. El archivo forecast.csv contiene 140 filas y recommendations.csv contiene diez filas agregadas, una por producto. Una recomendación no equivale a una fila diaria: resume la demanda dentro del periodo de cobertura y el estado de inventario.",
    ])
    cardinality_rows = [
        [7, 10, 70, 70, "Correcto"],
        [14, 10, 140, 140, "Correcto"],
        [30, 10, 300, 300, "Correcto"],
    ]
    cursor = add_table(doc, cursor, "Comprobación de cardinalidad por horizonte y modelo", ["Horizonte", "Productos", "Esperadas", "Obtenidas", "Resultado"], cardinality_rows, [1.15, 1.1, 1.2, 1.2, 1.25])
    cursor = add_heading(doc, cursor, "Métricas de error")
    cursor = add_paragraphs(doc, cursor, [
        "MAE expresa el error absoluto medio en unidades vendidas. Su interpretación directa permite indicar cuántas unidades se desvía el pronóstico en un producto-día típico. RMSE eleva los errores al cuadrado antes de promediarlos y, por tanto, penaliza con mayor intensidad las desviaciones grandes. Ambas métricas conservan la escala del objetivo, aunque RMSE no representa una cantidad lineal de error.",
        "MAPE divide el error absoluto por la demanda observada y lo expresa como porcentaje. Cuando el valor real es cero o muy pequeño, el cociente se vuelve indefinido o exagerado. La implementación excluye las observaciones cuya demanda real es cero, pero el resultado sigue siendo sensible en productos de bajo volumen. Por este motivo no se utilizó MAPE como único criterio.",
        "SMAPE normaliza el error por la suma de los valores absolutos observado y pronosticado. Se definió como 200/n por la suma de |y - ŷ|/(|y| + |ŷ|). Su rango facilita comparar productos con escalas distintas. Hyndman y Koehler (2006) advierten que ninguna métrica porcentual es universal; en este trabajo SMAPE se complementó con MAE, RMSE y resultados desagregados.",
        "El criterio de selección fue el promedio aritmético del SMAPE obtenido en 7, 14 y 30 días. Cada horizonte tuvo el mismo peso para reflejar el compromiso explícito del objetivo. Una ponderación por número de observaciones habría dado más influencia al horizonte de treinta días y habría respondido a una pregunta distinta.",
    ])
    cursor = add_heading(doc, cursor, "Evaluación auxiliar de la ventana final")
    simple_rows = [
        [
            display[row.model],
            f"{row.mae:.2f}",
            f"{row.rmse:.2f}",
            f"{row.mape:.2f} %",
            f"{row.smape:.2f} %",
        ]
        for row in metrics_simple.itertuples(index=False)
    ]
    cursor = add_table(doc, cursor, "Métricas auxiliares del endpoint de entrenamiento", ["Modelo", "MAE", "RMSE", "MAPE", "SMAPE"], simple_rows, [1.85, 1.0, 1.0, 1.1, 1.1])
    cursor = add_paragraphs(doc, cursor, [
        "El endpoint /train también produce metrics.csv con una comprobación auxiliar sobre los últimos treinta días del conjunto supervisado. En este cálculo, cada fila del tramo de prueba conserva retardos derivados de observaciones reales anteriores; por ello representa un escenario de actualización diaria a un paso y no un pronóstico recursivo de treinta días emitido desde un único origen.",
        "En esa comprobación Ridge Regression obtuvo SMAPE de 25,36 %, frente a 26,44 % de Random Forest. Esta tabla explica por qué el panel puede mostrar a Ridge como líder de la ventana auxiliar. No se utilizó para decidir el modelo operativo: la decisión se tomó con la evaluación recursiva de 7, 14 y 30 días descrita a continuación. Mantener ambas salidas permite distinguir dos preguntas experimentales en vez de mezclar sus cifras.",
    ])
    cursor = add_heading(doc, cursor, "Resultados por horizonte")
    metric_rows = []
    for _, row in metrics_h.sort_values(["horizon_days", "smape"]).iterrows():
        metric_rows.append([
            int(row.horizon_days), display[row.model], int(row.n_predictions), f"{row.mae:.2f}", f"{row.rmse:.2f}", f"{row.mape:.2f} %", f"{row.smape:.2f} %",
        ])
    cursor = add_table(doc, cursor, "Resultados predictivos por horizonte", ["Días", "Modelo", "n", "MAE", "RMSE", "MAPE", "SMAPE"], metric_rows, [0.55, 1.4, 0.55, 0.75, 0.75, 0.9, 0.9])
    cursor = add_paragraphs(doc, cursor, [
        "En siete días, Ridge obtuvo el menor MAE (6,19) y SMAPE (25,37 %). Random Forest quedó cerca, con MAE 6,45 y SMAPE 26,92 %. Ambos superaron al baseline, cuyo SMAPE fue 33,69 %. La diferencia indica que las variables de retardos, ventanas, calendario, precio y promoción aportaron información adicional frente a repetir exclusivamente la semana previa.",
        "En catorce días se mantuvo la ventaja de Ridge: MAE 6,02 y SMAPE 27,85 %, frente a 6,34 y 29,53 % de Random Forest. El baseline aumentó hasta 36,97 %. El deterioro respecto a siete días era esperable, porque la recursión incorpora predicciones previas en los retardos y propaga parte del error.",
        "En treinta días cambió el orden. Random Forest obtuvo MAE 6,15 y SMAPE 28,54 %, mientras Ridge alcanzó MAE 6,53 y SMAPE 40,43 %. El MAE de Ridge siguió siendo moderado, pero su error porcentual aumentó en observaciones de demanda baja. El bosque mantuvo un comportamiento más uniforme, lo que justificó su selección para un sistema que evalúa los tres horizontes.",
        "El baseline presentó MAE entre 7,53 y 8,62 y SMAPE entre 33,69 % y 36,97 %. No fue el mejor en ningún horizonte, pero estableció que la tarea contiene una señal semanal importante. La mejora del modelo operativo frente al SMAPE medio del baseline fue de aproximadamente 7,09 puntos porcentuales, equivalente a una reducción relativa cercana al 20 %.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "metricas_por_horizonte.png", "Comparación de MAE y SMAPE por modelo y horizonte")
    cursor = add_heading(doc, cursor, "Selección del modelo final")
    ranking_rows = []
    for _, row in ranking.sort_values("rank").iterrows():
        ranking_rows.append([int(row["rank"]), display[row.model], f"{row.mean_mae:.2f}", f"{row.mean_rmse:.2f}", f"{row.mean_mape:.2f} %", f"{row.mean_smape:.2f} %"])
    cursor = add_table(doc, cursor, "Ranking promedio multihorizonte", ["Puesto", "Modelo", "MAE medio", "RMSE medio", "MAPE medio", "SMAPE medio"], ranking_rows, [0.75, 1.5, 1.05, 1.1, 1.2, 1.2])
    cursor = add_paragraphs(doc, cursor, [
        "Random Forest fue seleccionado porque obtuvo el menor SMAPE medio, 28,33 %, y la mayor estabilidad al horizonte largo. Ridge alcanzó un MAE medio ligeramente menor, 6,25 frente a 6,31, pero su SMAPE medio fue 31,22 %. La decisión priorizó el criterio establecido antes del análisis y no la métrica que favorecía retrospectivamente a un modelo.",
        "La diferencia entre MAE y SMAPE ilustra por qué se reportaron varias métricas. Ridge cometió errores absolutos similares en unidades, pero concentró desviaciones relativas elevadas en días de menor demanda a treinta días. Random Forest evitó ese deterioro, aunque no fue superior en los horizontes cortos. En una aplicación real podría utilizarse selección por horizonte; el prototipo conserva un único modelo para simplificar operación y trazabilidad.",
        "La conclusión no es que Random Forest sea universalmente mejor. El resultado depende del conjunto sintético, la ventana temporal y los hiperparámetros. La evidencia solo permite afirmar que, bajo este protocolo reproducible, fue el candidato más equilibrado para los tres horizontes estudiados.",
    ])
    cursor = add_heading(doc, cursor, "Resultados por producto")
    rf_products = metrics_p[metrics_p.model == "random_forest"].groupby("product_id", as_index=False).agg(mae=("mae", "mean"), smape=("smape", "mean"), observations=("n_predictions", "sum"))
    product_rows = [[row.product_id, int(row.observations), f"{row.mae:.2f}", f"{row.smape:.2f} %"] for _, row in rf_products.sort_values("smape").iterrows()]
    cursor = add_table(doc, cursor, "Rendimiento medio de Random Forest por producto", ["Producto", "Observaciones", "MAE medio", "SMAPE medio"], product_rows, [1.25, 1.5, 1.4, 1.45])
    best = rf_products.sort_values("smape").iloc[0]
    worst = rf_products.sort_values("smape").iloc[-1]
    cursor = add_paragraphs(doc, cursor, [
        f"La desagregación mostró heterogeneidad entre referencias. {best.product_id} presentó el menor SMAPE medio ({best.smape:.2f} %), mientras {worst.product_id} alcanzó el mayor ({worst.smape:.2f} %). Reportar solo el promedio global habría ocultado esta diferencia y podría conducir a confiar por igual en recomendaciones con niveles de incertidumbre distintos.",
        "Los productos con demanda media alta tendieron a mostrar porcentajes más estables, porque una desviación de pocas unidades representa una fracción menor del valor real. En las referencias de bajo volumen, días cercanos a cero elevan MAPE y SMAPE. Para operación real convendría acompañar la predicción con intervalos y umbrales de confianza por producto.",
        "No se excluyeron productos difíciles para mejorar la cifra global. Cada uno aportó siete, catorce y treinta observaciones por modelo, es decir, 51 evaluaciones diarias acumuladas. Esta simetría facilita comparar referencias, aunque una cartera real exigiría ponderar por valor económico, criticidad o rotación.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "smape_por_producto.png", "SMAPE de Random Forest por producto y horizonte")
    cursor = add_heading(doc, cursor, "Resultados por categoría y sesgo")
    backtest = pd.read_csv(PROJECT / "reports" / "backtest_predictions.csv")
    rf_backtest = backtest[backtest.model_name == "random_forest"].copy()
    category_names = {
        "Aseo": "Aseo",
        "Basicos": "Básicos",
        "Bebidas": "Bebidas",
        "Despensa": "Despensa",
        "Frutas": "Frutas",
        "Lacteos": "Lácteos",
        "Panaderia": "Panadería",
    }
    category_rows = []
    for category, group in rf_backtest.groupby("category"):
        error = group.forecast_units - group.actual_units
        denominator = (group.forecast_units.abs() + group.actual_units.abs()) / 2
        valid = denominator != 0
        category_rows.append([
            category_names.get(category, category),
            int(group.product_id.nunique()),
            int(len(group)),
            f"{error.abs().mean():.2f}",
            f"{(error.pow(2).mean() ** 0.5):.2f}",
            f"{(error[valid].abs() / denominator[valid]).mean() * 100:.2f} %",
            f"{error.mean():.2f}",
        ])
    category_rows.sort(key=lambda row: float(row[5].replace(" %", "")))
    cursor = add_table(
        doc,
        cursor,
        "Rendimiento de Random Forest por categoría",
        ["Categoría", "Productos", "n", "MAE", "RMSE", "SMAPE", "Sesgo"],
        category_rows,
        [1.25, 0.9, 0.55, 0.75, 0.75, 0.95, 0.8],
    )
    cursor = add_paragraphs(doc, cursor, [
        "El análisis por categoría reutilizó las 510 predicciones de Random Forest: 51 observaciones por producto acumuladas entre los tres horizontes. Panadería obtuvo el menor SMAPE, 20,67 %, mientras Aseo alcanzó 38,61 %. Sin embargo, Panadería tuvo un MAE de 9,71 unidades y Aseo solo 3,21. Esta inversión confirma que un error porcentual alto no implica necesariamente una desviación grande en unidades: depende de la escala de demanda.",
        "El sesgo se calculó como pronóstico menos valor real. Un valor negativo indica subestimación media y uno positivo, sobreestimación. Bebidas presentó el sesgo negativo más pronunciado, -5,23 unidades por producto-día, seguido de Panadería con -3,68 y Básicos con -2,84. Lácteos y Aseo mostraron sesgos positivos pequeños. Esta lectura importa para inventario porque una subestimación persistente puede elevar el riesgo de rotura aunque el MAE global parezca aceptable.",
        "Las categorías no contienen el mismo número de productos: Lácteos, Bebidas y Despensa agrupan dos referencias, mientras las demás contienen una. Por ello, la tabla es descriptiva y no permite atribuir el rendimiento a la categoría como factor causal. Una validación real necesitaría más productos por familia, múltiples tiendas y periodos adicionales antes de afirmar diferencias estables entre categorías.",
        "El prototipo no corrige el sesgo después del entrenamiento. Hacerlo con la misma ventana de prueba introduciría optimismo. Una extensión podría estimar calibración con ventanas internas, mantener un residual por producto y supervisar cambios del sesgo en el tiempo. Para el alcance actual, la contribución consiste en hacerlo visible y vincularlo con la interpretación de la recomendación.",
    ])
    cursor = add_heading(doc, cursor, "Análisis del pronóstico operativo")
    cursor = add_paragraphs(doc, cursor, [
        "La ejecución operativa utilizó Random Forest reentrenado con todo el histórico disponible y produjo 140 filas para el horizonte de catorce días. Cada product_id aparece exactamente catorce veces y cada fecha contiene diez registros. La suma por producto alimentó el motor de inventario; el archivo diario se preservó para que el usuario pueda inspeccionar la trayectoria y no solo el total.",
        "La figura del producto de ejemplo superpone la historia reciente y los catorce días futuros. La separación vertical marca el comienzo del pronóstico. La curva futura conserva la escala y el patrón semanal sin valores negativos. No representa un intervalo de confianza; es una estimación puntual y debe leerse junto con el error histórico del producto.",
        "La promoción futura se configuró en cero. Si el comercio ya conoce campañas, ignorarlas puede subestimar demanda. En cambio, inventar promociones generaría una falsa precisión. El diseño deja esta variable explícita para que una versión posterior incorpore un calendario de acciones comerciales.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "forecast_producto_ejemplo.png", "Pronóstico operativo de catorce días para un producto de ejemplo")
    cursor = add_heading(doc, cursor, "Resultados de reposición")
    rec_rows = []
    for _, row in recommendations.sort_values(["priority", "suggested_order_qty"], ascending=[True, False]).iterrows():
        rec_rows.append([row.product_id, int(row.stock_on_hand), f"{row.demand_during_lead_time:.1f}", int(row.suggested_order_qty), row.priority])
    cursor = add_table(doc, cursor, "Recomendaciones generadas para el horizonte operativo", ["Producto", "Stock", "Demanda prevista", "Cantidad sugerida", "Prioridad"], rec_rows, [1.05, 0.8, 1.55, 1.55, 1.05])
    total_order = int(recommendations.suggested_order_qty.sum())
    high_count = int((recommendations.priority == "alta").sum())
    cursor = add_paragraphs(doc, cursor, [
        f"El escenario base produjo recomendaciones para los diez productos, con {total_order} unidades sugeridas en total y {high_count} referencias de prioridad alta. Las mayores cantidades correspondieron a productos cuya demanda durante la cobertura superó ampliamente el stock disponible. El redondeo a pack_size explica que la cantidad final pueda exceder ligeramente la necesidad neta.",
        "La prioridad no ordena únicamente por cantidad. Un producto de bajo volumen puede ser urgente si su cobertura es inferior al tiempo de reposición; otro de alto volumen puede tener prioridad media si dispone de stock suficiente para los primeros días. Esta separación mejora la interpretación operativa y evita equiparar volumen con riesgo de rotura.",
        "Las cantidades no constituyen órdenes automáticas. La interfaz las presenta como recomendaciones que deben revisarse ante promociones no registradas, pedidos en tránsito, caducidad, restricciones de caja o cambios del proveedor. Mantener a una persona en la decisión es coherente con el alcance de apoyo, no sustitución, de la herramienta.",
    ])
    cursor = add_figure(doc, cursor, FIGURES / "reposicion_sugerida.png", "Cantidad de reposición sugerida y prioridad por producto")
    cursor = add_heading(doc, cursor, "Análisis de sensibilidad de inventario")
    sensitivity_rows = []
    for _, row in sensitivity.iterrows():
        sensitivity_rows.append([int(row.safety_stock_days), int(row.review_period_days), int(row.products_with_order), int(row.high_priority_products), int(row.total_suggested_units), f"{row.mean_coverage_days:.2f}"])
    cursor = add_table(doc, cursor, "Sensibilidad a días de seguridad y revisión", ["Seguridad", "Revisión", "Productos con pedido", "Prioridad alta", "Unidades", "Cobertura media"], sensitivity_rows, [0.9, 0.9, 1.25, 1.15, 1.0, 1.15])
    cursor = add_paragraphs(doc, cursor, [
        "El análisis modificó safety_stock_days entre uno, tres y cinco días y review_period_days entre uno y tres días. Con un día de seguridad y uno de revisión se recomendaron 716 unidades para nueve productos. Al aumentar ambos valores a cinco y tres días, respectivamente, la recomendación alcanzó 1.984 unidades para los diez productos.",
        "El crecimiento fue monotónico: ampliar el periodo que debe cubrirse elevó la demanda acumulada y, por tanto, la orden. Los escenarios (seguridad 1, revisión 3) y (seguridad 3, revisión 1) produjeron el mismo total de 1.178 unidades porque, bajo la fórmula actual, ambos añaden dos días equivalentes de cobertura. Esta igualdad hace visible una simplificación del motor: no diferencia conceptualmente incertidumbre de entrega y frecuencia de revisión si ambas se expresan solo como días de demanda media.",
        "Las siete prioridades altas permanecieron constantes porque se calculan principalmente con stock actual y tiempo de reposición; los parámetros de seguridad y revisión modifican la cantidad, no necesariamente la urgencia inicial. Una política futura podría hacer que la prioridad considere también variabilidad pronosticada y nivel de servicio objetivo.",
        "El ejercicio confirma que la herramienta es ajustable, pero también muestra el coste de elegir parámetros conservadores. Pasar del escenario mínimo al máximo añadió 1.268 unidades, un incremento aproximado del 177 %. Sin información de costes y caducidad no puede afirmarse qué escenario es económicamente óptimo.",
    ])
    cursor = add_heading(doc, cursor, "Verificación funcional y reproducibilidad")
    test_rows = [
        ["Prueba de generación", "Esquema, fechas y productos del dataset", "Superada"],
        ["Prueba de variables", "Retardos y ventanas sin fuga temporal", "Superada"],
        ["Prueba de métricas", "MAE, RMSE, MAPE y SMAPE con casos controlados", "Superada"],
        ["Prueba de reposición", "Truncamiento, seguridad y redondeo por empaque", "Superada"],
        ["Prueba multihorizonte", "70, 140 y 300 filas por modelo", "Superada"],
        ["Prueba de artefactos", "Creación de modelo y reportes", "Superada"],
    ]
    cursor = add_table(doc, cursor, "Resumen de pruebas automatizadas", ["Prueba", "Criterio", "Resultado"], test_rows, [1.65, 3.65, 1.15])
    performance_rows = [
        ["Entrenamiento", "scripts/train.py", "43,64", "< 60", "Cumple"],
        ["Evaluación multihorizonte", "scripts/evaluate.py", "30,09", "< 60", "Cumple"],
        ["Predicción operativa", "scripts/predict.py --horizon-days 14", "8,52", "< 60", "Cumple"],
    ]
    cursor = add_table(
        doc,
        cursor,
        "Tiempos de ejecución en el entorno local",
        ["Proceso", "Comando", "Segundos", "Criterio", "Estado"],
        performance_rows,
        [1.55, 2.9, 0.8, 0.75, 0.8],
    )
    cursor = add_paragraphs(doc, cursor, [
        "La ejecución completa de pytest produjo seis pruebas superadas. Además de las aserciones unitarias, se ejecutaron train.py, evaluate.py y predict.py desde la línea de comandos. Se verificó la existencia y el contenido de demand_model.joblib, forecast.csv, recommendations.csv, metrics_by_horizon.csv, metrics_by_product.csv y model_ranking.csv.",
        "Los tiempos se midieron con Python 3.13.4 en el entorno virtual del proyecto y con cada comando ejecutado de forma independiente en el equipo Windows del autor. Los tres procesos cumplieron el límite local de 60 segundos. Estas cifras sirven para verificar viabilidad académica y no constituyen un benchmark generalizable a otro hardware o a catálogos de mayor tamaño.",
        "La ejecución conjunta con Docker Compose inició el panel Streamlit, la API y n8n. Los tres servicios respondieron correctamente en sus comprobaciones HTTP: Streamlit en /_stcore/health, FastAPI en /health y n8n en /healthz. El workflow está versionado y configurado para invocar la API dentro de la red de Docker; su captura y ejecución visual permanecen como evidencia manual pendiente del autor.",
        "La reproducción requiere crear el entorno, instalar requirements.txt, generar o verificar los datos, entrenar, evaluar y predecir. Los comandos exactos se incorporan en el anexo. Los CSV de resultados permiten comprobar las cifras sin depender de la memoria ni de capturas estáticas.",
    ])
    cursor = add_heading(doc, cursor, "Matriz de verificación de requisitos")
    req_rows = [
        ["RF-01", "Ingestar ventas", "data.py y sales_sample.csv", "Verificado"],
        ["RF-02", "Validar esquema y calidad", "data.py, reporte de calidad y prueba de generación", "Verificado"],
        ["RF-03", "Generar variables temporales", "features.py y prueba de variables sin fuga", "Verificado"],
        ["RF-04", "Entrenar baseline y candidatos", "train.py, evaluate.py y métricas por modelo", "Verificado"],
        ["RF-05", "Seleccionar el modelo", "model_ranking.csv y metadatos del modelo", "Verificado"],
        ["RF-06", "Predecir 7/14/30 días", "70/140/300 filas por modelo y selector del panel", "Verificado"],
        ["RF-07", "Modificar seguridad y revisión", "controles Streamlit y análisis de sensibilidad", "Verificado"],
        ["RF-08", "Calcular la reposición", "recommendations.csv, fórmula documentada y pruebas", "Verificado"],
        ["RF-09", "Persistir resultados", "storage.py y SQLite con tres tablas reemplazables", "Verificado"],
        ["RF-10", "Automatizar con n8n", "workflow versionado y servicios Docker; captura pendiente", "Implementado"],
        ["RNF-01", "Reproducibilidad", "requirements, Dockerfile, Compose y scripts", "Verificado"],
        ["RNF-02", "Usabilidad", "Protocolo y SUS preparados; sin participantes reales", "Pendiente"],
        ["RNF-03", "Explicabilidad", "fórmula, ayudas contextuales y variables visibles", "Verificado"],
        ["RNF-04", "Ejecución local", "tres servicios en localhost con respuesta HTTP", "Verificado"],
        ["RNF-05", "Mantenibilidad", "módulos separados, estructura y pruebas", "Verificado"],
        ["RNF-06", "Privacidad", "datos sintéticos y ausencia de datos personales", "Verificado"],
        ["RNF-07", "Rendimiento razonable", "43,64/30,09/8,52 s; todos bajo 60 s", "Verificado"],
    ]
    cursor = add_table(doc, cursor, "Matriz de verificación con evidencias", ["Id.", "Requisito", "Evidencia", "Estado"], req_rows, [0.8, 1.65, 3.1, 1.25])
    cursor = add_heading(doc, cursor, "Discusión de validez")
    cursor = add_paragraphs(doc, cursor, [
        "La validez interna se reforzó mediante cortes temporales, igualdad de muestras y prevención de fuga. Sin embargo, solo se utilizó un origen de evaluación por horizonte. Un backtesting con múltiples orígenes proporcionaría una estimación menos dependiente del último mes y permitiría calcular dispersión entre ventanas.",
        "La validez de constructo está limitada por las métricas puntuales. MAE y SMAPE miden error, pero no directamente roturas de stock, desperdicio ni coste. El análisis de reposición demuestra coherencia matemática, no ahorro económico. Para evaluar impacto real serían necesarios costes unitarios, nivel de servicio, ventas perdidas y caducidad.",
        "La validez externa es la principal limitación. Los datos son sintéticos y representan diez productos de una sola tienda ficticia. No se puede generalizar el rendimiento a un comercio real, a promociones complejas, festivos locales o demanda intermitente. El valor del resultado reside en demostrar el proceso reproducible y detectar qué debe validarse con datos reales.",
        "La confiabilidad técnica se apoyó en semilla fija, parámetros explícitos, pruebas y artefactos versionados. Aun así, los resultados pueden variar con versiones futuras de bibliotecas o hardware. Docker reduce esa variación, pero no elimina diferencias de arquitectura ni garantiza determinismo perfecto en todas las operaciones paralelas.",
        "La evaluación con usuarios expertos prevista inicialmente no se ejecutó. No se reportaron respuestas inventadas ni una puntuación SUS sin participantes. En el anexo se conserva el instrumento y el protocolo para una validación posterior. Por esta razón, el objetivo de usabilidad permanece cumplido solo a nivel de diseño funcional.",
    ])
    cursor = add_heading(doc, cursor, "Comparación con el estado del arte")
    cursor = add_paragraphs(doc, cursor, [
        "Los resultados coinciden con la literatura en que no existe un único modelo dominante para todas las series y horizontes. Ridge fue competitivo en el corto plazo y Random Forest mostró mayor robustez a treinta días. Esta observación es consistente con el énfasis de Hyndman y Athanasopoulos (2021) en evaluar contra baselines y mediante separación temporal, en lugar de asumir superioridad por complejidad.",
        "El prototipo se diferencia de plataformas comerciales por su transparencia y alcance reducido. Cada transformación, parámetro y métrica puede inspeccionarse; no se ofrecen integraciones empresariales, optimización avanzada ni soporte operativo. Esa simplicidad es adecuada para un TFE y para pequeños comercios que requieren entender la recomendación, pero no sustituye una solución industrial.",
        "La incorporación de n8n y Docker aproxima el trabajo a prácticas MLOps: automatiza pasos repetibles, conserva artefactos y desacopla servicios. No se implementaron monitoreo de deriva, promoción automática entre ambientes ni rollback. Por ello se emplea la expresión MLOps ligero y no una plataforma completa de operación de modelos.",
    ])
    cursor = add_heading(doc, cursor, "Interpretación operativa de los errores")
    cursor = add_paragraphs(doc, cursor, [
        "Un MAE cercano a seis unidades por producto-día no determina por sí solo si el sistema es aceptable. Para una referencia que vende cuarenta unidades diarias puede ser tolerable; para otra que vende cinco puede alterar completamente la decisión. La tabla por producto y SMAPE complementa la lectura, pero una validación real debería convertir los errores en faltantes, excedentes y coste. El prototipo evita fijar un umbral universal sin esos datos económicos.",
        "Los errores positivos y negativos tienen consecuencias diferentes. Sobreestimar demanda aumenta inventario y puede elevar merma; subestimarla favorece roturas de stock y ventas perdidas. MAE, RMSE y SMAPE son simétricas y no reflejan esa asimetría. Una ampliación podría calcular bias medio por producto y una función de pérdida ponderada según margen, perecibilidad y nivel de servicio.",
        "El horizonte modifica la forma de usar la salida. El pronóstico de siete días puede apoyar una revisión frecuente y reaccionar con rapidez a cambios recientes. El de treinta días sirve para planificación, pero acumula incertidumbre y depende más de supuestos sobre promociones y disponibilidad. Por ello la interfaz debe mostrar el horizonte junto a la recomendación y no presentar las cantidades como equivalentes entre escenarios.",
        "Random Forest fue estable en promedio, pero no garantiza coherencia temporal fuera del rango observado. Los árboles interpolan patrones aprendidos y no extrapolan una tendencia de forma natural. Si el comercio creciera de manera sostenida, Ridge o un modelo temporal con tendencia podría responder de otra forma. El monitoreo debería comparar demanda real y prevista por fecha, producto y horizonte para identificar esa pérdida de adecuación.",
        "La variabilidad entre productos sugiere una política gradual de adopción. Las referencias con error bajo podrían usarse primero como apoyo rutinario, mientras las de error alto permanecerían bajo revisión estrecha. Este planteamiento es más responsable que activar órdenes automáticas para toda la cartera. También permitiría recopilar evidencia de utilidad sin asumir desde el inicio que todos los productos comparten el mismo patrón.",
        "El análisis de sensibilidad mostró que el parámetro de cobertura puede tener un efecto mayor sobre las unidades sugeridas que pequeñas diferencias entre modelos. En el escenario extremo se añadieron más de mil unidades respecto al mínimo. Esto implica que la capacitación del usuario y la configuración de inventario son tan importantes como el algoritmo. Una interfaz correcta debe mostrar los parámetros activos y registrar sus cambios.",
        "Finalmente, las métricas se calcularon sobre demanda observada, que en un comercio real puede estar censurada por falta de stock. Si un producto se agota, las ventas registradas subestiman la demanda. Entrenar sin identificar stockouts enseñaría al modelo a reproducir esa limitación. La futura incorporación de movimientos de inventario y disponibilidad permitiría distinguir cero demanda de cero venta por ausencia de producto.",
    ])
    return cursor


def build_conclusions(doc: Document):
    cursor = clear_section(doc, "Conclusiones y trabajo futuro")
    cursor = add_heading(doc, cursor, "Conclusiones generales")
    cursor = add_paragraphs(doc, cursor, [
        "El trabajo alcanzó su propósito principal: se desarrolló un prototipo reproducible que transforma ventas diarias en pronósticos y estos, a su vez, en recomendaciones explicables de reposición. La contribución no se limita a una propuesta arquitectónica; existe código ejecutable, datos, modelos serializados, reportes, pruebas, API, panel y archivos de automatización.",
        "La evaluación demostró que al menos un modelo de aprendizaje automático superó al baseline estacional en cada horizonte. Random Forest obtuvo el mejor SMAPE promedio, 28,33 %, y fue seleccionado como modelo operativo por su estabilidad a treinta días. Ridge fue mejor en siete y catorce días, pero quedó por detrás del baseline en SMAPE a treinta días; esto muestra que la selección depende del horizonte y que una única cifra agregada debe discutirse, no presentarse de forma aislada.",
        "La discrepancia de cardinalidad detectada en el predepósito quedó resuelta. Diez productos generan 70, 140 y 300 predicciones por modelo en 7, 14 y 30 días. El escenario operativo de catorce días produjo 140 filas diarias y diez recomendaciones agregadas. La distinción entre ambos niveles quedó documentada y protegida mediante una prueba automatizada.",
        "La regla de reposición conectó la IA con una decisión concreta. Al combinar demanda durante cobertura, stock de seguridad, existencias y tamaño de empaque, la salida puede explicarse y ajustarse. El análisis de sensibilidad mostró que parámetros más conservadores elevan sustancialmente el inventario sugerido, por lo que deben configurarse con información económica y de servicio.",
    ])
    cursor = add_heading(doc, cursor, "Cumplimiento de los objetivos")
    cursor = add_paragraphs(doc, cursor, [
        "El primer objetivo específico se cumplió mediante un dataset sintético reproducible de 7.310 filas y un pipeline de validación, calendario y variables temporales. El uso de datos sintéticos permitió ejecutar el sistema completo, aunque limita la generalización y debe reemplazarse por datos anonimizados en una validación externa.",
        "El segundo objetivo se cumplió con la comparación justa de baseline estacional, Ridge y Random Forest. La evaluación por horizonte y producto aportó más información que la tabla agregada inicial. Los artefactos CSV permiten reproducir la selección sin copiar manualmente las cifras de la memoria.",
        "El tercer objetivo se cumplió con una regla configurable y sometida a pruebas. La recomendación muestra demanda prevista, stock, cantidad redondeada y prioridad. No se afirmó optimización económica porque no se modelaron costes ni restricciones de caducidad.",
        "El cuarto objetivo se cumplió técnicamente mediante Streamlit, FastAPI y SQLite. La API y el panel se ejecutaron correctamente y los resultados se persistieron. Quedan por incorporar al documento las capturas finales del panel, n8n y Docker obtenidas por el autor con los servicios visibles, señaladas en amarillo para evitar que se omitan antes de entregar.",
        "El quinto objetivo se cumplió en su componente cuantitativo y funcional, pero no en la validación con usuarios expertos. Se ejecutaron métricas y seis pruebas automatizadas; la encuesta de usabilidad no se aplicó. Esta situación se reconoce expresamente y el instrumento queda preparado como trabajo pendiente, sin fabricar resultados.",
    ])
    cursor = add_heading(doc, cursor, "Contribuciones del trabajo")
    cursor = add_paragraphs(doc, cursor, [
        "La primera contribución es un flujo de referencia de extremo a extremo para un comercio pequeño: datos, variables, modelos, evaluación temporal, reposición, consulta y automatización. Cada componente es reemplazable y utiliza herramientas abiertas, lo que facilita adaptación y aprendizaje.",
        "La segunda contribución es metodológica. El sistema distingue evaluación a un paso, backtesting multihorizonte y forecast operativo; explicita el número de observaciones esperado; y conserva resultados por producto. Esta trazabilidad reduce errores frecuentes en trabajos de forecasting, como mezclar periodos, elegir el modelo con la misma muestra de prueba o presentar un promedio sin desagregar.",
        "La tercera contribución es la conexión transparente entre predicción y reposición. El usuario puede seguir la cantidad desde el forecast hasta la cobertura y el redondeo por empaque. El enfoque favorece adopción responsable, porque una recomendación explicable puede cuestionarse y ajustarse.",
    ])
    cursor = add_heading(doc, cursor, "Limitaciones")
    limitations = [
        "Los datos son sintéticos y no reflejan toda la variabilidad, errores de captura, festivos ni sustitución entre productos de un comercio real.",
        "La evaluación emplea un único origen por horizonte; múltiples ventanas ofrecerían intervalos y mayor robustez estadística.",
        "Los pronósticos son puntuales y no incluyen intervalos de incertidumbre ni probabilidades de rotura de stock.",
        "La promoción futura se supone igual a cero y no se modelan precios o campañas planificadas.",
        "La reposición descuenta las unidades agregadas en tránsito y limita el pedido por stock máximo, pero no representa fechas de llegada por pedido, costes, presupuesto, caducidad activa, mínimos de proveedor ni restricciones conjuntas.",
        "No se completó la validación con usuarios expertos ni se dispone de una puntuación de usabilidad.",
        "Docker y n8n se ejecutaron y respondieron a sus comprobaciones de salud; sin embargo, aún falta incorporar la captura de una ejecución manual completa del workflow desde la interfaz de n8n.",
        "SQLite y el despliegue local no están diseñados para concurrencia elevada, múltiples tiendas ni requisitos de seguridad empresarial.",
    ]
    for item in limitations:
        cursor = add_list_item(doc, cursor, item)
    cursor = add_heading(doc, cursor, "Trabajo futuro")
    cursor = add_paragraphs(doc, cursor, [
        "La prioridad es sustituir o complementar los datos sintéticos con un histórico anonimizado de una tienda real. Antes de entrenar deberían documentarse calidad, días cerrados, cambios de catálogo, promociones y stockouts, porque una venta cero no siempre equivale a demanda cero. La evaluación debería repetirse con múltiples orígenes y separarse por familias de rotación.",
        "En modelado se propone incorporar modelos especializados de series temporales, intervalos conformales y selección por producto u horizonte. La comparación deberá mantener el baseline y evitar búsquedas de hiperparámetros sobre el conjunto final. También sería útil pronosticar demanda probabilística para vincular el stock de seguridad con un nivel de servicio explícito.",
        "En inventario se recomienda desagregar cada pedido en tránsito con su fecha prevista de llegada e incorporar costes de mantener y romper stock, vida útil, restricciones de proveedor y presupuesto. Con esas variables podría evaluarse no solo error predictivo, sino coste total y tasa de servicio mediante simulación histórica.",
        "En experiencia de usuario debe aplicarse el protocolo del anexo a tres o más participantes con experiencia en compras o inventario. Las tareas medirían éxito, tiempo y errores, seguidas de la escala SUS de Brooke (1996). Las observaciones cualitativas deberían traducirse en cambios concretos del panel.",
        "En operación se propone añadir monitorización de datos y rendimiento, historial de ejecuciones, registro de versiones del modelo y alertas cuando el error o la distribución cambien. Una base PostgreSQL, autenticación y gestión de secretos serían apropiadas para múltiples usuarios. Estas mejoras convertirían el prototipo en una prueba piloto más cercana a producción.",
    ])
    cursor = add_heading(doc, cursor, "Reflexión final")
    cursor = add_paragraphs(doc, cursor, [
        "El resultado muestra que una solución de IA útil para inventario no depende únicamente de elegir un algoritmo. La mayor parte del trabajo consistió en definir la unidad temporal, impedir fuga, conservar cardinalidad, comparar de forma justa y traducir el resultado a una decisión. La automatización y la interfaz aportan valor cuando preservan esa trazabilidad.",
        "Con el alcance controlado, el prototipo demuestra viabilidad técnica y ofrece una base válida para experimentación real. Sus límites están identificados y no se ocultan detrás de métricas agregadas. Esa combinación de evidencia, reproducibilidad y cautela permite valorar la contribución sin atribuirle una madurez operativa que todavía no posee.",
    ])
    return cursor


def build_annex(doc: Document):
    cursor = clear_section(doc, "Anexo A. Código fuente y datos analizados")
    cursor = add_heading(doc, cursor, "Contenido del repositorio")
    cursor = add_paragraphs(doc, cursor, [
        "El repositorio sistema-reposicion-mini-market contiene el código y los artefactos necesarios para reproducir la experimentación. Los directorios data/raw y data/processed separan entradas de transformaciones; models conserva el estimador serializado; reports reúne métricas, predicciones, recomendaciones y figuras; src contiene lógica de dominio; app expone interfaces; scripts ofrece comandos; tests verifica comportamiento; n8n almacena el workflow; y los archivos Docker describen el entorno.",
        "Los archivos generados de mayor importancia son reports/metrics_by_horizon.csv, reports/metrics_by_product.csv, reports/model_ranking.csv, reports/backtest_predictions.csv, reports/forecast.csv, reports/recommendations.csv y reports/replenishment_sensitivity.csv. La memoria toma sus cifras de estos artefactos. Si se modifica el dataset o el código, deben regenerarse antes de actualizar las tablas.",
    ])
    cursor = add_heading(doc, cursor, "Comandos de reproducción")
    commands = [
        "python -m venv .venv",
        ".venv\\Scripts\\activate",
        "python -m pip install -r requirements.txt",
        "python scripts/generate_sample_data.py",
        "python scripts/train.py",
        "python scripts/evaluate.py",
        "python scripts/predict.py --horizon-days 14",
        "python scripts/analyze_replenishment_sensitivity.py",
        "python scripts/generate_evaluation_figures.py",
        "python -m pytest -q",
        "uvicorn app.api:app --host 127.0.0.1 --port 8000",
        "streamlit run app/streamlit_app.py",
        "docker compose up --build",
    ]
    for command in commands:
        cursor = add_code(doc, cursor, command)
    cursor = add_paragraphs(doc, cursor, [
        "Los comandos se ejecutan desde la raíz del repositorio. Para una reproducción estricta se recomienda eliminar únicamente los artefactos generados de reports y models, conservar los CSV de entrada, instalar las versiones fijadas y ejecutar la secuencia completa. La semilla del dataset y random_state del bosque permanecen constantes.",
        "La API se consulta en http://127.0.0.1:8000/docs y Streamlit en http://127.0.0.1:8501. n8n utiliza el puerto definido en docker-compose.yml. Si un puerto está ocupado debe modificarse la asignación de host sin alterar el puerto interno del servicio.",
    ])
    cursor = add_heading(doc, cursor, "Endpoints de la API")
    endpoint_rows = [
        ["GET /health", "Comprueba disponibilidad", "Estado y modelo cargado"],
        ["POST /sample-data", "Regenera el dataset sintético", "Rutas de los CSV"],
        ["POST /train?test_days=30", "Entrena y selecciona el modelo", "Ruta del modelo y métricas"],
        ["GET /predict?horizon_days=14", "Genera forecast y reposición", "Predicciones y recomendaciones"],
        ["GET /docs", "Abre OpenAPI", "Interfaz interactiva"],
    ]
    cursor = add_table(doc, cursor, "Contrato resumido de la API", ["Ruta", "Propósito", "Respuesta"], endpoint_rows, [2.15, 2.5, 2.1])
    cursor = add_heading(doc, cursor, "Diccionario de artefactos de evaluación")
    artifact_rows = [
        ["metrics_by_horizon.csv", "horizonte-modelo", "n, MAE, RMSE, MAPE, SMAPE"],
        ["metrics_by_product.csv", "horizonte-modelo-producto", "métricas desagregadas"],
        ["model_ranking.csv", "modelo", "promedios y puesto"],
        ["backtest_predictions.csv", "modelo-producto-fecha", "real y predicción"],
        ["forecast.csv", "producto-fecha futura", "predicción operativa"],
        ["recommendations.csv", "producto", "stock, forecast, cantidad y prioridad"],
        ["replenishment_sensitivity.csv", "escenario", "parámetros y total sugerido"],
    ]
    cursor = add_table(doc, cursor, "Archivos de resultados reproducibles", ["Archivo", "Granularidad", "Contenido"], artifact_rows, [2.25, 2.15, 2.45])
    cursor = add_heading(doc, cursor, "Protocolo pendiente de validación con usuarios")
    cursor = add_paragraphs(doc, cursor, [
        "La validación se propone para tres a cinco participantes con experiencia en compras, inventario o administración de pequeños comercios. Cada sesión debe comenzar con consentimiento informado, explicación de que se evalúa la herramienta y no a la persona, y un escenario idéntico. No deben revelarse las cantidades correctas antes de completar las tareas.",
        "Las tareas son: identificar el producto con mayor urgencia; consultar el forecast de catorce días; modificar los días de seguridad y explicar el cambio; descargar recomendaciones; y localizar la métrica del modelo. Para cada tarea se registra éxito, tiempo, errores y comentarios. Después se aplica la escala SUS de diez ítems y tres preguntas abiertas sobre claridad, confianza y mejoras.",
        "El análisis debe reportar participantes de forma anónima, tasa de éxito por tarea, mediana de tiempo, problemas observados y puntuación SUS calculada según Brooke (1996). No debe afirmarse representatividad estadística con una muestra pequeña. El propósito es detectar problemas de interacción y utilidad percibida.",
    ])
    survey_items = [
        [1, "Creo que usaría esta herramienta con frecuencia"],
        [2, "La herramienta me pareció innecesariamente compleja"],
        [3, "La herramienta me pareció fácil de usar"],
        [4, "Necesitaría apoyo técnico para usarla"],
        [5, "Las funciones están bien integradas"],
        [6, "Encontré demasiada inconsistencia"],
        [7, "La mayoría aprendería a usarla rápidamente"],
        [8, "La herramienta me pareció incómoda"],
        [9, "Me sentí seguro al usarla"],
        [10, "Necesité aprender muchas cosas antes de empezar"],
    ]
    cursor = add_table(doc, cursor, "Instrumento SUS propuesto", ["Ítem", "Enunciado; escala 1 totalmente en desacuerdo a 5 totalmente de acuerdo"], survey_items, [0.75, 5.95])
    cursor = add_action_note(doc, cursor, "VALIDACIÓN CON USUARIOS", "Aplica el protocolo al menos a tres participantes reales antes de afirmar que el objetivo de usabilidad está cumplido. Añade una tabla anónima con rol, tareas superadas, tiempo y SUS; resume observaciones sin nombres ni datos personales. Si no puedes realizarla, conserva la limitación tal como está y no inventes resultados.")
    cursor = add_heading(doc, cursor, "Lista de evidencias antes de entrega")
    checklist = [
        "Confirmar la fecha oficial del depósito extraordinario y sustituir el aviso amarillo de la portada.",
        "Sustituir los recuadros amarillos por dos capturas de Streamlit, una de n8n y una de Docker, todas reales y con sus pies.",
        "Decidir si se aplicará la validación con usuarios. Si no se realiza, conservar explícitamente la limitación y no afirmar una puntuación SUS.",
        "No incluir rutas personales, tokens, correos ni información identificable en capturas o anexos.",
        "Después de insertar las imágenes, actualizar índices y numeración con Ctrl+A y F9, revisar el PDF completo y eliminar todos los recuadros amarillos antes del depósito.",
    ]
    for number, item in enumerate(checklist, start=1):
        cursor = add_body(doc, cursor, f"{number}. {item}")
        paragraph = paragraph_from_element(doc, cursor)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.left_indent = Inches(0.3)
        paragraph.paragraph_format.first_line_indent = Inches(-0.3)
    return cursor


def remove_uncited_references(doc: Document):
    uncited_prefixes = ("Doran, G. T.", "Virtanen, P.")
    for paragraph in list(doc.paragraphs):
        if paragraph.text.strip().startswith(uncited_prefixes):
            paragraph._element.getparent().remove(paragraph._element)


def repair_existing_headings(doc: Document):
    replacements = {
        "4.1.1. Priorización del alcance mínimo": "Priorización del alcance mínimo",
        "4.7. Trazabilidad y validación de requisitos": "Trazabilidad y validación de requisitos",
    }
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() in replacements:
            paragraph.text = replacements[paragraph.text.strip()]
            paragraph.style = "Heading 2"
        if paragraph.style.name == "Heading 1" and paragraph.text.startswith("Anexo A. Código fuente"):
            paragraph.text = "Anexo A. Código fuente y datos analizados"
            paragraph.style = "Heading 1"


def _add_static_index_entry(
    doc: Document,
    after,
    text: str,
    page_number: int | None,
    level: int = 1,
    emphasize: bool | None = None,
):
    element = insert_paragraph(doc, after, "", "Normal")
    paragraph = paragraph_from_element(doc, element)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.left_indent = Inches(0.18 * max(0, level - 1))
    paragraph.paragraph_format.first_line_indent = Inches(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(6.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
    )
    run = paragraph.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(9.5 if level == 1 else 9)
    if emphasize if emphasize is not None else level == 1:
        run.bold = True
    page_run = paragraph.add_run(f"\t{page_number if page_number is not None else '--'}")
    page_run.font.name = "Arial"
    page_run.font.size = run.font.size
    return element


def _add_page_break(doc: Document, after):
    element = insert_paragraph(doc, after, "", "Normal")
    paragraph_from_element(doc, element).add_run().add_break(WD_BREAK.PAGE)
    return element


def rebuild_indexes(doc: Document):
    heading_entries = []
    figure_entries = []
    table_entries = []
    intro_found = False
    counters = [0, 0, 0]
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Heading 1" and "Introducción" in paragraph.text:
            intro_found = True
        if not intro_found:
            continue
        if paragraph.style.name in {"Heading 1", "Heading 2", "Heading 3"} and paragraph.text.strip():
            level = int(paragraph.style.name[-1])
            counters[level - 1] += 1
            for index in range(level, 3):
                counters[index] = 0
            number = ".".join(str(value) for value in counters[:level])
            raw_text = paragraph.text.strip()
            heading_entries.append((f"{number}. {raw_text}", raw_text, level))
        elif paragraph.text.strip().startswith("Figura "):
            figure_entries.append(paragraph.text.strip())
        elif paragraph.text.strip().startswith("Tabla "):
            table_entries.append(paragraph.text.strip())

    heading = None
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Título Índices" and "Índice de contenidos" in paragraph.text:
            heading = paragraph
            break
    if heading is None:
        raise ValueError("Index heading not found")
    heading.paragraph_format.page_break_before = True
    current = heading._p.getnext()
    intro = find_heading(doc, "Introducción")._p
    while current is not None and current is not intro:
        nxt = current.getnext()
        current.getparent().remove(current)
        current = nxt
    cursor = heading._p
    for display_text, raw_text, level in heading_entries:
        cursor = _add_static_index_entry(doc, cursor, display_text, INDEX_PAGE_MAP.get(raw_text), level)

    cursor = _add_page_break(doc, cursor)
    cursor = insert_paragraph(doc, cursor, "Índice de figuras", "Título Índices")
    paragraph_from_element(doc, cursor).paragraph_format.space_before = Pt(18)
    for text in figure_entries:
        cursor = _add_static_index_entry(doc, cursor, text, FIGURE_PAGE_MAP.get(text), 1, False)

    cursor = _add_page_break(doc, cursor)
    cursor = insert_paragraph(doc, cursor, "Índice de tablas", "Título Índices")
    paragraph_from_element(doc, cursor).paragraph_format.space_before = Pt(18)
    for text in table_entries:
        cursor = _add_static_index_entry(doc, cursor, text, TABLE_PAGE_MAP.get(text), 1, False)
    return cursor


def replace_summary(doc: Document):
    for index, paragraph in enumerate(doc.paragraphs):
        if paragraph.text.strip() == "Resumen":
            target = doc.paragraphs[index + 1]
            target.text = (
                "Este Trabajo Fin de Estudios presenta el desarrollo y la evaluación de un prototipo local para la predicción de demanda y la recomendación de reposición en un mini-market. Se construyó un conjunto sintético reproducible con 7.310 registros diarios, diez productos y siete categorías. El pipeline genera variables de calendario, precio, promoción, retardos y ventanas móviles; compara un baseline estacional, Ridge y Random Forest mediante backtesting para horizontes de 7, 14 y 30 días. Random Forest obtuvo el mejor SMAPE medio multihorizonte (28,33 %), mientras Ridge presentó mejores resultados a 7 y 14 días. El sistema produjo 70, 140 y 300 predicciones por modelo según el horizonte y generó 140 registros operativos para diez productos en el escenario de 14 días. La solución integra Streamlit, FastAPI, SQLite, Docker y un flujo n8n, además de pruebas automatizadas y una regla explicable de reposición. Los resultados demuestran viabilidad técnica, aunque permanecen como limitaciones el carácter sintético de los datos y la ausencia de validación con usuarios reales."
            )
        if paragraph.text.strip() == "Abstract":
            target = doc.paragraphs[index + 1]
            target.text = (
                "This Final Project presents the development and evaluation of a local prototype for demand forecasting and inventory replenishment recommendations in a mini-market. A reproducible synthetic dataset was built with 7,310 daily records, ten products and seven categories. The pipeline creates calendar, price, promotion, lag and rolling-window features and compares a seasonal baseline, Ridge and Random Forest through backtesting at 7-, 14- and 30-day horizons. Random Forest achieved the best average multi-horizon SMAPE (28.33%), while Ridge performed better at 7 and 14 days. The system produced 70, 140 and 300 predictions per model depending on the horizon, as well as 140 operational records for ten products in the 14-day scenario. The solution integrates Streamlit, FastAPI, SQLite, Docker and an n8n workflow, together with automated tests and an explainable replenishment rule. Results support the technical feasibility of the approach, while synthetic data and the lack of validation with real users remain explicit limitations."
            )


def mark_pending_cover_date(doc: Document):
    """Evita presentar una fecha de depósito no confirmada como dato definitivo."""
    if not doc.tables:
        return
    for row in doc.tables[0].rows:
        if row.cells[0].text.strip().startswith("Fecha:"):
            cell = row.cells[1]
            cell.text = "PENDIENTE: confirmar fecha de depósito extraordinario"
            shading = cell._tc.get_or_add_tcPr().find(qn("w:shd"))
            if shading is None:
                shading = OxmlElement("w:shd")
                cell._tc.get_or_add_tcPr().append(shading)
            shading.set(qn("w:fill"), "FFF2CC")
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(9)
            return
def ensure_page_break_before_headings(doc: Document):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Heading 1":
            paragraph.paragraph_format.page_break_before = True


def add_page_numbering(doc: Document):
    for section in doc.sections:
        section.different_first_page_header_footer = True
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.clear()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_field(paragraph, "PAGE", "1")
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9)


def normalize_new_text(doc: Document):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Normal":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            if not run.font.name:
                run.font.name = "Arial"
                run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Arial")
                run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Arial")


def normalize_heading_alignment(doc: Document):
    """Impide que los títulos cortos se expandan por la justificación."""
    for paragraph in doc.paragraphs:
        if paragraph.style.name in {"Heading 1", "Heading 2", "Heading 3"}:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def format_references(doc: Document):
    """Aplica alineación izquierda y sangría francesa a las referencias APA."""
    in_references = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "Referencias bibliográficas":
            in_references = True
            continue
        if in_references and paragraph.style.name == "Heading 1":
            break
        if in_references and text:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Inches(0.5)
            paragraph.paragraph_format.first_line_indent = Inches(-0.5)
            paragraph.paragraph_format.space_after = Pt(6)


def normalize_table_geometry(doc: Document, max_width_twips: int = 9000):
    for table in doc.tables:
        grid = table._tbl.tblGrid
        columns = list(grid.gridCol_lst)
        widths = [int(column.get(qn("w:w"), "0")) for column in columns]
        if not widths or sum(widths) <= 0:
            widths = [max_width_twips // len(table.columns)] * len(table.columns)
        total = sum(widths)
        if total > max_width_twips:
            factor = max_width_twips / total
            scaled = [max(240, round(width * factor)) for width in widths]
            scaled[-1] += max_width_twips - sum(scaled)
            widths = scaled
            total = sum(widths)

        tbl_pr = table._tbl.tblPr
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.insert(0, tbl_w)
        tbl_w.set(qn("w:type"), "dxa")
        tbl_w.set(qn("w:w"), str(total))
        tbl_ind = tbl_pr.find(qn("w:tblInd"))
        if tbl_ind is None:
            tbl_ind = OxmlElement("w:tblInd")
            tbl_pr.append(tbl_ind)
        tbl_ind.set(qn("w:type"), "dxa")
        tbl_ind.set(qn("w:w"), "120")

        for column, width in zip(columns, widths):
            column.set(qn("w:w"), str(width))
        for row in table.rows:
            for cell, width in zip(row.cells, widths):
                tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
                if tc_w is None:
                    tc_w = OxmlElement("w:tcW")
                    cell._tc.get_or_add_tcPr().append(tc_w)
                tc_w.set(qn("w:type"), "dxa")
                tc_w.set(qn("w:w"), str(width))


def tune_table_layout(doc: Document):
    width_profiles = {
        ("Código", "Caso de uso", "Actor principal", "Descripción", "Resultado esperado"): [850, 1450, 1300, 2700, 2700],
        ("Código", "Requisito funcional", "Caso de uso", "Prioridad", "Criterio de aceptación"): [850, 2750, 900, 900, 3600],
        ("Código", "Requisito no funcional", "Justificación", "Criterio de aceptación"): [850, 1850, 2550, 3750],
        ("Elemento", "Descripción", "Criterio de aceptación"): [1550, 3500, 3950],
        ("Contenedor", "Componente", "Puerto", "Comprobación", "Resultado"): [3650, 1450, 800, 1600, 1500],
        ("Proceso", "Comando", "Segundos", "Criterio", "Estado"): [1800, 3500, 1200, 1100, 1400],
    }
    for table in doc.tables:
        headers = tuple(cell.text.strip() for cell in table.rows[0].cells)
        widths = width_profiles.get(headers)
        table.autofit = False

        if widths:
            for grid_column, width in zip(table._tbl.tblGrid.gridCol_lst, widths):
                grid_column.set(qn("w:w"), str(width))
            for row in table.rows:
                for cell, width in zip(row.cells, widths):
                    tc_pr = cell._tc.get_or_add_tcPr()
                    tc_w = tc_pr.find(qn("w:tcW"))
                    if tc_w is None:
                        tc_w = OxmlElement("w:tcW")
                        tc_pr.append(tc_w)
                    tc_w.set(qn("w:type"), "dxa")
                    tc_w.set(qn("w:w"), str(width))

        header_properties = table.rows[0]._tr.get_or_add_trPr()
        if header_properties.find(qn("w:tblHeader")) is None:
            repeat = OxmlElement("w:tblHeader")
            repeat.set(qn("w:val"), "true")
            header_properties.append(repeat)

        for row_index, row in enumerate(table.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.0
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
                        if row_index == 0:
                            run.bold = True


def main() -> None:
    shutil.copy2(REFERENCE, OUTPUT)
    doc = Document(OUTPUT)
    replace_summary(doc)
    mark_pending_cover_date(doc)
    revise_existing_sections(doc)
    repair_existing_headings(doc)

    sales = pd.read_csv(PROJECT / "data" / "raw" / "sales_sample.csv", parse_dates=["date"])
    inventory = pd.read_csv(PROJECT / "data" / "raw" / "inventory_sample.csv")
    metrics_h = pd.read_csv(PROJECT / "reports" / "metrics_by_horizon.csv")
    metrics_simple = pd.read_csv(PROJECT / "reports" / "metrics.csv")
    metrics_p = pd.read_csv(PROJECT / "reports" / "metrics_by_product.csv")
    ranking = pd.read_csv(PROJECT / "reports" / "model_ranking.csv")
    recommendations = pd.read_csv(PROJECT / "reports" / "recommendations.csv")
    sensitivity = pd.read_csv(PROJECT / "reports" / "replenishment_sensitivity.csv")

    build_methodology(doc, sales)
    build_development(doc, sales, inventory)
    build_evaluation(doc, sales, metrics_simple, metrics_h, metrics_p, ranking, recommendations, sensitivity)
    build_conclusions(doc)
    build_annex(doc)
    remove_uncited_references(doc)
    repair_spelling(doc)
    italicize_technical_terms(doc)
    italicize_reference_titles(doc)
    renumber_captions(doc)
    rebuild_indexes(doc)
    ensure_page_break_before_headings(doc)
    add_page_numbering(doc)
    normalize_new_text(doc)
    normalize_heading_alignment(doc)
    format_references(doc)
    normalize_table_geometry(doc)
    tune_table_layout(doc)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
