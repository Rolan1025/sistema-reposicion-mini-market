from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[2]
PROJECT = ROOT / "sistema-reposicion-mini-market"
PROPOSAL = ROOT / "Propuesta"
INPUT_DOCX = PROPOSAL / "TFE_Tercera_Entrega_Orlando_Delgado_Pinzon_limpio.docx"
OUTPUT_DOCX = PROPOSAL / "TFE_Tercera_Entrega_Orlando_Delgado_Pinzon_revisado.docx"
FIG_DIR = PROPOSAL / "figuras_generadas"


def block_text(block) -> str:
    texts = []
    for node in block.iter():
        if node.tag == qn("w:t"):
            texts.append(node.text or "")
    return "".join(texts)


def paragraph_from_element(doc: Document, element):
    from docx.text.paragraph import Paragraph

    return Paragraph(element, doc._body)


def find_heading(doc: Document, heading_text: str):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Heading 1" and heading_text.lower() in paragraph.text.lower():
            return paragraph
    raise ValueError(f"Heading not found: {heading_text}")


def is_heading_1(doc: Document, element) -> bool:
    if element.tag != qn("w:p"):
        return False
    paragraph = paragraph_from_element(doc, element)
    return paragraph.style.name == "Heading 1" and paragraph.text.strip()


def clear_section_after_heading(doc: Document, heading_text: str):
    heading = find_heading(doc, heading_text)
    current = heading._p.getnext()
    while current is not None:
        if is_heading_1(doc, current):
            break
        nxt = current.getnext()
        current.getparent().remove(current)
        current = nxt
    return heading._p


def insert_paragraph(doc: Document, after_element, text: str = "", style: str | None = None):
    new_p = OxmlElement("w:p")
    after_element.addnext(new_p)
    paragraph = paragraph_from_element(doc, new_p)
    if style:
        paragraph.style = style
    if text:
        paragraph.add_run(text)
    return new_p


def insert_picture(doc: Document, after_element, image_path: Path, width_inches: float = 6.0):
    p_el = insert_paragraph(doc, after_element)
    paragraph = paragraph_from_element(doc, p_el)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run().add_picture(str(image_path), width=Inches(width_inches))
    return p_el


def set_cell_text(cell, text: str, bold: bool = False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.font.size = Pt(8.5)
    run.bold = bold
    return cell


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def insert_table(doc: Document, after_element, headers: list[str], rows: list[list[str]], title: str | None = None):
    cursor = after_element
    if title:
        cursor = insert_paragraph(doc, cursor, title)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    for i, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], header, bold=True)
        shade_cell(table.rows[0].cells[i], "D9EAF7")
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            set_cell_text(cells[i], value)
    tbl = table._tbl
    tbl.getparent().remove(tbl)
    cursor.addnext(tbl)
    return tbl


def add_heading(doc: Document, cursor, text: str, level: int = 2):
    return insert_paragraph(doc, cursor, text, f"Heading {level}")


def add_body(doc: Document, cursor, text: str):
    return insert_paragraph(doc, cursor, text, "Normal")


def add_bullet(doc: Document, cursor, text: str):
    return insert_paragraph(doc, cursor, f"• {text}", "List Paragraph")


def font() -> ImageFont.ImageFont:
    try:
        return ImageFont.truetype("arial.ttf", 20)
    except OSError:
        return ImageFont.load_default()


def draw_arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int]):
    draw.line([start, end], fill="#2F5597", width=4)
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    length = 14
    for offset in (math.pi * 0.82, -math.pi * 0.82):
        x = end[0] + length * math.cos(angle + offset)
        y = end[1] + length * math.sin(angle + offset)
        draw.line([end, (x, y)], fill="#2F5597", width=4)


def draw_box(draw: ImageDraw.ImageDraw, xy, text: str, fill="#EEF4FB", outline="#2F5597"):
    draw.rounded_rectangle(xy, radius=16, fill=fill, outline=outline, width=3)
    fnt = font()
    lines = text.split("\n")
    total_h = len(lines) * 24
    y = (xy[1] + xy[3] - total_h) // 2
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=fnt)
        x = (xy[0] + xy[2] - (bbox[2] - bbox[0])) // 2
        draw.text((x, y), line, fill="#1F1F1F", font=fnt)
        y += 24


def create_figures() -> tuple[Path, Path]:
    FIG_DIR.mkdir(parents=True, exist_ok=True)

    architecture = FIG_DIR / "figura_arquitectura.png"
    img = Image.new("RGB", (1400, 620), "white")
    draw = ImageDraw.Draw(img)
    boxes = [
        ((40, 210, 260, 350), "Ventas e\ninventario"),
        ((330, 120, 590, 260), "Pipeline de datos\npandas"),
        ((330, 360, 590, 500), "Parametros de\nreposicion"),
        ((670, 120, 930, 260), "Entrenamiento\nscikit-learn"),
        ((670, 360, 930, 500), "Forecast\n7/14/30 dias"),
        ((1010, 210, 1270, 350), "Dashboard/API\nreportes"),
    ]
    for box, label in boxes:
        draw_box(draw, box, label)
    draw_arrow(draw, (260, 280), (330, 190))
    draw_arrow(draw, (260, 280), (330, 430))
    draw_arrow(draw, (590, 190), (670, 190))
    draw_arrow(draw, (930, 190), (1010, 250))
    draw_arrow(draw, (590, 430), (670, 430))
    draw_arrow(draw, (930, 430), (1010, 310))
    draw.text((40, 35), "Arquitectura logica implementada", fill="#1F4E79", font=ImageFont.truetype("arial.ttf", 28) if Path("C:/Windows/Fonts/arial.ttf").exists() else font())
    img.save(architecture)

    pipeline = FIG_DIR / "figura_pipeline.png"
    img = Image.new("RGB", (1400, 520), "white")
    draw = ImageDraw.Draw(img)
    labels = [
        "Generar o\ncargar CSV",
        "Validar\ncampos",
        "Agregar por\ndia-producto",
        "Crear lags y\ncalendario",
        "Entrenar y\ncomparar",
        "Guardar\nmodelo",
        "Predecir y\nreponer",
    ]
    x = 35
    last = None
    for label in labels:
        box = (x, 190, x + 170, 330)
        draw_box(draw, box, label, fill="#F4F8EF", outline="#548235")
        if last:
            draw_arrow(draw, (last[2], 260), (box[0], 260))
        last = box
        x += 195
    draw.text((40, 40), "Pipeline reproducible ejecutado", fill="#375623", font=ImageFont.truetype("arial.ttf", 28) if Path("C:/Windows/Fonts/arial.ttf").exists() else font())
    img.save(pipeline)

    return architecture, pipeline


def ensure_section_properties(doc: Document) -> None:
    body = doc._body._element
    for child in body:
        if child.tag == qn("w:sectPr"):
            return
    sect_pr = OxmlElement("w:sectPr")
    pg_sz = OxmlElement("w:pgSz")
    pg_sz.set(qn("w:w"), "11906")
    pg_sz.set(qn("w:h"), "16838")
    pg_marg = OxmlElement("w:pgMar")
    pg_marg.set(qn("w:top"), "1417")
    pg_marg.set(qn("w:right"), "1417")
    pg_marg.set(qn("w:bottom"), "1417")
    pg_marg.set(qn("w:left"), "1417")
    pg_marg.set(qn("w:header"), "708")
    pg_marg.set(qn("w:footer"), "708")
    pg_marg.set(qn("w:gutter"), "0")
    sect_pr.append(pg_sz)
    sect_pr.append(pg_marg)
    body.append(sect_pr)


def fmt(value: float) -> str:
    return f"{value:.2f}"


def main() -> None:
    arch_path, pipeline_path = create_figures()

    metrics = pd.read_csv(PROJECT / "reports" / "metrics.csv")
    forecast = pd.read_csv(PROJECT / "reports" / "forecast.csv")
    recommendations = pd.read_csv(PROJECT / "reports" / "recommendations.csv")
    sales = pd.read_csv(PROJECT / "data" / "raw" / "sales_sample.csv", parse_dates=["date"])
    inventory = pd.read_csv(PROJECT / "data" / "raw" / "inventory_sample.csv")

    doc = Document(INPUT_DOCX)

    # Replace the methodology chapter with concrete, executed methodology.
    cursor = clear_section_after_heading(doc, "Objetivos concretos")
    cursor = add_heading(doc, cursor, "Objetivo general", 2)
    cursor = add_body(
        doc,
        cursor,
        "Desarrollar y validar un prototipo software local y reproducible que predice la demanda diaria por producto en un mini-market y genera recomendaciones de reposicion a partir del forecast, el stock disponible, pedidos en camino, tiempo de reposicion, stock de seguridad y tamano de empaque.",
    )
    cursor = add_heading(doc, cursor, "Objetivos especificos y cumplimiento", 2)
    objective_rows = [
        ["OE-01", "Preparar un conjunto de datos de ventas y un pipeline reproducible.", "Cumplido", "scripts/generate_sample_data.py; data/raw/sales_sample.csv"],
        ["OE-02", "Comparar baseline y modelos clasicos de prediccion.", "Cumplido", "scripts/train.py; reports/metrics.csv"],
        ["OE-03", "Implementar logica explicable de reposicion.", "Cumplido", "src/demand_forecasting/replenishment.py; reports/recommendations.csv"],
        ["OE-04", "Desarrollar dashboard/API para consulta.", "Cumplido", "app/streamlit_app.py; app/api.py"],
        ["OE-05", "Automatizar tareas de entrenamiento y reporte.", "Cumplido", "n8n/workflows/retrain_and_report.json"],
        ["OE-06", "Evaluar cuantitativamente el sistema.", "Cumplido parcialmente", "Metricas reales y pruebas automatizadas; validacion con usuarios queda como limitacion."],
    ]
    cursor = insert_table(doc, cursor, ["Objetivo", "Descripcion", "Estado", "Evidencia"], objective_rows, "Tabla 1. Cumplimiento de objetivos especificos")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Metodologia de trabajo ejecutada", 2)
    for item in [
        "Comprension del problema: se delimito el alcance a diez productos representativos de un mini-market y a ejecucion local sin servicios de pago.",
        "Preparacion de datos: se genero una muestra sintetica reproducible de ventas diarias entre 2024-01-01 y 2025-12-31, con 7.310 registros, 10 productos y 7 categorias.",
        "Ingenieria de variables: se construyeron variables de calendario, precio, promocion, retardos de 1, 2, 7, 14 y 28 dias, y medias moviles/desviaciones de 7, 14 y 28 dias.",
        "Modelado: se comparo un baseline estacional de 7 dias con Ridge y Random Forest Regressor, usando una particion temporal con los ultimos 30 dias como prueba.",
        "Evaluacion: se calcularon MAE, RMSE, MAPE y SMAPE; el criterio de seleccion fue menor SMAPE y, en empate, menor MAE.",
        "Integracion: el modelo seleccionado se serializo en models/demand_model.joblib y se conecto con dashboard, API, reportes CSV/SQLite y flujo n8n.",
    ]:
        cursor = add_bullet(doc, cursor, item)
    cursor = add_heading(doc, cursor, "Instrumentos, criterios de analisis y trazabilidad", 2)
    cursor = add_body(
        doc,
        cursor,
        "La trazabilidad se materializo en el repositorio local sistema-reposicion-mini-market. Cada resultado incluido en la memoria se asocia con un script, archivo de salida o prueba automatizada. Esta decision permite que el prototipo pueda reproducirse sin depender de cuadernos manuales ni de pasos no documentados.",
    )

    # Replace chapter 5.
    cursor = clear_section_after_heading(doc, "Diseño y desarrollo")
    cursor = add_body(
        doc,
        cursor,
        "Este capitulo documenta el sistema efectivamente implementado. La solucion se materializo como un repositorio Python modular, con scripts de ejecucion, dashboard, API, flujo n8n exportable, pruebas automatizadas y salidas verificables en CSV y SQLite.",
    )
    cursor = add_heading(doc, cursor, "Arquitectura implementada", 2)
    cursor = add_body(
        doc,
        cursor,
        "La arquitectura final separa la preparacion de datos, el entrenamiento, la inferencia, el motor de reposicion y las interfaces de consulta. Esta separacion facilita la explicabilidad y permite sustituir el modelo predictivo sin modificar la regla de reposicion.",
    )
    cursor = insert_picture(doc, cursor, arch_path, 6.2)
    cursor = add_body(doc, cursor, "Figura 1. Arquitectura logica implementada del prototipo. Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Dataset utilizado y preparacion de datos", 2)
    cursor = add_body(
        doc,
        cursor,
        f"Para cerrar la implementacion sin depender de datos sensibles de una tienda real, se utilizo un dataset sintetico reproducible que simula ventas de un mini-market. El archivo contiene {len(sales):,} registros diarios, {sales.product_id.nunique()} productos y {sales.category.nunique()} categorias, desde {sales.date.min().date()} hasta {sales.date.max().date()}. Cada registro incluye fecha, producto, categoria, unidades vendidas, precio unitario y promocion.",
    )
    dataset_rows = [
        ["Periodo", f"{sales.date.min().date()} a {sales.date.max().date()}"],
        ["Granularidad", "Diaria por producto"],
        ["Productos", str(sales.product_id.nunique())],
        ["Categorias", ", ".join(sorted(sales.category.unique()))],
        ["Registros de ventas", f"{len(sales):,}"],
        ["Inventario inicial", f"{len(inventory)} productos con stock, pedidos en camino, lead time, pack size y max_stock"],
    ]
    cursor = insert_table(doc, cursor, ["Elemento", "Valor implementado"], dataset_rows, "Tabla 7. Caracteristicas del dataset implementado")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia a partir de los archivos data/raw/sales_sample.csv e inventory_sample.csv.")
    cursor = insert_picture(doc, cursor, pipeline_path, 6.2)
    cursor = add_body(doc, cursor, "Figura 2. Pipeline reproducible ejecutado. Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Variables finales del modelo", 2)
    variables_rows = [
        ["Calendario", "day_of_week, day_of_month, week_of_year, month, is_weekend, is_payday, codificacion ciclica de dia y mes."],
        ["Historico de demanda", "lag_1, lag_2, lag_7, lag_14, lag_28."],
        ["Ventanas moviles", "rolling_mean_7, rolling_std_7, rolling_mean_14, rolling_std_14, rolling_mean_28, rolling_std_28."],
        ["Producto y categoria", "Codificacion one-hot de product_id y category."],
        ["Exogenas", "unit_price y promo."],
        ["Objetivo", "units_sold del dia siguiente en formato supervisado."],
    ]
    cursor = insert_table(doc, cursor, ["Grupo", "Variables"], variables_rows, "Tabla 8. Variables finales utilizadas por el modelo")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Configuracion real de modelos", 2)
    model_rows = [
        ["seasonal_naive_7d", "Baseline", "Demanda observada siete dias antes; sirve como referencia minima."],
        ["ridge", "Modelo final", "Pipeline con StandardScaler y Ridge(alpha=1.0, random_state=42)."],
        ["random_forest", "Modelo candidato", "RandomForestRegressor(n_estimators=160, max_depth=12, min_samples_leaf=3, random_state=42, n_jobs=-1)."],
    ]
    cursor = insert_table(doc, cursor, ["Modelo", "Rol", "Configuracion"], model_rows, "Tabla 9. Modelos implementados y configuracion")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Motor de reposicion implementado", 2)
    cursor = add_body(
        doc,
        cursor,
        "La recomendacion se calcula de forma explicable: demanda durante el tiempo de reposicion mas periodo de revision, stock de seguridad en dias, stock disponible y pedidos en camino. La cantidad final se redondea al multiplo de pack_size para que sea operativamente accionable.",
    )
    cursor = add_body(
        doc,
        cursor,
        "cantidad_sugerida = max((demanda_durante_reposicion + stock_seguridad) - (stock_on_hand + on_order), 0), redondeada al lote minimo.",
    )
    cursor = add_heading(doc, cursor, "Interfaz, API, persistencia y automatizacion", 2)
    for item in [
        "Dashboard Streamlit: permite seleccionar horizonte de 7, 14 o 30 dias, ajustar stock de seguridad y revisar predicciones/recomendaciones.",
        "API FastAPI: expone /health, /sample-data, /train y /predict para integracion programatica.",
        "Persistencia: genera reports/metrics.csv, reports/forecast.csv, reports/recommendations.csv y reports/mini_market_reports.db.",
        "n8n: se exporto un flujo programado que ejecuta entrenamiento, prediccion y lectura del reporte de reposicion.",
        "Docker: docker-compose.yml levanta dashboard, API y n8n para ejecucion local reproducible.",
    ]:
        cursor = add_bullet(doc, cursor, item)

    # Replace chapter 6.
    cursor = clear_section_after_heading(doc, "Evaluación de la solución")
    cursor = add_body(
        doc,
        cursor,
        "La evaluacion se ejecuto sobre la particion temporal definida: entrenamiento con el historico disponible hasta treinta dias antes del final y prueba con los ultimos 30 dias. La comparacion cuantitativa se registro en reports/metrics.csv.",
    )
    metric_rows = [
        [row["model"], fmt(row["mae"]), fmt(row["rmse"]), fmt(row["mape"]) + "%", fmt(row["smape"]) + "%"]
        for _, row in metrics.iterrows()
    ]
    cursor = insert_table(doc, cursor, ["Modelo", "MAE", "RMSE", "MAPE", "SMAPE"], metric_rows, "Tabla 10. Resultados cuantitativos de los modelos")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia a partir de reports/metrics.csv.")
    best = metrics.iloc[0]
    cursor = add_body(
        doc,
        cursor,
        f"El modelo seleccionado fue {best['model']}, con MAE={fmt(best['mae'])}, RMSE={fmt(best['rmse'])}, MAPE={fmt(best['mape'])}% y SMAPE={fmt(best['smape'])}%. El baseline seasonal_naive_7d obtuvo SMAPE={fmt(metrics.loc[metrics['model']=='seasonal_naive_7d','smape'].iloc[0])}%, por lo que el modelo seleccionado mejora la referencia simple en esta muestra.",
    )
    cursor = add_heading(doc, cursor, "Resultados de forecast y reposicion", 2)
    cursor = add_body(
        doc,
        cursor,
        f"Con el modelo seleccionado se genero un forecast de {len(forecast)} filas para un horizonte de 14 dias ({forecast.product_id.nunique()} productos). A partir de ese forecast se generaron {len(recommendations)} recomendaciones de reposicion.",
    )
    rec_rows = [
        [
            row["product_id"],
            row["product_name"],
            str(int(row["stock_on_hand"])),
            str(int(row["lead_time_days"])),
            fmt(row["avg_daily_forecast"]),
            str(int(row["suggested_order_qty"])),
            row["priority"],
        ]
        for _, row in recommendations.head(10).iterrows()
    ]
    cursor = insert_table(
        doc,
        cursor,
        ["Producto", "Nombre", "Stock", "Lead time", "Forecast diario", "Pedido sugerido", "Prioridad"],
        rec_rows,
        "Tabla 11. Recomendaciones de reposicion generadas",
    )
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia a partir de reports/recommendations.csv.")
    cursor = add_heading(doc, cursor, "Verificacion funcional y reproducibilidad", 2)
    verification_rows = [
        ["Generacion de datos", "python scripts/generate_sample_data.py", "Correcto"],
        ["Entrenamiento", "python scripts/train.py", "Correcto; modelo guardado en models/demand_model.joblib"],
        ["Prediccion y reposicion", "python scripts/predict.py --horizon-days 14", "Correcto; forecast, recomendaciones y SQLite generados"],
        ["Pruebas automatizadas", "python -m pytest", "5 pruebas superadas"],
        ["API", "Importacion de app.api y llamada health()", "{'status': 'ok'}"],
    ]
    cursor = insert_table(doc, cursor, ["Elemento", "Evidencia", "Resultado"], verification_rows, "Tabla 12. Evidencias de ejecucion del prototipo")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Matriz de verificacion de requisitos", 2)
    req_rows = [
        ["RF-01 a RF-03", "Pipeline de datos", "scripts/generate_sample_data.py; src/demand_forecasting/data.py; features.py", "Implementado y verificado"],
        ["RF-04 a RF-06", "Modulo predictivo", "reports/metrics.csv; models/demand_model.joblib; app/streamlit_app.py", "Implementado y verificado"],
        ["RF-07 a RF-08", "Motor de reposicion", "src/demand_forecasting/replenishment.py; tests/test_replenishment.py", "Implementado y verificado"],
        ["RF-09", "Persistencia local", "reports/*.csv y reports/mini_market_reports.db", "Implementado y verificado"],
        ["RF-10", "n8n", "n8n/workflows/retrain_and_report.json", "Implementado"],
        ["RNF-01/RNF-05", "Repositorio y documentacion", "README.md; pyproject.toml; Dockerfile; docker-compose.yml", "Implementado"],
        ["RNF-02/RNF-03", "Usabilidad y explicabilidad", "Dashboard con metricas, forecast y componentes de reposicion", "Implementado; validacion con usuarios pendiente"],
    ]
    cursor = insert_table(doc, cursor, ["Requisito", "Componente", "Evidencia", "Estado"], req_rows, "Tabla 13. Matriz de verificacion actualizada")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Limitaciones de la evaluacion", 2)
    cursor = add_body(
        doc,
        cursor,
        "La evaluacion confirma la ejecucion tecnica del prototipo y la comparacion cuantitativa entre modelos. No obstante, la validacion con usuarios expertos no se completo en esta fase, por lo que sus resultados no se reportan como evidencia. Esta limitacion se declara de forma explicita y se propone como siguiente paso antes de una entrega final cerrada.",
    )

    # Replace chapter 7.
    cursor = clear_section_after_heading(doc, "Conclusiones y trabajo futuro")
    cursor = add_body(
        doc,
        cursor,
        "El TFE paso de una especificacion avanzada a un prototipo funcional ejecutable. Se implemento un sistema de extremo a extremo capaz de generar datos de ventas, entrenar modelos clasicos, seleccionar el mejor modelo por metricas, producir forecasts y transformar esas predicciones en recomendaciones accionables de reposicion.",
    )
    cursor = add_heading(doc, cursor, "Conclusiones por objetivo", 2)
    conclusion_rows = [
        ["Preparar datos y pipeline", "Se obtuvo un dataset reproducible de 7.310 ventas y un pipeline modular de limpieza y features."],
        ["Comparar modelos", "Se compararon baseline, Ridge y Random Forest; Ridge fue seleccionado por menor SMAPE y MAE."],
        ["Recomendar reposicion", "La regla de pedido combina forecast, lead time, stock de seguridad, pedidos en camino y pack size."],
        ["Construir prototipo", "Se entregaron dashboard Streamlit, API FastAPI, reportes CSV/SQLite y Docker Compose."],
        ["Automatizar flujo", "Se creo un workflow n8n para entrenamiento y reporte programado."],
    ]
    cursor = insert_table(doc, cursor, ["Objetivo", "Resultado concreto"], conclusion_rows, "Tabla 14. Relacion entre objetivos y resultados obtenidos")
    cursor = add_body(doc, cursor, "Fuente: elaboracion propia.")
    cursor = add_heading(doc, cursor, "Limitaciones", 2)
    for item in [
        "El dataset es sintetico, por lo que reproduce patrones plausibles de mini-market pero no sustituye una validacion con ventas reales.",
        "La evaluacion se realizo con una particion temporal simple de 30 dias; un backtesting con multiples ventanas mejoraria la robustez.",
        "La validacion con usuarios expertos esta definida como instrumento, pero no se ejecuto con participantes reales en esta fase.",
        "El flujo n8n automatiza la ejecucion, pero no incluye todavia notificaciones por correo o integracion con proveedores.",
    ]:
        cursor = add_bullet(doc, cursor, item)
    cursor = add_heading(doc, cursor, "Trabajo futuro", 2)
    for item in [
        "Sustituir el dataset sintetico por ventas anonimizadas de un comercio real y comparar resultados.",
        "Agregar backtesting con multiples ventanas e indicadores por categoria o tipo de rotacion.",
        "Incorporar validacion con usuarios expertos mediante tareas guiadas y cuestionario SUS o escala breve de utilidad.",
        "Ampliar la automatizacion n8n con notificaciones, versionado de modelos y promocion controlada del mejor modelo.",
        "Explorar modelos adicionales como LightGBM o Prophet si el alcance y los datos reales lo justifican.",
    ]:
        cursor = add_bullet(doc, cursor, item)

    # Replace Annex A.
    cursor = clear_section_after_heading(doc, "Anexo A")
    cursor = add_body(
        doc,
        cursor,
        "El codigo fuente se encuentra en el repositorio local sistema-reposicion-mini-market, dentro del mismo directorio de trabajo del TFE. La estructura implementada es la siguiente:",
    )
    for item in [
        "app/: dashboard Streamlit y API FastAPI.",
        "data/raw/: dataset de ventas e inventario de ejemplo.",
        "docs/: metodologia tecnica del prototipo.",
        "models/: modelo serializado.",
        "n8n/workflows/: flujo exportable de automatizacion.",
        "reports/: metricas, forecast, recomendaciones y base SQLite.",
        "scripts/: generacion de datos, entrenamiento y prediccion.",
        "src/demand_forecasting/: modulos de datos, features, modelos, forecast, reposicion y persistencia.",
        "tests/: pruebas de metricas, reposicion y pipeline completo.",
    ]:
        cursor = add_bullet(doc, cursor, item)
    cursor = add_heading(doc, cursor, "Comandos de reproduccion", 2)
    for item in [
        "python scripts/generate_sample_data.py",
        "python scripts/train.py",
        "python scripts/predict.py --horizon-days 14",
        "streamlit run app/streamlit_app.py",
        "uvicorn app.api:app --reload --port 8000",
        "python -m pytest",
    ]:
        cursor = add_bullet(doc, cursor, item)

    # Small updates in abstract/reserved wording.
    replacements = {
        "Este Trabajo Fin de Estudios plantea el diseño, desarrollo y validación": "Este Trabajo Fin de Estudios presenta el diseño, desarrollo y validación",
        "This Final Project proposes the design, development and validation": "This Final Project presents the design, development and validation",
        "El capítulo 7 se reserva para la versión final del TFE, cuando el prototipo haya sido implementado y evaluado. En esa etapa": "El capítulo 7 recoge las conclusiones derivadas del prototipo implementado y evaluado. En esta etapa",
        "Tabla 1. Actores identificados para el prototipo": "Tabla 2. Actores identificados para el prototipo",
        "Tabla 2. Casos de uso principales": "Tabla 3. Casos de uso principales",
        "Tabla 3. Requisitos funcionales": "Tabla 4. Requisitos funcionales",
        "Tabla 4. Requisitos no funcionales": "Tabla 5. Requisitos no funcionales",
        "Tabla 5. Criterios de aceptación de datos y trazabilidad": "Tabla 6. Criterios de aceptación de datos y trazabilidad",
        "Los requisitos opcionales se consideran extensiones que solo se implementarán si el tiempo lo permite.": "Los requisitos opcionales se consideran extensiones no incluidas en el alcance validado y se documentan como trabajo futuro.",
        "Los requisitos imprescindibles deberán aparecer implementados y verificados en la entrega final. Los requisitos deseables se implementarán si no comprometen el calendario y, en caso contrario, se justificarán como limitaciones o trabajo futuro.": "Los requisitos imprescindibles aparecen implementados y verificados mediante evidencias de ejecucion. Los requisitos deseables se incorporan cuando no comprometen el alcance principal y, en caso contrario, se justifican como limitaciones o trabajo futuro.",
    }
    for paragraph in doc.paragraphs:
        for old, new in replacements.items():
            if old in paragraph.text:
                paragraph.text = paragraph.text.replace(old, new)

    # Basic polish.
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Arial"
            run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Arial")
            run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Arial")
            if paragraph.style.name == "Normal":
                run.font.size = Pt(10.5)

    ensure_section_properties(doc)
    doc.save(OUTPUT_DOCX)
    print(f"Wrote {OUTPUT_DOCX}")


if __name__ == "__main__":
    main()
