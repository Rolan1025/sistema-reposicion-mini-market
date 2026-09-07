from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import DEFAULT_INVENTORY_CSV, DEFAULT_MODEL_PATH, DEFAULT_SALES_CSV, REPORTS_DIR
from demand_forecasting.data import load_sales_csv, write_sample_data
from demand_forecasting.models import load_model_bundle
from demand_forecasting.pipeline import predict_pipeline, train_pipeline


MODEL_NAMES = {
    "random_forest": "Random Forest",
    "ridge": "Ridge Regression",
    "seasonal_naive_7d": "Seasonal Naive (7 días)",
}
CATEGORY_NAMES = {
    "Lacteos": "Lácteos",
    "Panaderia": "Panadería",
    "Basicos": "Básicos",
}
PRODUCT_NAMES = {
    "Atun lata": "Atún lata",
    "Cafe molido 250g": "Café molido 250 g",
    "Papel higienico x4": "Papel higiénico x4",
}


def localized_values(data: pd.DataFrame) -> pd.DataFrame:
    """Traduce valores para la interfaz sin modificar los CSV originales."""
    display = data.copy()
    if "model" in display.columns:
        display["model"] = display["model"].replace(MODEL_NAMES)
    if "model_name" in display.columns:
        display["model_name"] = display["model_name"].replace(MODEL_NAMES)
    if "category" in display.columns:
        display["category"] = display["category"].replace(CATEGORY_NAMES)
    if "product_name" in display.columns:
        display["product_name"] = display["product_name"].replace(PRODUCT_NAMES)
    return display


FORECAST_COLUMNS = {
    "date": st.column_config.DateColumn(
        "Fecha",
        help="Día futuro para el que se calcula la predicción.",
        format="DD/MM/YYYY",
    ),
    "product_id": st.column_config.TextColumn(
        "Código",
        help="Identificador único del producto dentro del sistema.",
    ),
    "product_name": st.column_config.TextColumn(
        "Producto",
        help="Nombre comercial del producto.",
    ),
    "category": st.column_config.TextColumn(
        "Categoría",
        help="Familia a la que pertenece el producto.",
    ),
    "forecast_units": st.column_config.NumberColumn(
        "Unidades pronosticadas",
        help="Cantidad estimada que se venderá en esa fecha.",
        format="%.2f",
    ),
    "unit_price": st.column_config.NumberColumn(
        "Precio unitario",
        help="Último precio conocido utilizado como variable del modelo.",
        format="$ %.2f",
    ),
    "model_name": st.column_config.TextColumn(
        "Modelo utilizado",
        help="Nombre técnico internacional del algoritmo que generó la predicción mostrada.",
    ),
}

SALES_COLUMNS = {
    "date": FORECAST_COLUMNS["date"],
    "product_id": FORECAST_COLUMNS["product_id"],
    "product_name": FORECAST_COLUMNS["product_name"],
    "category": FORECAST_COLUMNS["category"],
    "units_sold": st.column_config.NumberColumn(
        "Unidades vendidas",
        help="Cantidad realmente vendida durante el día.",
        format="%d",
    ),
    "unit_price": FORECAST_COLUMNS["unit_price"],
    "promo": st.column_config.CheckboxColumn(
        "Promoción",
        help="Indica si el producto estuvo en promoción: marcado significa sí.",
    ),
    "revenue": st.column_config.NumberColumn(
        "Ingreso",
        help="Unidades vendidas multiplicadas por el precio unitario.",
        format="$ %.2f",
    ),
}


st.set_page_config(page_title="Reposición mini-market", layout="wide")

st.title("Sistema de predicción de demanda y reposición")

# Controles que el usuario puede modificar sin cambiar el código.
with st.sidebar:
    st.header("Parámetros")
    horizon_days = st.selectbox(
        "Horizonte de predicción",
        [7, 14, 30],
        index=1,
        help="Número de días futuros que pronosticará el sistema. Un horizonte mayor permite planificar más lejos, pero acumula más incertidumbre.",
    )
    safety_stock_days = st.slider(
        "Stock de seguridad (días)",
        min_value=0,
        max_value=14,
        value=3,
        help="Reserva adicional expresada en días de demanda. Al aumentarla, normalmente aumenta la cantidad recomendada y disminuye el riesgo de quedarse sin stock.",
    )
    review_period_days = st.slider(
        "Periodo de revisión (días)",
        min_value=1,
        max_value=7,
        value=1,
        help="Tiempo hasta la próxima revisión del inventario. Al aumentarlo, la compra debe cubrir más días.",
    )
    test_days = st.slider(
        "Ventana de evaluación (días)",
        min_value=14,
        max_value=90,
        value=30,
        help="Cantidad de días históricos que se ocultan para medir el modelo. Solo se utiliza al reentrenar.",
    )

    if st.button(
        "Restaurar datos de ejemplo",
        help="Sobrescribe los CSV actuales con el escenario sintético reproducible del TFE.",
    ):
        write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
        st.success("Datos de ejemplo generados.")

    train_clicked = st.button(
        "Reentrenar modelo",
        type="primary",
        help="Vuelve a comparar los modelos con los datos actuales y guarda el mejor.",
    )


# La aplicación se puede demostrar desde cero: crea datos o modelo solo cuando faltan.
if not DEFAULT_SALES_CSV.exists() or not DEFAULT_INVENTORY_CSV.exists():
    write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)

if train_clicked or not DEFAULT_MODEL_PATH.exists():
    with st.spinner("Entrenando modelos y seleccionando el mejor..."):
        _, metrics = train_pipeline(
            DEFAULT_SALES_CSV,
            DEFAULT_MODEL_PATH,
            REPORTS_DIR / "metrics.csv",
            test_days=test_days,
        )
    st.success("Modelo entrenado.")
else:
    metrics_path = REPORTS_DIR / "metrics.csv"
    metrics = pd.read_csv(metrics_path) if metrics_path.exists() else pd.DataFrame()

# Cada cambio de controles recalcula el forecast y la regla de reposición.
forecast, recommendations = predict_pipeline(
    DEFAULT_SALES_CSV,
    DEFAULT_INVENTORY_CSV,
    DEFAULT_MODEL_PATH,
    horizon_days=horizon_days,
    forecast_path=REPORTS_DIR / "forecast.csv",
    recommendations_path=REPORTS_DIR / "recommendations.csv",
    safety_stock_days=safety_stock_days,
    review_period_days=review_period_days,
)

sales = load_sales_csv(DEFAULT_SALES_CSV)

model_bundle = load_model_bundle(DEFAULT_MODEL_PATH)
operational_model = MODEL_NAMES.get(model_bundle.model_name, model_bundle.model_name)
metric_cols = st.columns(5)
if not metrics.empty:
    best = metrics.iloc[0]
    metric_cols[0].metric(
        "Modelo operativo",
        operational_model,
        help="Modelo seleccionado por el menor SMAPE promedio en los horizontes de 7, 14 y 30 días.",
    )
    metric_cols[1].metric(
        "Líder auxiliar",
        MODEL_NAMES.get(str(best["model"]), str(best["model"])),
        help="Mejor modelo en la comprobación auxiliar de la ventana final. No sustituye la selección multihorizonte del modelo operativo.",
    )
    metric_cols[2].metric(
        "MAE",
        f"{best['mae']:.2f}",
        help="Error absoluto medio del líder auxiliar: unidades que falla en promedio. Menor es mejor.",
    )
    metric_cols[3].metric(
        "SMAPE",
        f"{best['smape']:.2f}%",
        help="Error porcentual simétrico del líder auxiliar. Menor es mejor.",
    )
    metric_cols[4].metric(
        "MAPE",
        f"{best['mape']:.2f}%",
        help="Error porcentual del líder auxiliar respecto a las ventas reales. Puede ser inestable con demanda baja.",
    )

# Las pestañas separan predicción, decisión de compra, evaluación y datos fuente.
tab_forecast, tab_replenishment, tab_metrics, tab_data = st.tabs(
    ["Predicción", "Reposición", "Evaluación", "Datos"]
)

with tab_forecast:
    product_names = sorted(forecast["product_name"].unique())
    product_options = {PRODUCT_NAMES.get(name, name): name for name in product_names}
    selected_labels = st.multiselect(
        "Productos",
        list(product_options),
        default=list(product_options)[:4],
        help="Selecciona qué productos deseas comparar en el gráfico y la tabla de predicción.",
    )
    selected_products = [product_options[label] for label in selected_labels]
    filtered = forecast[forecast["product_name"].isin(selected_products)]
    filtered_display = localized_values(filtered)
    fig = px.line(
        filtered_display,
        x="date",
        y="forecast_units",
        color="product_name",
        markers=True,
        labels={"forecast_units": "Unidades pronosticadas", "date": "Fecha"},
    )
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(
        filtered_display,
        use_container_width=True,
        hide_index=True,
        column_config=FORECAST_COLUMNS,
    )

with tab_replenishment:
    # ProgressColumn colorea la cantidad con componentes propios de Streamlit.
    # Así el dashboard no depende de Matplotlib para mostrar esta tabla.
    max_order = max(int(recommendations["suggested_order_qty"].max()), 1)
    replenishment_columns = {
        "product_id": FORECAST_COLUMNS["product_id"],
        "product_name": FORECAST_COLUMNS["product_name"],
        "category": FORECAST_COLUMNS["category"],
        "stock_on_hand": st.column_config.NumberColumn(
            "Stock actual", help="Unidades disponibles físicamente en la tienda.", format="%d"
        ),
        "on_order": st.column_config.NumberColumn(
            "Pedido en camino", help="Unidades ya solicitadas que todavía no han llegado.", format="%d"
        ),
        "lead_time_days": st.column_config.NumberColumn(
            "Tiempo de entrega (días)", help="Días que tarda el proveedor en entregar.", format="%d"
        ),
        "avg_daily_forecast": st.column_config.NumberColumn(
            "Promedio diario previsto", help="Venta diaria media estimada dentro del horizonte.", format="%.2f"
        ),
        "demand_during_lead_time": st.column_config.NumberColumn(
            "Demanda durante cobertura", help="Unidades previstas durante el tiempo que debe cubrir el pedido.", format="%.2f"
        ),
        "safety_stock": st.column_config.NumberColumn(
            "Stock de seguridad", help="Reserva adicional calculada con el ajuste de días de seguridad.", format="%.2f"
        ),
        "reorder_point": st.column_config.NumberColumn(
            "Punto de pedido", help="Nivel requerido para cubrir demanda y stock de seguridad.", format="%.2f"
        ),
        "coverage_days": st.column_config.NumberColumn(
            "Cobertura actual (días)", help="Días aproximados que puede cubrir el stock disponible.", format="%.2f"
        ),
        "suggested_order_qty": st.column_config.ProgressColumn(
            "Cantidad sugerida",
            help="Unidades recomendadas, redondeadas al tamaño de empaque.",
            min_value=0,
            max_value=max_order,
            format="%d",
        ),
        "priority": st.column_config.TextColumn(
            "Prioridad",
            help="Alta: riesgo de quedarse sin stock antes de recibir el pedido. Media: se recomienda comprar sin riesgo inmediato. Baja: no se necesita compra.",
        ),
    }
    st.dataframe(
        localized_values(recommendations),
        use_container_width=True,
        hide_index=True,
        column_config=replenishment_columns,
    )
    csv = recommendations.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Descargar recomendaciones CSV",
        csv,
        "recommendations.csv",
        "text/csv",
        help="Descarga el formato técnico con nombres de columnas en inglés para conservar compatibilidad con el pipeline.",
    )

with tab_metrics:
    st.caption(
        "Esta tabla corresponde a la evaluación auxiliar de la ventana final. "
        "La selección del modelo operativo utiliza el ranking multihorizonte de 7, 14 y 30 días."
    )
    metric_display = localized_values(metrics)
    st.dataframe(
        metric_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "model": st.column_config.TextColumn(
                "Modelo",
                help="Nombre técnico internacional del algoritmo evaluado.",
            ),
            "mae": st.column_config.NumberColumn("MAE (unidades)", help="Error medio en unidades; menor es mejor.", format="%.2f"),
            "rmse": st.column_config.NumberColumn("RMSE (unidades)", help="Penaliza más los errores grandes; menor es mejor.", format="%.2f"),
            "mape": st.column_config.NumberColumn("MAPE (%)", help="Error porcentual respecto al valor real; menor es mejor.", format="%.2f %%"),
            "smape": st.column_config.NumberColumn("SMAPE (%)", help="Error porcentual simétrico y métrica principal; menor es mejor.", format="%.2f %%"),
        },
    )
    if not metrics.empty:
        chart = metric_display.melt(
            id_vars="model",
            value_vars=["mae", "smape"],
            var_name="metric_name",
            value_name="metric_value",
        )
        chart["metric_name"] = chart["metric_name"].map(
            {"mae": "MAE (unidades)", "smape": "SMAPE (%)"}
        )
        fig = px.bar(
            chart,
            x="model",
            y="metric_value",
            color="metric_name",
            barmode="group",
            labels={
                "model": "Modelo",
                "metric_value": "Valor de la métrica",
                "metric_name": "Métrica",
            },
        )
        st.plotly_chart(fig, use_container_width=True)

with tab_data:
    col1, col2 = st.columns(2)
    col1.subheader("Ventas históricas")
    col1.dataframe(
        localized_values(sales.tail(200)),
        use_container_width=True,
        hide_index=True,
        column_config=SALES_COLUMNS,
    )
    col2.subheader("Inventario y recomendaciones")
    col2.dataframe(
        localized_values(recommendations),
        use_container_width=True,
        hide_index=True,
        column_config=replenishment_columns,
    )
