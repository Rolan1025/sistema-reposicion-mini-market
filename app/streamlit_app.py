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
from demand_forecasting.pipeline import predict_pipeline, train_pipeline


st.set_page_config(page_title="Reposicion mini-market", layout="wide")

st.title("Sistema de prediccion de demanda y reposicion")

with st.sidebar:
    st.header("Parametros")
    horizon_days = st.selectbox("Horizonte de prediccion", [7, 14, 30], index=1)
    safety_stock_days = st.slider("Stock de seguridad (dias)", min_value=0, max_value=14, value=3)
    review_period_days = st.slider("Periodo de revision (dias)", min_value=1, max_value=7, value=1)
    test_days = st.slider("Ventana de evaluacion (dias)", min_value=14, max_value=90, value=30)

    if st.button("Generar datos de ejemplo"):
        write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
        st.success("Datos de ejemplo generados.")

    train_clicked = st.button("Reentrenar modelo", type="primary")


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

metric_cols = st.columns(4)
if not metrics.empty:
    best = metrics.iloc[0]
    metric_cols[0].metric("Mejor modelo", str(best["model"]))
    metric_cols[1].metric("MAE", f"{best['mae']:.2f}")
    metric_cols[2].metric("SMAPE", f"{best['smape']:.2f}%")
    metric_cols[3].metric("MAPE", f"{best['mape']:.2f}%")

tab_forecast, tab_replenishment, tab_metrics, tab_data = st.tabs(
    ["Prediccion", "Reposicion", "Evaluacion", "Datos"]
)

with tab_forecast:
    product_names = sorted(forecast["product_name"].unique())
    selected_products = st.multiselect("Productos", product_names, default=product_names[:4])
    filtered = forecast[forecast["product_name"].isin(selected_products)]
    fig = px.line(
        filtered,
        x="date",
        y="forecast_units",
        color="product_name",
        markers=True,
        labels={"forecast_units": "Unidades pronosticadas", "date": "Fecha"},
    )
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(filtered, use_container_width=True)

with tab_replenishment:
    st.dataframe(
        recommendations.style.background_gradient(subset=["suggested_order_qty"], cmap="YlOrRd"),
        use_container_width=True,
    )
    csv = recommendations.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar recomendaciones CSV", csv, "recommendations.csv", "text/csv")

with tab_metrics:
    st.dataframe(metrics, use_container_width=True)
    if not metrics.empty:
        fig = px.bar(metrics, x="model", y=["mae", "smape"], barmode="group")
        st.plotly_chart(fig, use_container_width=True)

with tab_data:
    col1, col2 = st.columns(2)
    col1.subheader("Ventas historicas")
    col1.dataframe(sales.tail(200), use_container_width=True)
    col2.subheader("Inventario y recomendaciones")
    col2.dataframe(recommendations, use_container_width=True)

