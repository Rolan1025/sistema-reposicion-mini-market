# Sistema de prediccion de demanda y reposicion para mini-market

Aplicacion funcional para el TFE: predice demanda por producto a corto plazo y genera recomendaciones de reposicion usando modelos clasicos de aprendizaje automatico, una interfaz Streamlit, una API FastAPI y automatizacion con n8n.

## Componentes

- Pipeline reproducible de datos en `src/demand_forecasting`.
- Dataset sintetico inicial en `data/raw` generado por script.
- Modelos comparados: baseline estacional 7 dias, Ridge y Random Forest.
- Metricas: MAE, RMSE, MAPE y SMAPE.
- Reposicion explicable: forecast, lead time, stock de seguridad, pedidos en camino y tamano de empaque.
- Dashboard Streamlit en `app/streamlit_app.py`.
- API FastAPI en `app/api.py`.
- Workflow n8n importable en `n8n/workflows/retrain_and_report.json`.

## Estructura

```text
sistema-reposicion-mini-market/
  app/                    Dashboard y API
  data/raw/               CSV de ventas e inventario
  docs/                   Notas tecnicas para el informe
  models/                 Modelo entrenado
  n8n/workflows/          Automatizacion importable en n8n
  reports/                Metricas, forecasts y recomendaciones
  scripts/                Comandos de generacion, entrenamiento y prediccion
  src/demand_forecasting/ Logica principal
  tests/                  Pruebas automatizadas
```

## Ejecucion local

Requiere Python 3.11 o superior.

```bash
cd sistema-reposicion-mini-market
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python scripts/generate_sample_data.py
python scripts/train.py
python scripts/evaluate.py
python scripts/predict.py --horizon-days 14
streamlit run app/streamlit_app.py
```

El dashboard queda disponible en `http://localhost:8501`.

Para ejecutar la API:

```bash
uvicorn app.api:app --reload --port 8000
```

Endpoints principales:

- `GET /health`
- `POST /sample-data`
- `POST /train`
- `GET /predict?horizon_days=14&safety_stock_days=3`

## Ejecucion con Docker

```bash
docker compose up --build
```

Servicios:

- `tfe_mini_market_panel_streamlit`: panel visual en `http://localhost:8501`.
- `tfe_mini_market_api_prediccion`: API y documentacion en `http://localhost:8000/docs`.
- `tfe_mini_market_automatizacion_n8n`: automatizacion en `http://localhost:5678`.

Docker Desktop agrupa estos tres contenedores bajo el nombre
`tfe-mini-market-prediccion-reposicion`, para distinguirlos de otros proyectos.

En n8n se puede importar `n8n/workflows/retrain_and_report.json`. El flujo llama a los endpoints `/train` y `/predict` del servicio API dentro de Docker.

## Datos reales

Para reemplazar el dataset sintetico, usa un CSV de ventas con columnas:

```text
date,product_id,product_name,category,units_sold,unit_price,promo
```

`promo` es opcional. El inventario debe tener:

```text
product_id,product_name,category,stock_on_hand,on_order,lead_time_days,pack_size,max_stock
```

## Evidencia para el TFE

Archivos utiles para documentar resultados:

- `reports/metrics.csv`: comparacion cuantitativa de modelos.
- `reports/metrics_by_horizon.csv`: MAE, RMSE, MAPE y SMAPE para 7, 14 y 30 dias.
- `reports/metrics_by_product.csv`: resultados por producto, modelo y horizonte.
- `reports/backtest_predictions.csv`: predicciones y valores reales usados en el backtesting.
- `reports/model_ranking.csv`: ranking promedio multihorizonte usado para seleccionar el modelo operativo.
- `reports/forecast.csv`: predicciones por fecha y producto.
- `reports/recommendations.csv`: sugerencias de reposicion.
- `docs/metodologia_tecnica.md`: resumen tecnico para trasladar al documento academico.
- `scripts/generate_execution_evidence.py`: ejecuta pruebas y evaluacion antes de regenerar su figura de evidencia.
- `scripts/update_tfe_docx.py`: prepara la version revisada de la tercera entrega.
- `scripts/build_tfe_extraordinaria.py`: construye la memoria ampliada a partir de la version revisada.

## Pruebas

```bash
pytest
```

Las pruebas cubren metricas, logica de reposicion y un flujo completo de entrenamiento/prediccion con datos sinteticos.
