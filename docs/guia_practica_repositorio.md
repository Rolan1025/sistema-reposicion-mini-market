# Guia practica del repositorio

## 1. Que hace el proyecto

El sistema toma un historico diario de ventas y una fotografia del inventario para:

1. Preparar una serie diaria por producto.
2. Crear variables de calendario, precio, promocion, retardos y ventanas moviles.
3. Comparar un baseline estacional, Ridge y Random Forest.
4. Elegir el modelo con mejor SMAPE medio en 7, 14 y 30 dias.
5. Pronosticar la demanda futura de cada producto.
6. Convertir el pronostico en una cantidad sugerida de compra.
7. Mostrar los resultados en Streamlit, FastAPI, CSV y SQLite.
8. Permitir que n8n solicite reentrenamiento y prediccion mediante la API.

El modelo operativo actual es Random Forest. El conjunto de ejemplo contiene 7.310 filas, diez productos y fechas entre 2024-01-01 y 2025-12-31.

## 2. Palabras importantes en lenguaje sencillo

- **Modelo**: regla matemática aprendida a partir de los datos para calcular una predicción.
- **Entrenar**: mostrarle al modelo ventas históricas para que aprenda patrones.
- **Predecir o pronosticar**: estimar unidades de una fecha que el modelo no ha visto.
- **Forecast**: palabra inglesa usada en el código para referirse al pronóstico.
- **Horizonte**: cantidad de días futuros que se desean predecir, por ejemplo 7, 14 o 30.
- **Baseline**: método sencillo que sirve como punto de comparación. Aquí repite las ventas de hace siete días.
- **Ridge**: regresión lineal que combina las variables y limita coeficientes exagerados.
- **Random Forest**: conjunto de muchos árboles de decisión; cada árbol propone una predicción y el resultado se promedia.
- **Feature o variable predictora**: pista entregada al modelo, como precio, día de la semana o ventas anteriores.
- **Lag o retardo**: venta pasada. `lag_7` significa las unidades vendidas siete días antes.
- **Ventana móvil**: resumen de los últimos días, como el promedio de las últimas siete ventas.
- **Backtesting**: ocultar el final del histórico, predecirlo y comparar la predicción con lo que realmente ocurrió.
- **MAE**: error absoluto medio. Indica cuántas unidades falla el modelo en promedio. MAE 6 significa aproximadamente seis unidades de diferencia por producto y día.
- **RMSE**: raíz del error cuadrático medio. También se expresa en unidades, pero castiga más los errores muy grandes.
- **MAPE**: error porcentual medio respecto a la venta real. Puede ser problemático cuando la venta real es cero o muy pequeña.
- **SMAPE**: error porcentual simétrico. Compara el error con el tamaño combinado del valor real y el pronosticado; en este proyecto es la métrica principal para ordenar modelos.
- **Lead time o tiempo de reposición**: días que tarda el proveedor en entregar un pedido.
- **Stock de seguridad**: unidades adicionales para protegerse frente a variaciones o retrasos.
- **Periodo de revisión**: días hasta que el comercio volverá a revisar el inventario.
- **Pack size o tamaño de empaque**: cantidad mínima o múltiplo en que el proveedor vende el producto.
- **SQLite**: base de datos guardada en un único archivo local.
- **API**: puerta de entrada para que otro programa solicite entrenamiento o predicciones.
- **n8n**: herramienta visual que automatiza llamadas a la API en un horario.

## 3. Flujo mental del sistema

```text
sales_sample.csv + inventory_sample.csv
                  |
                  v
       limpieza y serie diaria
                  |
                  v
        variables predictoras
                  |
                  v
 baseline + Ridge + Random Forest
                  |
                  v
     evaluacion 7, 14 y 30 dias
                  |
                  v
       demand_model.joblib
                  |
                  v
       forecast diario futuro
                  |
                  v
 regla de reposicion + inventario
                  |
                  v
 CSV / SQLite / Streamlit / API / n8n
```

## 4. Carpetas y archivos

### Entradas

- `data/raw/sales_sample.csv`: ventas historicas. Una fila por producto y dia.
- `data/raw/inventory_sample.csv`: stock actual, pedidos en camino, tiempo de entrega, empaque y stock maximo.

### Logica principal

- `src/demand_forecasting/config.py`: rutas y constantes generales.
- `src/demand_forecasting/data.py`: valida CSV, completa fechas y genera datos sinteticos.
- `src/demand_forecasting/features.py`: crea calendario, retardos y ventanas moviles.
- `src/demand_forecasting/metrics.py`: calcula errores en unidades (MAE y RMSE) y porcentajes (MAPE y SMAPE).
- `src/demand_forecasting/models.py`: configura Ridge y Random Forest y guarda/carga el modelo.
- `src/demand_forecasting/evaluation.py`: ejecuta el backtesting de 7, 14 y 30 dias.
- `src/demand_forecasting/forecasting.py`: genera predicciones futuras de forma recursiva.
- `src/demand_forecasting/replenishment.py`: transforma el forecast en una compra sugerida.
- `src/demand_forecasting/pipeline.py`: une entrenamiento y prediccion en funciones de alto nivel.
- `src/demand_forecasting/storage.py`: escribe el ultimo reporte en SQLite.

### Comandos

- `scripts/generate_sample_data.py`: regenera los dos CSV sinteticos. Sobrescribe los datos actuales.
- `scripts/train.py`: compara modelos, selecciona el operativo y guarda `models/demand_model.joblib`.
- `scripts/evaluate.py`: calcula resultados detallados por horizonte y producto.
- `scripts/predict.py`: genera forecast, recomendaciones y la base SQLite.
- `scripts/analyze_replenishment_sensitivity.py`: compara configuraciones de seguridad y revision.
- `scripts/generate_evaluation_figures.py`: reconstruye las figuras del informe.
- `scripts/generate_execution_evidence.py`: ejecuta las pruebas y la evaluacion antes de crear la figura de evidencia.
- `scripts/update_tfe_docx.py`: genera la version revisada que sirve como base documental.
- `scripts/build_tfe_extraordinaria.py`: reconstruye el Word ampliado. No es necesario para usar la aplicacion.

### Interfaces

- `app/streamlit_app.py`: dashboard visual para el usuario.
- `app/api.py`: API para que otros sistemas soliciten entrenamiento o predicciones.
- `n8n/workflows/retrain_and_report.json`: flujo importable que llama a la API.
- `Dockerfile` y `docker-compose.yml`: ejecutan dashboard, API y n8n en contenedores.

### Salidas

- `models/demand_model.joblib`: modelo seleccionado y sus metadatos.
- `reports/metrics.csv`: comparacion de la ventana simple de entrenamiento.
- `reports/metrics_by_horizon.csv`: metricas para 7, 14 y 30 dias.
- `reports/metrics_by_product.csv`: metricas desagregadas por producto.
- `reports/model_ranking.csv`: ranking usado para elegir el modelo operativo.
- `reports/backtest_predictions.csv`: valores reales y predichos de la evaluacion.
- `reports/forecast.csv`: predicciones futuras por fecha y producto.
- `reports/recommendations.csv`: cantidad sugerida y prioridad por producto.
- `reports/mini_market_reports.db`: snapshot SQLite de forecast, recomendaciones y metricas.

## 5. Preparacion en Windows PowerShell

Abrir PowerShell y ejecutar:

```powershell
cd "C:\Users\armad\Documentos\Documentos\UNIR\Trabajo de Fin de Master\TrabajoFinal\sistema-reposicion-mini-market"
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Si PowerShell no permite activar el entorno, se puede utilizar directamente:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

En los comandos siguientes se reemplaza `python` por `.\.venv\Scripts\python.exe` si no se activa el entorno.

## 6. Cuando se debe generar la data

`generate_sample_data.py` no es un comando de arranque obligatorio. Se utiliza en dos situaciones:

1. La primera vez que se prepara la demostración.
2. Cuando se quiere restaurar exactamente el escenario sintético descrito en el Word.

La semilla aleatoria es `42`. Por eso, mientras no cambien el código del generador, las fechas ni el catálogo, cada ejecución produce exactamente las mismas ventas y el mismo inventario. El CSV se sobrescribe, pero su contenido vuelve a ser igual.

Después de regenerar los datos se deben volver a ejecutar entrenamiento, evaluación y predicción para que modelo y reportes correspondan a esos CSV:

```powershell
python scripts\generate_sample_data.py
python scripts\train.py
python scripts\evaluate.py
python scripts\predict.py --horizon-days 14
```

Para abrir normalmente un proyecto que ya tiene datos y modelo, no se genera nada. Basta con iniciar Streamlit o la API.

## 7. Ejecucion desde cero por consola

La secuencia normal es:

```powershell
python scripts\generate_sample_data.py
python scripts\train.py
python scripts\evaluate.py
python scripts\predict.py --horizon-days 14 --safety-stock-days 3 --review-period-days 1
python -m pytest -q
```

No es necesario regenerar los datos en cada ejecución. El primer comando restaura el escenario sintético original y sobrescribe cualquier CSV que estuviera allí.

La ejecución normal, cuando los CSV y el modelo ya existen, puede ser solamente:

```powershell
python scripts\predict.py --horizon-days 14
```

O se puede abrir directamente el dashboard:

```powershell
python -m streamlit run app\streamlit_app.py
```

Resultados esperados:

- `train.py` informa el modelo operativo y crea el archivo `.joblib`.
- `evaluate.py` crea las metricas multihorizonte.
- `predict.py` produce 140 filas para diez productos y catorce dias, mas diez recomendaciones.
- `pytest` debe mostrar seis pruebas superadas.

## 8. Abrir el dashboard

Despues de entrenar:

```powershell
python -m streamlit run app\streamlit_app.py
```

Abrir `http://localhost:8501`. En la barra lateral se seleccionan horizonte, stock de seguridad y periodo de revision. Las pestanas muestran forecast, reposicion, evaluacion y datos.

Para detenerlo, volver a PowerShell y pulsar `Ctrl+C`.

## 9. Abrir la API

En otra terminal con el entorno activado:

```powershell
python -m uvicorn app.api:app --reload --port 8000
```

Abrir `http://localhost:8000/docs`. Desde esa pagina se pueden probar:

- `GET /health`: comprueba que la API responde.
- `POST /sample-data`: regenera los datos sinteticos.
- `POST /train?test_days=30`: entrena el modelo.
- `GET /predict?horizon_days=14&safety_stock_days=3&review_period_days=1`: devuelve forecast y recomendaciones.

## 10. Ejecucion con Docker y n8n

Con Docker Desktop iniciado:

```powershell
docker compose up --build
```

Servicios:

- `tfe_mini_market_panel_streamlit`: panel Streamlit en `http://localhost:8501`.
- `tfe_mini_market_api_prediccion`: API FastAPI en `http://localhost:8000/docs`.
- `tfe_mini_market_automatizacion_n8n`: automatizacion n8n en `http://localhost:5678`.

Docker Desktop agrupa los tres componentes como
`tfe-mini-market-prediccion-reposicion`. En n8n se importa
`n8n/workflows/retrain_and_report.json`. El flujo diario llama primero a
`http://api-prediccion:8000/train` y luego a
`http://api-prediccion:8000/predict`. El nombre `api-prediccion` solo funciona
dentro de la red de Docker Compose.

Para cerrar los servicios:

```powershell
docker compose down
```

## 11. Como se calcula la reposicion

Para cada producto:

```text
periodo_cubierto = lead_time_days + review_period_days
demanda_cubierta = suma del forecast durante ese periodo
stock_seguridad = promedio diario pronosticado * safety_stock_days
stock_disponible = stock_on_hand + on_order
necesidad = max(demanda_cubierta + stock_seguridad - stock_disponible, 0)
```

La necesidad se limita por `max_stock`, se redondea al multiplo de `pack_size` y se clasifica con prioridad alta, media o baja.

## 12. Uso de datos reales

El CSV de ventas debe conservar:

```text
date,product_id,product_name,category,units_sold,unit_price,promo
```

El CSV de inventario debe conservar:

```text
product_id,product_name,category,stock_on_hand,on_order,lead_time_days,pack_size,max_stock
```

Antes de reemplazar los archivos de ejemplo se debe guardar una copia. Despues se ejecutan nuevamente `train.py`, `evaluate.py` y `predict.py`.

## 13. Demostracion corta para la defensa

1. Mostrar los dos CSV de entrada.
2. Ejecutar `python scripts\train.py` y explicar los tres modelos.
3. Abrir `reports/model_ranking.csv` y justificar Random Forest.
4. Ejecutar `python scripts\predict.py --horizon-days 14`.
5. Mostrar que `forecast.csv` tiene 140 filas y `recommendations.csv` diez.
6. Abrir Streamlit y cambiar los dias de seguridad.
7. Abrir `/docs` de FastAPI y ejecutar `/health` y `/predict`.
8. Mostrar las seis pruebas superadas.

La idea central para explicar el proyecto es: el modelo predice unidades, pero la regla de inventario convierte esas unidades en una decision de compra comprensible.
