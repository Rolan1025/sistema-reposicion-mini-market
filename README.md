# Sistema de predicción de demanda y reposición para mini-market

Prototipo funcional desarrollado para el TFE. El sistema utiliza ventas históricas para predecir la demanda diaria por producto y transformar esas predicciones en recomendaciones explicables de reposición de inventario.

## Qué hace el sistema

1. Lee las ventas históricas y el inventario actual.
2. Limpia los datos y crea una serie diaria por producto.
3. Genera variables de calendario, precio, promoción, retardos y medias móviles.
4. Compara un modelo base estacional, Ridge Regression y Random Forest.
5. Evalúa los modelos en horizontes de 7, 14 y 30 días.
6. Selecciona el modelo operativo mediante su desempeño multihorizonte.
7. Pronostica la demanda futura.
8. Calcula cuánto pedir considerando inventario, pedidos en camino, tiempo de entrega, stock de seguridad y tamaño de empaque.
9. Presenta los resultados en Streamlit, FastAPI, archivos CSV y SQLite.
10. Permite automatizar el reentrenamiento y la predicción mediante n8n.

El conjunto de ejemplo es sintético y reproducible. Contiene diez productos y 7.310 registros diarios entre el 1 de enero de 2024 y el 31 de diciembre de 2025.

## Componentes

- **Streamlit:** panel visual para consultar predicciones, recomendaciones y métricas.
- **FastAPI:** API para integrar el entrenamiento y la predicción con otros sistemas.
- **n8n:** automatización visual que llama a la API en un horario programado.
- **scikit-learn:** entrenamiento y comparación de los modelos de aprendizaje automático.
- **SQLite y CSV:** almacenamiento local de resultados y evidencias.
- **Docker Compose:** ejecución conjunta de Streamlit, FastAPI y n8n.

## Conceptos y métricas

- **Baseline estacional:** referencia sencilla que repite la demanda observada siete días antes.
- **Ridge Regression:** modelo lineal regularizado que limita coeficientes excesivos.
- **Random Forest:** conjunto de árboles de decisión cuyas predicciones se promedian.
- **Horizonte:** número de días futuros que se desean predecir.
- **MAE (error absoluto medio):** unidades que falla el modelo en promedio. Un valor menor es mejor.
- **RMSE (raíz del error cuadrático medio):** error en unidades que penaliza con más fuerza las diferencias grandes. Un valor menor es mejor.
- **MAPE (error porcentual absoluto medio):** error porcentual respecto a la demanda real; puede ser inestable cuando la demanda es cero o muy baja.
- **SMAPE (error porcentual absoluto medio simétrico):** error porcentual que considera conjuntamente el valor real y el pronosticado. Es la métrica principal del proyecto.
- **Lead time o tiempo de reposición:** días que tarda el proveedor en entregar un pedido.
- **Stock de seguridad:** reserva adicional para cubrir variaciones de demanda o retrasos.
- **Periodo de revisión:** días que deben pasar antes de volver a revisar el inventario.

## Estructura del repositorio

```text
sistema-reposicion-mini-market/
  app/                    Panel Streamlit y API FastAPI
  data/raw/               Ventas e inventario de ejemplo
  models/                 Modelo entrenado, generado localmente
  n8n/workflows/          Flujo importable en n8n
  reports/                Resultados, base SQLite y figuras del TFE
  scripts/                Entrenamiento, evaluación, predicción y evidencias
  src/demand_forecasting/ Lógica principal del sistema
  tests/                  Pruebas automatizadas
  Dockerfile              Imagen de la aplicación Python
  docker-compose.yml      Ejecución de los tres servicios
```

## Inicio rápido con Docker

Requisito: Docker Desktop debe estar abierto y mostrar el motor en ejecución.

Desde la raíz del repositorio:

```powershell
docker compose up -d --build
docker compose ps
```

La primera ejecución puede tardar mientras se construye la imagen. Si no existe un modelo local, Streamlit genera los datos de ejemplo, entrena el modelo y guarda los resultados automáticamente.

Servicios disponibles:

- Panel Streamlit: `http://localhost:8501`
- API y documentación interactiva: `http://localhost:8000/docs`
- n8n: `http://localhost:5678`

Los contenedores aparecen en Docker Desktop dentro del grupo `tfe-mini-market-prediccion-reposicion` con nombres descriptivos:

- `tfe_mini_market_panel_streamlit`
- `tfe_mini_market_api_prediccion`
- `tfe_mini_market_automatizacion_n8n`

Comandos de administración:

```powershell
# Consultar el estado
docker compose ps

# Consultar los registros
docker compose logs -f

# Reiniciar los servicios
docker compose restart

# Detenerlos sin borrar datos
docker compose down
```

No es necesario pulsar **Run** por separado en cada contenedor. El comando `docker compose up -d` inicia el conjunto completo.

## Configuración de n8n

En el primer acceso a `http://localhost:5678`, n8n solicita crear la cuenta propietaria local. Sus datos se guardan en `n8n/.n8n`, una carpeta ignorada por Git.

Para incorporar la automatización:

1. Abrir n8n.
2. Seleccionar la opción para importar un flujo desde archivo.
3. Elegir `n8n/workflows/retrain_and_report.json`.
4. Guardar el flujo y ejecutarlo manualmente una vez para comprobarlo.
5. Activarlo únicamente cuando se desee la ejecución diaria.

El flujo llama primero a `http://api-prediccion:8000/train` y después a `http://api-prediccion:8000/predict`. El nombre `api-prediccion` funciona dentro de la red creada por Docker Compose.

n8n puede mostrar en sus registros un aviso sobre la ausencia del ejecutor Python interno. No afecta este proyecto porque el flujo utiliza únicamente el disparador de horario y nodos HTTP, no nodos de código Python. Para una instalación productiva con nodos de código se recomienda configurar el ejecutor externo indicado por n8n.

## Ejecución local con Python

Requiere Python 3.11 o superior.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Para reproducir todo el experimento:

```powershell
python scripts\generate_sample_data.py
python scripts\train.py
python scripts\evaluate.py
python scripts\predict.py --horizon-days 14 --safety-stock-days 3 --review-period-days 1
python -m pytest -q
```

`generate_sample_data.py` no debe ejecutarse cada vez que se abre el proyecto. Solo se utiliza para crear por primera vez o restaurar el escenario sintético. La semilla aleatoria es 42, por lo que genera los mismos datos mientras no cambien el catálogo, las fechas o el algoritmo de generación.

Después de entrenar, se pueden iniciar las interfaces por separado:

```powershell
python -m streamlit run app\streamlit_app.py
python -m uvicorn app.api:app --reload --port 8000
```

## API

Endpoints principales:

- `GET /health`: comprueba que la API responde.
- `POST /sample-data`: restaura los datos sintéticos de ejemplo.
- `POST /train?test_days=30`: reentrena y compara los modelos.
- `GET /predict?horizon_days=14&safety_stock_days=3&review_period_days=1`: genera predicciones y recomendaciones.

## Datos de entrada

El archivo de ventas debe contener:

```text
date,product_id,product_name,category,units_sold,unit_price,promo
```

`promo` es opcional. El archivo de inventario debe contener:

```text
product_id,product_name,category,stock_on_hand,on_order,lead_time_days,pack_size,max_stock
```

Para trabajar con datos reales se reemplazan los dos CSV de `data/raw` y se vuelven a ejecutar el entrenamiento, la evaluación y la predicción.

## Cálculo de reposición

Para cada producto se aplica la siguiente lógica:

```text
periodo_cubierto = tiempo_reposicion + periodo_revision
demanda_cubierta = suma del pronostico durante el periodo_cubierto
stock_seguridad = promedio_diario_pronosticado * dias_stock_seguridad
stock_disponible = stock_actual + pedidos_en_camino
necesidad = max(demanda_cubierta + stock_seguridad - stock_disponible, 0)
```

La necesidad se limita mediante el stock máximo, se redondea al múltiplo del tamaño de empaque y se clasifica con prioridad alta, media o baja.

## Resultados y evidencias

Los principales archivos generados son:

- `models/demand_model.joblib`: modelo operativo y metadatos.
- `reports/metrics.csv`: comparación auxiliar en la ventana final.
- `reports/metrics_by_horizon.csv`: métricas para 7, 14 y 30 días.
- `reports/metrics_by_product.csv`: métricas por producto, modelo y horizonte.
- `reports/model_ranking.csv`: clasificación multihorizonte de modelos.
- `reports/backtest_predictions.csv`: valores reales y pronosticados del backtesting.
- `reports/forecast.csv`: predicciones diarias futuras.
- `reports/recommendations.csv`: cantidades sugeridas de compra.
- `reports/mini_market_reports.db`: resultados almacenados en SQLite.

Los CSV, la base SQLite y el modelo son salidas reproducibles y están ignorados por Git. Las figuras de `reports/figures` sí están versionadas porque forman parte de las evidencias del TFE.

Scripts de análisis que se conservan por reproducibilidad:

- `scripts/analyze_replenishment_sensitivity.py`: compara configuraciones de inventario.
- `scripts/generate_evaluation_figures.py`: reconstruye las figuras de evaluación.
- `scripts/generate_execution_evidence.py`: ejecuta pruebas y genera la evidencia de su resultado.

## Pruebas

```powershell
python -m pytest -q
```

Las pruebas cubren las métricas, la evaluación multihorizonte, la lógica de reposición y el flujo completo de entrenamiento y predicción.

## Demostración breve para la sustentación

1. Mostrar los CSV de ventas e inventario.
2. Explicar los tres modelos y las cuatro métricas.
3. Mostrar `reports/model_ranking.csv` y justificar el modelo seleccionado.
4. Abrir Streamlit y cambiar el horizonte o el stock de seguridad.
5. Abrir `/docs` de FastAPI y ejecutar `/health` y `/predict`.
6. Ejecutar el flujo de n8n.
7. Mostrar que las pruebas automatizadas terminan correctamente.
