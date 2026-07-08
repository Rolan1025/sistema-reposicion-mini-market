# Metodologia tecnica del prototipo

## Alcance

El prototipo implementa un flujo reproducible para un mini-market:

1. Ingesta de ventas historicas por dia y producto.
2. Limpieza y agregacion diaria.
3. Generacion de variables de calendario, rezagos y medias moviles.
4. Comparacion entre baseline estacional y modelos clasicos.
5. Seleccion del mejor modelo por SMAPE y MAE.
6. Pronostico iterativo a 7, 14 o 30 dias.
7. Recomendacion de reposicion con punto de pedido, stock de seguridad y lead time.
8. Automatizacion del entrenamiento y reporte mediante n8n.

## Dataset

El repositorio incluye un dataset sintetico para facilitar la demostracion.
Para usar datos reales anonimizados, el archivo de ventas debe tener estas columnas:

- `date`: fecha de venta.
- `product_id`: identificador del producto.
- `product_name`: nombre del producto.
- `category`: categoria.
- `units_sold`: unidades vendidas.
- `unit_price`: precio unitario.
- `promo`: opcional, 1 si hubo promocion y 0 en caso contrario.

El inventario debe incluir:

- `product_id`
- `product_name`
- `category`
- `stock_on_hand`
- `on_order`
- `lead_time_days`
- `pack_size`
- `max_stock`

## Modelos

Se comparan:

- Baseline estacional de 7 dias.
- Regresion Ridge.
- Random Forest Regressor.

Las metricas usadas son MAE, RMSE, MAPE y SMAPE. Para el informe academico,
SMAPE es util porque tolera mejor productos con demanda baja o intermitente.

## Formula de reposicion

La cantidad sugerida se calcula de forma explicable:

```text
demanda_durante_reposicion = suma del forecast hasta lead_time + periodo_revision
stock_seguridad = demanda_promedio_diaria_forecast * dias_stock_seguridad
punto_pedido = demanda_durante_reposicion + stock_seguridad
cantidad_sugerida = max(punto_pedido - stock_disponible - pedidos_en_camino, 0)
```

La cantidad se redondea al multiplo de `pack_size`.

