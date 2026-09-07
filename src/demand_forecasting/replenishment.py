"""Regla explicable que convierte predicciones en cantidades de compra."""

from __future__ import annotations

import math

import numpy as np
import pandas as pd


def round_to_pack(quantity: float, pack_size: float) -> int:
    """Redondea hacia arriba al empaque completo que vende el proveedor."""
    if quantity <= 0:
        return 0
    pack_size = max(int(pack_size), 1)
    return int(math.ceil(quantity / pack_size) * pack_size)


def recommend_replenishment(
    forecast: pd.DataFrame,
    inventory: pd.DataFrame,
    safety_stock_days: int = 3,
    review_period_days: int = 1,
) -> pd.DataFrame:
    """Calcula una recomendación de reposición para cada producto.

    Considera demanda prevista, tiempo de entrega, revisión, stock disponible,
    pedidos en camino, stock de seguridad, empaque y máximo permitido.
    """
    forecast = forecast.copy()
    inventory = inventory.copy()
    forecast["date"] = pd.to_datetime(forecast["date"])

    recommendations = []
    for item in inventory.itertuples(index=False):
        item_forecast = forecast[forecast["product_id"] == item.product_id].sort_values("date")
        if item_forecast.empty:
            continue

        # Se debe cubrir lo que tarda el proveedor más el tiempo hasta la
        # siguiente revisión de inventario.
        lead_time = int(getattr(item, "lead_time_days", 1))
        horizon = max(lead_time + review_period_days, 1)
        demand_window = item_forecast.head(horizon)
        avg_daily = float(item_forecast["forecast_units"].mean())
        demand_during_lead_time = float(demand_window["forecast_units"].sum())
        safety_stock = avg_daily * safety_stock_days
        reorder_point = demand_during_lead_time + safety_stock
        available_stock = float(item.stock_on_hand) + float(item.on_order)
        # Nunca se recomienda una compra negativa.
        raw_quantity = max(reorder_point - available_stock, 0)

        max_stock = getattr(item, "max_stock", np.nan)
        if not pd.isna(max_stock):
            raw_quantity = min(raw_quantity, max(float(max_stock) - available_stock, 0))

        suggested_qty = round_to_pack(raw_quantity, getattr(item, "pack_size", 1))
        coverage_days = available_stock / avg_daily if avg_daily > 0 else np.inf
        # Riesgo alto cuando el stock no alcanza hasta que llegue el proveedor.
        stockout_risk = coverage_days < lead_time
        priority = "alta" if stockout_risk else "media" if suggested_qty > 0 else "baja"

        recommendations.append(
            {
                "product_id": item.product_id,
                "product_name": item.product_name,
                "category": getattr(item, "category", ""),
                "stock_on_hand": int(item.stock_on_hand),
                "on_order": int(item.on_order),
                "lead_time_days": lead_time,
                "avg_daily_forecast": round(avg_daily, 2),
                "demand_during_lead_time": round(demand_during_lead_time, 2),
                "safety_stock": round(safety_stock, 2),
                "reorder_point": round(reorder_point, 2),
                "coverage_days": round(float(coverage_days), 2) if np.isfinite(coverage_days) else 999,
                "suggested_order_qty": suggested_qty,
                "priority": priority,
            }
        )

    return pd.DataFrame(recommendations).sort_values(
        ["priority", "suggested_order_qty"],
        ascending=[True, False],
        key=lambda col: col.map({"alta": 0, "media": 1, "baja": 2}) if col.name == "priority" else col,
    )
