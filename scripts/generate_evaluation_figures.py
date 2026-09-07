from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
FIGURES = REPORTS / "figures"
COLORS = {
    "ridge": "#0098CD",
    "random_forest": "#4C8C2B",
    "seasonal_naive_7d": "#8C8C8C",
}
DISPLAY_NAMES = {
    "Atun lata": "Atún lata",
    "Cafe molido 250g": "Café molido 250g",
    "Papel higienico x4": "Papel higiénico x4",
}


def save(fig: plt.Figure, name: str) -> None:
    FIGURES.mkdir(parents=True, exist_ok=True)
    fig.savefig(FIGURES / name, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def metrics_by_horizon() -> None:
    metrics = pd.read_csv(REPORTS / "metrics_by_horizon.csv")
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
    for model, group in metrics.groupby("model"):
        group = group.sort_values("horizon_days")
        label = model.replace("_", " ").title()
        axes[0].plot(group["horizon_days"], group["mae"], marker="o", linewidth=2.2, label=label, color=COLORS[model])
        axes[1].plot(group["horizon_days"], group["smape"], marker="o", linewidth=2.2, label=label, color=COLORS[model])
    axes[0].set_title("MAE por horizonte")
    axes[0].set_ylabel("Unidades")
    axes[1].set_title("SMAPE por horizonte")
    axes[1].set_ylabel("Porcentaje")
    for ax in axes:
        ax.set_xlabel("Horizonte (días)")
        ax.set_xticks([7, 14, 30])
        ax.grid(alpha=0.25)
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="lower center", ncol=3, frameon=False)
    fig.suptitle("Comparación multihorizonte de modelos", fontsize=14, fontweight="bold")
    fig.tight_layout(rect=(0, 0.12, 1, 0.93))
    save(fig, "metricas_por_horizonte.png")


def product_heatmap() -> None:
    data = pd.read_csv(REPORTS / "metrics_by_product.csv")
    data = data[data["model"] == "random_forest"]
    data["product_name"] = data["product_name"].replace(DISPLAY_NAMES)
    pivot = data.pivot(index="product_name", columns="horizon_days", values="smape").sort_index()
    fig, ax = plt.subplots(figsize=(8.5, 5.8))
    image = ax.imshow(pivot.values, cmap="YlOrRd", aspect="auto")
    ax.set_xticks(range(len(pivot.columns)), [f"{int(x)} días" for x in pivot.columns])
    ax.set_yticks(range(len(pivot.index)), pivot.index)
    for row in range(pivot.shape[0]):
        for col in range(pivot.shape[1]):
            value = pivot.iloc[row, col]
            ax.text(col, row, f"{value:.1f}%", ha="center", va="center", fontsize=8, color="black")
    ax.set_title("SMAPE de Random Forest por producto y horizonte", fontweight="bold")
    fig.colorbar(image, ax=ax, label="SMAPE (%)", fraction=0.03, pad=0.03)
    fig.tight_layout()
    save(fig, "smape_por_producto.png")


def forecast_example() -> None:
    sales = pd.read_csv(ROOT / "data" / "raw" / "sales_sample.csv", parse_dates=["date"])
    forecast = pd.read_csv(REPORTS / "forecast.csv", parse_dates=["date"])
    history = sales[sales["product_id"] == "P001"].sort_values("date").tail(30)
    future = forecast[forecast["product_id"] == "P001"].sort_values("date")
    fig, ax = plt.subplots(figsize=(10.5, 4.2))
    ax.plot(history["date"], history["units_sold"], marker="o", markersize=3, label="Demanda histórica", color="#1F1F1F")
    ax.plot(future["date"], future["forecast_units"], marker="o", label="Predicción Random Forest", color="#4C8C2B")
    ax.axvline(history["date"].max(), color="#8C8C8C", linestyle="--", linewidth=1.2, label="Inicio del pronóstico")
    ax.set_title("Pronóstico operativo: Leche entera 1L, horizonte 14 días", fontweight="bold")
    ax.set_ylabel("Unidades diarias")
    ax.grid(alpha=0.25)
    ax.legend(frameon=False)
    fig.autofmt_xdate()
    fig.tight_layout()
    save(fig, "forecast_producto_ejemplo.png")


def replenishment_bar() -> None:
    rec = pd.read_csv(REPORTS / "recommendations.csv").sort_values("suggested_order_qty")
    rec["product_name"] = rec["product_name"].replace(DISPLAY_NAMES)
    colors = rec["priority"].map({"alta": "#C00000", "media": "#E6A700", "baja": "#4C8C2B"}).fillna("#8C8C8C")
    fig, ax = plt.subplots(figsize=(9.5, 5.2))
    ax.barh(rec["product_name"], rec["suggested_order_qty"], color=colors)
    for index, value in enumerate(rec["suggested_order_qty"]):
        ax.text(value + 2, index, str(int(value)), va="center", fontsize=8)
    ax.set_title("Cantidad de reposición sugerida por producto", fontweight="bold")
    ax.set_xlabel("Unidades a solicitar")
    ax.grid(axis="x", alpha=0.2)
    fig.tight_layout()
    save(fig, "reposicion_sugerida.png")


def dataset_history() -> None:
    sales = pd.read_csv(ROOT / "data" / "raw" / "sales_sample.csv", parse_dates=["date"])
    monthly = sales.groupby([pd.Grouper(key="date", freq="MS"), "category"], as_index=False)["units_sold"].sum()
    fig, ax = plt.subplots(figsize=(10.5, 4.8))
    for category, group in monthly.groupby("category"):
        ax.plot(group["date"], group["units_sold"], linewidth=1.8, label=category)
    ax.set_title("Evolución mensual de unidades vendidas por categoría", fontweight="bold")
    ax.set_ylabel("Unidades mensuales")
    ax.grid(alpha=0.2)
    ax.legend(ncol=4, frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.18))
    fig.autofmt_xdate()
    fig.tight_layout()
    save(fig, "ventas_mensuales_categoria.png")


def main() -> None:
    metrics_by_horizon()
    product_heatmap()
    forecast_example()
    replenishment_bar()
    dataset_history()
    print(f"Figures written to: {FIGURES}")


if __name__ == "__main__":
    main()
