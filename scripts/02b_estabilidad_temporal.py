# %% [markdown]
# # 02b · Estabilidad temporal del consumo (ajuste al ACA 1)
# La retroalimentación pide comprobar si la caída del consumo de 2025 y 2026 frente a 2024
# fue generalizada o se concentró en familias específicas, porque usar un promedio histórico
# sobre una tendencia a la baja puede generar sobreinventario.

# %%
import os, sys, json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, "src")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import inventario as inv

DATA_DIR = os.environ.get("DATA_DIR", "data")
BASE = os.environ.get("BASE_ANALITICA", f"{DATA_DIR}/base_analitica.parquet")
TAB, FIG = os.environ.get("OUT_TAB", "outputs/tablas"), os.environ.get("OUT_FIG", "outputs/figuras")
os.makedirs(TAB, exist_ok=True); os.makedirs(FIG, exist_ok=True)
plt.rcParams.update({"figure.dpi": 130, "axes.spines.top": False, "axes.spines.right": False})

base = inv.leer_base_analitica(BASE)
biblia = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
agr = inv.agregar_base(base)
meta = biblia.drop_duplicates("CODIGO_PROPIO")[["CODIGO_PROPIO", "CONCEPTO", "VEN"]]
ene_may = lambda a: [f"{a}-{m:02d}" for m in range(1, 6)]
res = {}

# %% [markdown]
# ## 1. ¿Cuántos productos bajan y cuánto de la caída explican los que más bajan?

# %%
for a, b in [("2024", "2025"), ("2025", "2026")]:
    res[f"concentracion_ene_may_{a}_vs_{b}"] = inv.concentracion_caida(agr, ene_may(a), ene_may(b))
print(json.dumps(res, indent=2, ensure_ascii=False, default=str))

# %% [markdown]
# ## 2. Caída por tipo de producto (CONCEPTO) y por servicio farmacéutico

# %%
for a, b in [("2024", "2025"), ("2025", "2026")]:
    t1 = inv.descomposicion_caida(agr, meta, "CONCEPTO", ene_may(a), ene_may(b))
    t2 = inv.descomposicion_caida(agr, meta, "SERVICIO_FARMACEUTICO", ene_may(a), ene_may(b))
    t1.round(2).to_csv(f"{TAB}/est_caida_concepto_{a}_{b}.csv", encoding="utf-8-sig")
    t2.round(2).to_csv(f"{TAB}/est_caida_servicio_{a}_{b}.csv", encoding="utf-8-sig")

# %% [markdown]
# ## 3. Serie mensual por tipo de producto (índice base = promedio de 2024)

# %%
mens = (agr.merge(meta, on="CODIGO_PROPIO", how="left").assign(CONCEPTO=lambda d: d["CONCEPTO"].fillna("Sin Biblia"))
        .groupby(["PERIODO", "CONCEPTO"])["CANTIDAD"].sum().unstack(fill_value=0))
idx = mens / mens.loc[[p for p in mens.index if p.startswith("2024")]].mean() * 100
fig, ax = plt.subplots(figsize=(8, 3.8))
idx.plot(ax=ax, lw=1.6)
ax.set_ylabel("Índice (promedio mensual 2024 = 100)"); ax.set_xlabel("")
ax.tick_params(axis="x", rotation=60); ax.legend(fontsize=8, title=None)
ax.set_title("Consumo mensual por tipo de producto")
fig.tight_layout(); fig.savefig(f"{FIG}/est_indice_por_concepto.png"); plt.close(fig)
idx.round(1).to_csv(f"{TAB}/est_indice_por_concepto.csv", encoding="utf-8-sig")

with open(f"{TAB}/est_cifras_clave.json", "w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False, indent=2, default=str)
