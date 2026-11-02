# %% [markdown]
# # 04b · Diagnóstico del criterio C1 (disponibilidad de productos vitales)
# Descompone la disponibilidad ante demanda de los productos vitales por patrón de frecuencia, para explicar
# por qué la propuesta no alcanza el umbral. No cambia los umbrales del protocolo; solo documenta la causa.

# %%
import os, sys, json
import numpy as np
import pandas as pd

sys.path.insert(0, "src")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import inventario as inv

DATA_DIR = os.environ.get("DATA_DIR", "data")
BASE = os.environ.get("BASE_ANALITICA", f"{DATA_DIR}/base_analitica.parquet")
TAB = os.environ.get("OUT_TAB", "outputs/tablas")
W, L, R = 12, 0.5, 0.5

base = inv.leer_base_analitica(BASE)
biblia = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
agr = inv.agregar_base(base)
periodos = sorted(agr["PERIODO"].unique())
clas = pd.read_csv(f"{TAB}/clasificacion_abc_ved_por_producto.csv", dtype={"CODIGO_PROPIO": str}).set_index("CODIGO_PROPIO")
val = inv.asignar_costo(agr, biblia)
con = val.dropna(subset=["COSTO_UNITARIO"])
costo_prod = con.groupby("CODIGO_PROPIO")["VALOR"].sum() / con.groupby("CODIGO_PROPIO")["CANTIDAD"].sum()
M_df = inv.matriz_demanda(agr, ["CODIGO_PROPIO"], periodos)
M = M_df.to_numpy(dtype=float)
ids = M_df.index
ved = clas["VED"].reindex(ids).fillna("E").to_numpy()
costo = np.nan_to_num(costo_prod.reindex(ids).to_numpy(dtype=float), nan=0.0)

# %%
filas = []
n, T = M.shape
for t in range(W - 1, T - 1):
    win, D = M[:, t - W + 1: t + 1], M[:, t + 1]
    abc = inv.clasificar_abc(pd.Series(win.sum(axis=1) * costo)).to_numpy()
    p = inv.politica_propuesta(win, abc, ved, L, R)
    S = p["maximo"]
    m = (ved == "V") & (D > 0)
    sup_max_ventana = D > win.max(axis=1)           # demanda mayor que cualquier mes de la ventana
    for pat in np.unique(p["patron"][m]):
        mm = m & (p["patron"] == pat)
        filas.append({"origen": t, "PATRON": pat, "series_con_demanda": int(mm.sum()),
                      "series_cubiertas": int((mm & (S >= D)).sum()),
                      "faltantes": int((mm & (S < D)).sum()),
                      "faltantes_con_demanda_sobre_maximo_de_ventana": int((mm & (S < D) & sup_max_ventana).sum()),
                      "faltantes_con_nivel_cero": int((mm & (S < D) & (S == 0)).sum())})
d = pd.DataFrame(filas).groupby("PATRON").sum().drop(columns="origen")
d["disponibilidad_%"] = (100 * d["series_cubiertas"] / d["series_con_demanda"]).round(2)
d["%_de_los_faltantes_vitales"] = (100 * d["faltantes"] / d["faltantes"].sum()).round(2)
d["%_faltantes_con_demanda_sobre_maximo_ventana"] = (100 * d["faltantes_con_demanda_sobre_maximo_de_ventana"] / d["faltantes"]).round(2)
print(d.to_string())
d.to_csv(f"{TAB}/o3_diagnostico_vitales.csv", encoding="utf-8-sig")
