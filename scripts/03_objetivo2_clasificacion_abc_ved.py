# %% [markdown]
# # 03 · Objetivo específico 2: clasificación ABC, VED y matriz ABC-VED
# Entrada: `base_analitica.parquet` (exportada del cuaderno del Objetivo 1), `Biblia (1).xlsx`,
# `Novedades.xlsx`. Salida: tablas en `outputs/tablas` y figuras en `outputs/figuras`.

# %%
import os, sys, json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import spearmanr

sys.path.insert(0, "src")
try:
    sys.stdout.reconfigure(encoding="utf-8")   # consola de Windows
except Exception:
    pass
import inventario as inv

DATA_DIR = os.environ.get("DATA_DIR", "data")
BASE = os.environ.get("BASE_ANALITICA", f"{DATA_DIR}/base_analitica.parquet")
TAB, FIG = os.environ.get("OUT_TAB", "outputs/tablas"), os.environ.get("OUT_FIG", "outputs/figuras")
os.makedirs(TAB, exist_ok=True); os.makedirs(FIG, exist_ok=True)
plt.rcParams.update({"figure.dpi": 130, "axes.spines.top": False, "axes.spines.right": False})

VENTANA_ABC = 12            # meses más recientes usados para el ABC principal (ver bitácora D-12)
res = {}                    # cifras clave que se reutilizan al redactar el documento

# %% [markdown]
# ## 1. Carga, agregación y asignación de costo (8.2.1)

# %%
base = inv.leer_base_analitica(BASE)
biblia = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
nov = inv.leer_novedades(f"{DATA_DIR}/Novedades.xlsx")
agr = inv.agregar_base(base)
periodos = sorted(agr["PERIODO"].unique())
ult = periodos[-VENTANA_ABC:]
res["periodos"] = [periodos[0], periodos[-1], len(periodos)]
res["ventana_abc"] = [ult[0], ult[-1]]
desc = base.drop_duplicates("CODIGO_PROPIO").set_index("CODIGO_PROPIO")["DES_MEDICAMENTO"]

val = inv.asignar_costo(agr, biblia)
cob = (val.groupby("FUENTE_COSTO").agg(filas=("CANTIDAD", "size"), unidades=("CANTIDAD", "sum"))
          .assign(pct_filas=lambda x: 100 * x["filas"] / x["filas"].sum(),
                  pct_unidades=lambda x: 100 * x["unidades"] / x["unidades"].sum()).round(2))
cob.to_csv(f"{TAB}/o2_cobertura_costo.csv", encoding="utf-8-sig")
sin_costo = (val[val["FUENTE_COSTO"] == "4_sin_costo"].groupby("CODIGO_PROPIO")["CANTIDAD"].sum()
             .sort_values(ascending=False).rename("unidades").to_frame()
             .assign(DES_MEDICAMENTO=lambda x: x.index.map(desc)))
sin_costo.to_csv(f"{TAB}/o2_productos_sin_costo.csv", encoding="utf-8-sig")
res["productos_sin_costo"] = int(len(sin_costo))
res["productos_total"] = int(val["CODIGO_PROPIO"].nunique())
print(cob)

# %% [markdown]
# ## 2. ABC por impacto económico (8.2.2)

# %%
def valor_producto(v, periodos_):
    return v[v["PERIODO"].isin(periodos_)].groupby("CODIGO_PROPIO")["VALOR"].sum()

v12, v29 = valor_producto(val, ult), valor_producto(val, periodos)
todos = pd.Index(val["CODIGO_PROPIO"].unique(), name="CODIGO_PROPIO")
v12, v29 = v12.reindex(todos).fillna(0), v29.reindex(todos).fillna(0)
abc12, abc29 = inv.clasificar_abc(v12), inv.clasificar_abc(v29)

def resumen_abc(abc, v):
    t = pd.DataFrame({"ABC": abc, "VALOR": v})
    t = t[t["ABC"].isin(list("ABC"))]
    r = t.groupby("ABC").agg(productos=("VALOR", "size"), valor=("VALOR", "sum"))
    r["pct_productos"] = 100 * r["productos"] / r["productos"].sum()
    r["pct_valor"] = 100 * r["valor"] / r["valor"].sum()
    return r.round(2)

r12 = resumen_abc(abc12, v12); r12.to_csv(f"{TAB}/o2_abc_resumen_12m.csv", encoding="utf-8-sig")
resumen_abc(abc29, v29).to_csv(f"{TAB}/o2_abc_resumen_29m.csv", encoding="utf-8-sig")
res["abc12"] = r12.to_dict()
res["productos_sin_valor_12m"] = int((abc12 == "SIN_VALOR").sum())
print(r12)

# Curva de Pareto
s = v12[v12 > 0].sort_values(ascending=False)
x = np.arange(1, len(s) + 1) / len(s) * 100
y = s.cumsum() / s.sum() * 100
fig, ax = plt.subplots(figsize=(6.2, 4))
ax.plot(x, y, color="#3b6ea5", lw=2)
for corte, c in zip(inv.CORTES_ABC, ["#c44e52", "#dd8452"]):
    ax.axhline(corte * 100, color=c, ls="--", lw=1)
ax.set_xlabel("% acumulado de productos"); ax.set_ylabel("% acumulado del valor de consumo")
ax.set_title(f"Curva de Pareto del valor de consumo ({ult[0]} a {ult[-1]})")
fig.tight_layout(); fig.savefig(f"{FIG}/o2_pareto.png"); plt.close(fig)

# %% [markdown]
# ## 3. VED por criticidad clínica (8.2.3)

# %%
ved = inv.tabla_ved(todos.to_series(), biblia, nov).set_index("CODIGO_PROPIO")
ved_res = (ved.reindex(todos)
           .assign(VALOR12=v12.values)
           .groupby("VED").agg(productos=("VED", "size"), valor=("VALOR12", "sum")))
ved_res["pct_productos"] = 100 * ved_res["productos"] / ved_res["productos"].sum()
ved_res["pct_valor"] = 100 * ved_res["valor"] / ved_res["valor"].sum()
ved_res.round(2).to_csv(f"{TAB}/o2_ved_resumen.csv", encoding="utf-8-sig")
fuente_ved = ved["FUENTE_VED"].value_counts()
fuente_ved.to_csv(f"{TAB}/o2_ved_fuente.csv", encoding="utf-8-sig")
res["ved_fuente"] = fuente_ved.to_dict()
print(ved_res.round(2))

# VED por tipo de producto
tipo = biblia.drop_duplicates("CODIGO_PROPIO").set_index("CODIGO_PROPIO")["CONCEPTO"]
(pd.crosstab(todos.map(tipo).fillna("Sin Biblia"), ved.reindex(todos)["VED"].values, normalize="index") * 100
 ).round(2).to_csv(f"{TAB}/o2_ved_por_concepto.csv", encoding="utf-8-sig")

# %% [markdown]
# ## 4. Matriz ABC-VED (8.2.4)

# %%
tabla = pd.DataFrame({"ABC12": abc12, "ABC29": abc29, "VED": ved.reindex(todos)["VED"].values,
                      "VALOR_12M": v12, "VALOR_29M": v29})
mat = inv.tabla_abc_ved(tabla["ABC12"], tabla["VED"], tabla["VALOR_12M"])
for k, t in mat.items():
    t.to_csv(f"{TAB}/o2_matriz_{k}.csv", encoding="utf-8-sig")
    print(k); print(t)

fig, ax = plt.subplots(1, 2, figsize=(9, 3.6))
for a, k, tt in zip(ax, ["pct_productos", "pct_valor"], ["% de productos", "% del valor de consumo"]):
    im = a.imshow(mat[k].values, cmap="Blues")
    a.set_xticks(range(3), list("VED")); a.set_yticks(range(3), list("ABC"))
    a.set_xlabel("Criticidad (VED)"); a.set_ylabel("Impacto económico (ABC)"); a.set_title(tt)
    for i in range(3):
        for j in range(3):
            a.text(j, i, f"{mat[k].values[i, j]:.1f}", ha="center", va="center",
                   color="white" if mat[k].values[i, j] > mat[k].values.max() / 2 else "black")
fig.tight_layout(); fig.savefig(f"{FIG}/o2_matriz_abc_ved.png"); plt.close(fig)

# %% [markdown]
# ## 5. Validación de la clasificación (8.2.5)

# %%
# 5.1 Estabilidad entre ventanas (12 meses frente a 29 meses)
ok = tabla[tabla["ABC12"].isin(list("ABC")) & tabla["ABC29"].isin(list("ABC"))]
trans = pd.crosstab(ok["ABC29"], ok["ABC12"])
trans.to_csv(f"{TAB}/o2_val_transicion_29m_a_12m.csv", encoding="utf-8-sig")
res["val_acuerdo_12m_vs_29m_pct"] = round(100 * (ok["ABC12"] == ok["ABC29"]).mean(), 2)
res["val_spearman_valor_12m_vs_29m"] = round(float(spearmanr(ok["VALOR_12M"], ok["VALOR_29M"])[0]), 4)

# 5.2 Estabilidad por año calendario
anios = {}
for a in sorted({p[:4] for p in periodos}):
    anios[a] = inv.clasificar_abc(valor_producto(val, [p for p in periodos if p.startswith(a)]).reindex(todos).fillna(0))
est = pd.DataFrame(anios)
pares = {}
claves = list(anios)
for i in range(len(claves) - 1):
    a, b_ = claves[i], claves[i + 1]
    m = est[a].isin(list("ABC")) & est[b_].isin(list("ABC"))
    pares[f"{a}→{b_}"] = round(100 * (est.loc[m, a] == est.loc[m, b_]).mean(), 2)
res["val_acuerdo_entre_anios_pct"] = pares

# 5.3 Sensibilidad a la regla de costo y a los cortes
act = biblia[biblia["ES_ACTIVO"] & biblia["COSTO_FINAL"].notna()]
cmin = act.groupby("CODIGO_PROPIO")["COSTO_FINAL"].min()
cmax = act.groupby("CODIGO_PROPIO")["COSTO_FINAL"].max()
u12 = agr[agr["PERIODO"].isin(ult)].groupby("CODIGO_PROPIO")["CANTIDAD"].sum().reindex(todos).fillna(0)
sens = {}
for nombre, c in [("costo_minimo_activo", cmin), ("costo_maximo_activo", cmax)]:
    alt = inv.clasificar_abc(u12 * todos.map(c).astype(float).fillna(0))
    m = alt.isin(list("ABC")) & tabla["ABC12"].isin(list("ABC"))
    sens[nombre] = round(100 * (alt[m] == tabla.loc[m, "ABC12"]).mean(), 2)
for nombre, cortes in [("cortes_70_90", (0.70, 0.90)), ("cortes_75_95", (0.75, 0.95))]:
    alt = inv.clasificar_abc(v12, cortes)
    m = alt.isin(list("ABC")) & tabla["ABC12"].isin(list("ABC"))
    sens[nombre] = round(100 * (alt[m] == tabla.loc[m, "ABC12"]).mean(), 2)
res["val_sensibilidad_acuerdo_pct"] = sens

# 5.4 Cruce con los patrones de frecuencia del Objetivo 1 (29 meses)
act29 = agr.groupby(["CODIGO_PROPIO", "PERIODO"])["CANTIDAD"].sum().gt(0).groupby("CODIGO_PROPIO").sum()
tabla["MESES_ACTIVOS"] = act29.reindex(todos).fillna(0).values
tabla["PATRON"] = inv.patron_frecuencia(tabla["MESES_ACTIVOS"] / len(periodos))
cruce = pd.crosstab(tabla["PATRON"], tabla["ABC12"]).reindex(index=inv.PATRONES)
cruce.to_csv(f"{TAB}/o2_val_patron_por_abc.csv", encoding="utf-8-sig")
(pd.crosstab(tabla["PATRON"], tabla["VED"]).reindex(index=inv.PATRONES)
 ).to_csv(f"{TAB}/o2_val_patron_por_ved.csv", encoding="utf-8-sig")

# 5.5 Casos para revisión del químico farmacéutico
tabla["DES_MEDICAMENTO"] = tabla.index.map(desc)
tabla["PCT_VALOR_12M"] = 100 * tabla["VALOR_12M"] / tabla["VALOR_12M"].sum()
top = tabla.sort_values("VALOR_12M", ascending=False).head(20)
top[["DES_MEDICAMENTO", "VALOR_12M", "PCT_VALOR_12M", "ABC12", "VED", "PATRON"]].round(2).to_csv(
    f"{TAB}/o2_val_top20_valor.csv", encoding="utf-8-sig")
rev = tabla[((tabla["ABC12"] == "A") & (tabla["VED"] == "D")) | ((tabla["ABC12"] == "C") & (tabla["VED"] == "V"))]
rev.sort_values("VALOR_12M", ascending=False)[["DES_MEDICAMENTO", "VALOR_12M", "ABC12", "VED", "PATRON"]].round(2).to_csv(
    f"{TAB}/o2_val_casos_para_revision.csv", encoding="utf-8-sig")
res["casos_revision_AD_o_CV"] = int(len(rev))

# %% [markdown]
# ## 6. Exportar la clasificación por producto (insumo del Objetivo 3)

# %%
tabla.index.name = "CODIGO_PROPIO"
tabla.reset_index().to_csv(f"{TAB}/clasificacion_abc_ved_por_producto.csv", index=False, encoding="utf-8-sig")
with open(f"{TAB}/o2_cifras_clave.json", "w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False, indent=2, default=str)
print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
