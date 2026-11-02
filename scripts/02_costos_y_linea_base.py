# %% [markdown]
# # 02 · Análisis descriptivo del costo unitario y línea base de inventario
# Responde a dos observaciones del ACA 1:
# * "incluir un análisis descriptivo del costo unitario de los productos" (insumo del ABC);
# * "definir la línea base (método empírico actual del hospital)" del protocolo de evaluación.
# Solo usa las fuentes maestras (Biblia, Disponibilidad, Novedades); no requiere el consumo.

# %%
import os, sys, re
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, "src")
try:
    sys.stdout.reconfigure(encoding="utf-8")   # consola de Windows
except Exception:
    pass
import inventario as inv

DATA_DIR = os.environ.get("DATA_DIR", "data")
TAB, FIG = "outputs/tablas", "outputs/figuras"
os.makedirs(TAB, exist_ok=True); os.makedirs(FIG, exist_ok=True)
plt.rcParams.update({"figure.dpi": 130, "axes.spines.top": False, "axes.spines.right": False})

b = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
nov = inv.leer_novedades(f"{DATA_DIR}/Novedades.xlsx")

# %% [markdown]
# ## 1. Costo unitario por producto con la regla de asignación

# %%
c_bod, c_cod, c_ina = inv.tablas_costo(b)
cod = b.drop_duplicates("CODIGO_PROPIO")[["CODIGO_PROPIO", "CONCEPTO", "VEN"]].copy()
cod["COSTO"] = cod["CODIGO_PROPIO"].map(c_cod.set_index("CODIGO_PROPIO")["COSTO_FINAL"])
cod["FUENTE"] = np.where(cod["COSTO"].notna(), "codigo_activo", "")
m = cod["COSTO"].isna()
cod.loc[m, "COSTO"] = cod.loc[m, "CODIGO_PROPIO"].map(c_ina.set_index("CODIGO_PROPIO")["COSTO_FINAL"])
cod.loc[m & cod["COSTO"].notna(), "FUENTE"] = "codigo_inactivo"
cod.loc[cod["COSTO"].isna(), "FUENTE"] = "sin_costo"

cobertura = cod["FUENTE"].value_counts().rename_axis("Fuente del costo").reset_index(name="Productos")
cobertura["%"] = (100 * cobertura["Productos"] / len(cod)).round(2)
cobertura.to_csv(f"{TAB}/costo_cobertura.csv", index=False, encoding="utf-8-sig")

# Variación del costo de un mismo producto entre bodegas (registros activos)
act = b[b["ES_ACTIVO"] & b["COSTO_FINAL"].notna()]
g = act.groupby("CODIGO_PROPIO")["COSTO_FINAL"].agg(["min", "max", "nunique"])
multi = g[g["nunique"] > 1].assign(razon=lambda x: x["max"] / x["min"])
var_costo = {
    "productos_con_costo_activo": len(g),
    "productos_con_mas_de_un_costo_activo": len(multi),
    "mediana_razon_max_min": round(multi["razon"].median(), 2),
    "p90_razon_max_min": round(multi["razon"].quantile(.9), 2),
}
pd.Series(var_costo).to_csv(f"{TAB}/costo_variacion_entre_bodegas.csv", header=["valor"], encoding="utf-8-sig")

# %% [markdown]
# ## 2. Estadísticos descriptivos del costo unitario

# %%
def desc(s):
    s = s.dropna()
    return pd.Series({"n": len(s), "media": s.mean(), "desv_est": s.std(), "mínimo": s.min(),
                      "p25": s.quantile(.25), "mediana": s.median(), "p75": s.quantile(.75),
                      "p90": s.quantile(.90), "p99": s.quantile(.99), "máximo": s.max(),
                      "CV_%": 100 * s.std() / s.mean()})

d_all = desc(cod["COSTO"]).to_frame("Todos").T
d_con = cod.groupby("CONCEPTO")["COSTO"].apply(desc).unstack()
d_ven = cod.groupby("VEN")["COSTO"].apply(desc).unstack()
descr = pd.concat([d_all, d_con, d_ven])
descr.round(2).to_csv(f"{TAB}/costo_descriptivo.csv", encoding="utf-8-sig")
print(descr.round(0))

# Valores extremos (regla de Tukey sobre log10 del costo) y los 15 más costosos
lc = np.log10(cod["COSTO"].dropna())
q1, q3 = lc.quantile([.25, .75]); lim = q3 + 1.5 * (q3 - q1)
n_ext = int((lc > lim).sum())
top = cod.dropna(subset=["COSTO"]).nlargest(15, "COSTO").merge(
    b.drop_duplicates("CODIGO_PROPIO")[["CODIGO_PROPIO", "DES_MEDICAMENTO"]], on="CODIGO_PROPIO")
top[["CODIGO_PROPIO", "DES_MEDICAMENTO", "CONCEPTO", "VEN", "COSTO"]].to_csv(
    f"{TAB}/costo_top15.csv", index=False, encoding="utf-8-sig")
pd.Series({"limite_superior_tukey_log10": lim, "costo_limite": 10 ** lim, "productos_extremos": n_ext}).to_csv(
    f"{TAB}/costo_extremos.csv", header=["valor"], encoding="utf-8-sig")

# Figuras
fig, ax = plt.subplots(1, 2, figsize=(10, 3.8))
ax[0].hist(lc, bins=40, color="#3b6ea5"); ax[0].set_xlabel("log10(costo unitario, COP)")
ax[0].set_ylabel("Número de productos"); ax[0].set_title("Distribución del costo unitario")
datos = [np.log10(g_["COSTO"].dropna()) for _, g_ in cod.groupby("CONCEPTO")]
ax[1].boxplot(datos, tick_labels=[k for k, _ in cod.groupby("CONCEPTO")], showfliers=False)
ax[1].set_ylabel("log10(costo unitario, COP)"); ax[1].set_title("Costo por tipo de producto")
ax[1].tick_params(axis="x", labelsize=8)
fig.tight_layout(); fig.savefig(f"{FIG}/costo_distribucion.png"); plt.close(fig)

# %% [markdown]
# ## 3. Criticidad institucional (VEN) por tipo de producto

# %%
ven = pd.crosstab(cod["CONCEPTO"], cod["VEN"])
ven_pct = (ven.div(ven.sum(axis=1), axis=0) * 100).round(2)
ven.to_csv(f"{TAB}/ven_por_concepto_n.csv", encoding="utf-8-sig")
ven_pct.to_csv(f"{TAB}/ven_por_concepto_pct.csv", encoding="utf-8-sig")
tot_ven = cod["VEN"].value_counts().rename_axis("VEN").reset_index(name="Productos")
tot_ven["%"] = (100 * tot_ven["Productos"] / len(cod)).round(2)
tot_ven.to_csv(f"{TAB}/ven_total.csv", index=False, encoding="utf-8-sig")
print(tot_ven)

# %% [markdown]
# ## 4. Línea base: inventario actual y niveles mínimo/máximo parametrizados
# El sistema institucional deja en `ALERTA_STOCK` el texto "Del medicamento hay: X
# para un stock mínimo de Y y máximo de Z". Se extraen los tres números.

# %%
d = pd.read_excel(f"{DATA_DIR}/Disponibilidad.xlsx")
d.columns = ["CANTIDAD", "CODIGO", "CODIGO_GENERICO", "LOTE", "FECHA_VENCIMIENTO", "NOMBRE_COMERCIAL",
             "PRINCIPIO_ACTIVO", "PRESENTACION", "CLASIF_RIESGO", "ALMACENAMIENTO", "LABORATORIO",
             "REG_INVIMA", "FORMA_FARM", "FOTOSENSIBILIDAD", "ESTABILIDAD", "VENCIDO", "STOCK_EN_MINIMOS",
             "STOCK_EN_MAXIMOS", "BODEGA", "GRUPO", "SUBGRUPO", "CUM", "ALERTA_STOCK"]
d["CODIGO_PROPIO"] = inv._codigo(d["CODIGO_GENERICO"].astype(str))
ex = d["ALERTA_STOCK"].str.extract(r"hay:\s*(-?\d+).*?m.nimo de (\d+) y m.ximo de (\d+)").astype(float)
d[["EXIST_PRODUCTO", "MIN_SISTEMA", "MAX_SISTEMA"]] = ex.values

# Una fila por producto-bodega (el alerta repite el total del producto en cada lote)
pb = (d.groupby(["CODIGO_PROPIO", "BODEGA"], as_index=False)
        .agg(EXIST=("EXIST_PRODUCTO", "first"), MIN=("MIN_SISTEMA", "first"), MAX=("MAX_SISTEMA", "first"),
             UNIDADES_LOTES=("CANTIDAD", "sum"), GRUPO=("GRUPO", "first")))
pb["TIENE_MINMAX"] = (pb["MAX"] > 0)
costo_cod = cod.set_index("CODIGO_PROPIO")["COSTO"]
d["COSTO"] = d["CODIGO_PROPIO"].map(costo_cod)
d["VALOR"] = d["CANTIDAD"] * d["COSTO"]

linea_base = {
    "lotes_registrados": len(d),
    "productos_distintos": d["CODIGO_PROPIO"].nunique(),
    "combinaciones_producto_bodega": len(pb),
    "producto_bodega_con_min_max_definido": int(pb["TIENE_MINMAX"].sum()),
    "pct_producto_bodega_con_min_max_definido": round(100 * pb["TIENE_MINMAX"].mean(), 2),
    "pct_producto_bodega_sin_min_max (0/0)": round(100 * (1 - pb["TIENE_MINMAX"].mean()), 2),
    "lotes_con_costo_asignado_%": round(100 * d["COSTO"].notna().mean(), 2),
    "valor_inventario_cop (lotes con costo)": round(d["VALOR"].sum()),
    "lotes_vencidos": int((d["VENCIDO"] == "Si").sum()),
    "valor_lotes_vencidos_cop": round(d.loc[d["VENCIDO"] == "Si", "VALOR"].sum()),
    "pct_valor_vencido_sobre_inventario": round(100 * d.loc[d["VENCIDO"] == "Si", "VALOR"].sum() / d["VALOR"].sum(), 3),
    "lotes_con_existencia_negativa": int((d["EXIST_PRODUCTO"] < 0).sum()),
    "pct_filas_indicador_Stock_en_Maximos=Si": round(100 * (d["STOCK_EN_MAXIMOS"] == "Si").mean(), 2),
}
pd.Series(linea_base).to_csv(f"{TAB}/linea_base_inventario_actual.csv", header=["valor"], encoding="utf-8-sig")
pb_def = pb[pb["TIENE_MINMAX"]]
(pb.groupby("GRUPO").agg(producto_bodega=("TIENE_MINMAX", "size"), con_min_max=("TIENE_MINMAX", "sum"))
   .assign(pct=lambda x: (100 * x["con_min_max"] / x["producto_bodega"]).round(2))
   .to_csv(f"{TAB}/linea_base_minmax_por_grupo.csv", encoding="utf-8-sig"))
pb_def.to_csv(f"{TAB}/linea_base_minmax_definidos.csv", index=False, encoding="utf-8-sig")
print(pd.Series(linea_base).to_string())

# %% [markdown]
# ## 5. Cobertura de las novedades respecto de la Biblia

# %%
cob_nov = {
    "productos_con_novedad": nov["CODIGO_PROPIO"].nunique(),
    "pct_en_biblia": round(100 * nov["CODIGO_PROPIO"].isin(b["CODIGO_PROPIO"]).mean(), 2),
    **{c: int(nov[c].sum()) for c in nov.columns if c.startswith("NOV_")},
}
pd.Series(cob_nov).to_csv(f"{TAB}/novedades_cobertura.csv", header=["valor"], encoding="utf-8-sig")
print(pd.Series(cob_nov).to_string())
