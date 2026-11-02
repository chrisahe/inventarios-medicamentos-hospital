# %% [markdown]
# # 04 · Objetivo específico 3: niveles de inventario (mínimo, reposición y máximo)
# Entrada: base analítica, Biblia, Novedades, Disponibilidad y la clasificación generada por
# `03_objetivo2_clasificacion_abc_ved.py`. Aplica el protocolo de evaluación de `docs/`.

# %%
import os, sys, json
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
BASE = os.environ.get("BASE_ANALITICA", f"{DATA_DIR}/base_analitica.parquet")
TAB, FIG = os.environ.get("OUT_TAB", "outputs/tablas"), os.environ.get("OUT_FIG", "outputs/figuras")
os.makedirs(TAB, exist_ok=True); os.makedirs(FIG, exist_ok=True)
plt.rcParams.update({"figure.dpi": 130, "axes.spines.top": False, "axes.spines.right": False})

W, L, R = 12, 0.5, 0.5     # ventana (meses), tiempo de reposición y período de revisión (meses)
res = {}

# %% [markdown]
# ## 1. Preparación

# %%
base = inv.leer_base_analitica(BASE)
biblia = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
nov = inv.leer_novedades(f"{DATA_DIR}/Novedades.xlsx")
agr = inv.agregar_base(base)
periodos = sorted(agr["PERIODO"].unique())
desc = base.drop_duplicates("CODIGO_PROPIO").set_index("CODIGO_PROPIO")["DES_MEDICAMENTO"]
clas = pd.read_csv(f"{TAB}/clasificacion_abc_ved_por_producto.csv", dtype={"CODIGO_PROPIO": str}).set_index("CODIGO_PROPIO")

val = inv.asignar_costo(agr, biblia)
con = val.dropna(subset=["COSTO_UNITARIO"])
costo_prod = (con.groupby("CODIGO_PROPIO")["VALOR"].sum() / con.groupby("CODIGO_PROPIO")["CANTIDAD"].sum())

M_df = inv.matriz_demanda(agr, ["CODIGO_PROPIO"], periodos)
M = M_df.to_numpy(dtype=float)
ids = M_df.index
ved_v = clas["VED"].reindex(ids).fillna("E").to_numpy()
costo_v = costo_prod.reindex(ids).to_numpy(dtype=float)
res["series_backtest"] = int(len(ids)); res["periodos"] = len(periodos)

# %% [markdown]
# ## 2. Evaluación retrospectiva de origen móvil (caso base)

# %%
bt0 = inv.backtest(M, costo_v, ved_v, W=W, L=L, R=R)
cap_prop = inv.metricas_backtest(bt0).query("politica == 'PROPUESTA' and VED == 'TODAS'")["inventario_prom_valor_mensual"].iat[0]
k_b2 = inv.calibrar_k_b2(M, costo_v, cap_prop, W, L, R)       # B2 a capital equivalente
res["k_b2_capital_equivalente"] = round(float(k_b2), 3)
bt = inv.backtest(M, costo_v, ved_v, W=W, L=L, R=R, k_b2=k_b2)
met = inv.metricas_backtest(bt)
met.round(2).to_csv(f"{TAB}/o3_backtest_metricas.csv", index=False, encoding="utf-8-sig")
print(met.round(2).to_string(index=False))
res["origenes_evaluados"] = int(bt["origen"].nunique())

# Intervalos de confianza (bootstrap sobre meses) y pruebas pareadas
sm = inv.serie_mensual_metrica(bt, "TODAS")
filas = []
for pol in ["PROPUESTA", "B1_ultimo_mes", "B2_promedio_historico", "A1_promedio_ventana", "B2_capital_equivalente"]:
    f, lo, hi = inv.ic_bootstrap(sm[("fill_valor", pol)].to_numpy() * 100)
    i, ilo, ihi = inv.ic_bootstrap(sm[("valor_inventario_prom", pol)].to_numpy())
    filas.append({"politica": pol, "fill_rate_valor_%": f, "IC95_inf": lo, "IC95_sup": hi,
                  "inventario_prom_mensual": i, "IC95_inv_inf": ilo, "IC95_inv_sup": ihi})
ic = pd.DataFrame(filas); ic.round(2).to_csv(f"{TAB}/o3_backtest_ic95.csv", index=False, encoding="utf-8-sig")

pruebas = []
for base_pol in ["B1_ultimo_mes", "B2_promedio_historico", "A1_promedio_ventana", "B2_capital_equivalente"]:
    pruebas.append({
        "comparacion": f"PROPUESTA vs {base_pol}",
        "meses_propuesta_mayor_fill": int((sm[("fill_valor", "PROPUESTA")] > sm[("fill_valor", base_pol)]).sum()),
        "p_wilcoxon_fill": inv.prueba_wilcoxon(sm[("fill_valor", "PROPUESTA")], sm[("fill_valor", base_pol)]),
        "meses_propuesta_menor_inventario": int((sm[("valor_inventario_prom", "PROPUESTA")] < sm[("valor_inventario_prom", base_pol)]).sum()),
        "p_wilcoxon_inventario": inv.prueba_wilcoxon(sm[("valor_inventario_prom", "PROPUESTA")], sm[("valor_inventario_prom", base_pol)]),
        "meses_evaluados": len(sm)})
pd.DataFrame(pruebas).round(4).to_csv(f"{TAB}/o3_backtest_pruebas_pareadas.csv", index=False, encoding="utf-8-sig")

# Figura: compromiso servicio-capital por política (clase TODAS)
g = met[met["VED"] == "TODAS"].set_index("politica")
fig, ax = plt.subplots(figsize=(6, 4))
for pol, fila in g.iterrows():
    ax.scatter(fila["inventario_prom_valor_mensual"] / 1e6, fila["fill_rate_valor_%"], s=70, label=pol)
ax.set_xlabel("Inventario promedio teórico (millones COP)"); ax.set_ylabel("Fill rate en valor (%)")
ax.legend(fontsize=8); ax.set_title("Nivel de servicio frente a capital inmovilizado")
fig.tight_layout(); fig.savefig(f"{FIG}/o3_servicio_vs_capital.png"); plt.close(fig)

# %% [markdown]
# ## 3. Criterios de éxito del protocolo

# %%
P, B1, B2, BK = (g.loc[k] for k in ["PROPUESTA", "B1_ultimo_mes", "B2_promedio_historico", "B2_capital_equivalente"])
gv = met[met["VED"] == "V"].set_index("politica")
crit = pd.DataFrame([
    ["C1", "Disponibilidad ante demanda en productos vitales ≥ 95 %",
     round(gv.loc["PROPUESTA", "disponibilidad_ante_demanda_%"], 2), "≥ 95", gv.loc["PROPUESTA", "disponibilidad_ante_demanda_%"] >= 95],
    ["C2", "Fill rate en valor de la propuesta > líneas base B1 y B2",
     round(P["fill_rate_valor_%"], 2), f"> {max(B1['fill_rate_valor_%'], B2['fill_rate_valor_%']):.2f}",
     P["fill_rate_valor_%"] > max(B1["fill_rate_valor_%"], B2["fill_rate_valor_%"])],
    ["C3", "A capital equivalente, fill rate de la propuesta > B2 escalada",
     round(P["fill_rate_valor_%"], 2), f"> {BK['fill_rate_valor_%']:.2f}", P["fill_rate_valor_%"] > BK["fill_rate_valor_%"]],
], columns=["Criterio", "Descripción", "Resultado", "Umbral", "Cumple"])
res["capital_propuesta_vs_B2"] = round(float(P["inventario_prom_valor_mensual"] / B2["inventario_prom_valor_mensual"]), 3)
print(crit.to_string(index=False))

# %% [markdown]
# ## 4. Sensibilidad (ventana, tiempo de reposición, nivel de servicio)

# %%
sens = []
for etiqueta, kw in ([(f"ventana={w}", dict(W=w)) for w in (6, 12, 24)] +
                     [(f"L={l}", dict(L=l, R=1 - l)) for l in (0.25, 0.5, 0.75)] +
                     [(f"Δ nivel servicio={d:+.2f}", dict(delta_sl=d)) for d in (-0.05, 0.0, 0.03)]):
    args = dict(W=W, L=L, R=R); args.update(kw)
    m2 = inv.metricas_backtest(inv.backtest(M, costo_v, ved_v, **args))
    m2 = m2[(m2["politica"] == "PROPUESTA") & m2["VED"].isin(["TODAS", "V"])]
    for _, f in m2.iterrows():
        sens.append({"escenario": etiqueta, "VED": f["VED"], "fill_rate_valor_%": f["fill_rate_valor_%"],
                     "disponibilidad_%": f["disponibilidad_ante_demanda_%"],
                     "inventario_prom_millones": f["inventario_prom_valor_mensual"] / 1e6})
pd.DataFrame(sens).round(2).to_csv(f"{TAB}/o3_sensibilidad.csv", index=False, encoding="utf-8-sig")

# %% [markdown]
# ## 5. Niveles finales por producto y servicio farmacéutico (corte: último mes)
# Se aplican las reglas de negocio de Novedades, que describen el estado vigente y por eso
# NO se usan en la evaluación retrospectiva (hacerlo sería fuga de información).

# %%
Ms = inv.matriz_demanda(agr, ["CODIGO_PROPIO", "SERVICIO_FARMACEUTICO"], periodos)
win = Ms.to_numpy(dtype=float)[:, -W:]
cod_s = Ms.index.get_level_values(0)
abc_s = clas["ABC12"].reindex(cod_s).fillna("SIN_VALOR").to_numpy()
ved_s = clas["VED"].reindex(cod_s).fillna("E").to_numpy()
pol = inv.politica_propuesta(win, abc_s, ved_s, L, R)

out = pd.DataFrame({
    "CODIGO_PROPIO": cod_s, "SERVICIO_FARMACEUTICO": Ms.index.get_level_values(1),
    "DES_MEDICAMENTO": cod_s.map(desc), "ABC": abc_s, "VED": ved_s, "PATRON": pol["patron"],
    "NIVEL_SERVICIO": pol["nivel_servicio"], "MEDIA_MENSUAL": pol["media_mensual"].round(2),
    "DESV_MENSUAL": pol["desv_mensual"].round(2), "MINIMO": pol["minimo"],
    "PUNTO_REPOSICION": pol["punto_reposicion"], "MAXIMO": pol["maximo"]})
out = out.merge(nov, on="CODIGO_PROPIO", how="left")
for c in [c for c in out.columns if c.startswith("NOV_")]:
    out[c] = out[c].fillna(False).astype(bool)
out["REGLA_NEGOCIO"] = ""
def regla(mask, texto, anular=False, minimo=None):
    global out
    out.loc[mask, "REGLA_NEGOCIO"] = (out.loc[mask, "REGLA_NEGOCIO"] + texto + "; ")
    if anular:
        out.loc[mask, ["MINIMO", "PUNTO_REPOSICION", "MAXIMO"]] = 0
    if minimo:
        out.loc[mask, "MINIMO"] = np.maximum(out.loc[mask, "MINIMO"], minimo)
        out.loc[mask, "PUNTO_REPOSICION"] = np.maximum(out.loc[mask, "PUNTO_REPOSICION"], minimo)
        out.loc[mask, "MAXIMO"] = np.maximum(out.loc[mask, "MAXIMO"], out.loc[mask, "PUNTO_REPOSICION"] + 1)
regla(out["NOV_NO_SE_UTILIZA"], "producto descontinuado (sin nivel)", anular=True)
regla(out["NOV_SOLO_NECESIDAD"] & ~out["NOV_NO_SE_UTILIZA"], "se solicita solo a necesidad (sin nivel)", anular=True)
regla(out["NOV_SIEMPRE_EXISTENCIAS"], "existencia permanente obligatoria (mínimo 1)", minimo=1)
regla(out["NOV_INVIMA_VENCIDO"], "REGISTRO INVIMA VENCIDO: verificar antes de comprar")
regla(out["NOV_DESABASTECIDO"], "desabastecimiento reportado: evaluar L mayor")
out.loc[out["PATRON"].eq(inv.PATRONES[3]) & out["VED"].ne("V"), "REGLA_NEGOCIO"] += "esporádico no vital: bajo demanda; "
out.to_csv(f"{TAB}/niveles_inventario_por_producto_servicio.csv", index=False, encoding="utf-8-sig")

resumen = (out.groupby(["PATRON"]).agg(series=("MAXIMO", "size"), con_nivel=("MAXIMO", lambda s: int((s > 0).sum())),
                                       maximo_total=("MAXIMO", "sum")))
resumen.to_csv(f"{TAB}/o3_resumen_por_patron.csv", encoding="utf-8-sig")
res["series_finales"] = int(len(out)); res["series_con_nivel"] = int((out["MAXIMO"] > 0).sum())
print(resumen)

# %% [markdown]
# ## 6. Comparación con el inventario actual (indicativa)
# `Disponibilidad.xlsx` no trae fecha de corte: la comparación es referencial.

# %%
ruta_disp = f"{DATA_DIR}/Disponibilidad.xlsx"
if os.path.exists(ruta_disp):
    d = pd.read_excel(ruta_disp)
    d.columns = [f"c{i}" for i in range(d.shape[1])]
    d["CODIGO_PROPIO"] = inv._codigo(d["c2"].astype(str)); d["BODEGA"] = d["c18"]
    ex = d["c22"].str.extract(r"hay:\s*(-?\d+)")[0].astype(float)
    stock = (d.assign(EX=ex).groupby(["CODIGO_PROPIO", "BODEGA"])["EX"].first().groupby("CODIGO_PROPIO").sum())
    nivel = out.groupby("CODIGO_PROPIO")[["PUNTO_REPOSICION", "MAXIMO"]].sum()
    cmp_ = nivel.join(stock.rename("STOCK_ACTUAL"), how="inner").join(costo_prod.rename("COSTO"))
    cmp_["EXCESO_UNID"] = (cmp_["STOCK_ACTUAL"] - cmp_["MAXIMO"]).clip(lower=0)
    cmp_["FALTANTE_VS_REPOSICION"] = (cmp_["PUNTO_REPOSICION"] - cmp_["STOCK_ACTUAL"]).clip(lower=0)
    cmp_["VALOR_EXCESO"] = cmp_["EXCESO_UNID"] * cmp_["COSTO"].fillna(0)
    ind = {"productos_comparados": len(cmp_),
           "productos_con_stock_sobre_maximo": int((cmp_["EXCESO_UNID"] > 0).sum()),
           "pct_productos_sobre_maximo": round(100 * (cmp_["EXCESO_UNID"] > 0).mean(), 2),
           "valor_exceso_cop": round(cmp_["VALOR_EXCESO"].sum()),
           "productos_bajo_punto_reposicion": int((cmp_["FALTANTE_VS_REPOSICION"] > 0).sum()),
           "pct_productos_bajo_punto_reposicion": round(100 * (cmp_["FALTANTE_VS_REPOSICION"] > 0).mean(), 2)}
    pd.Series(ind).to_csv(f"{TAB}/o3_comparacion_inventario_actual.csv", header=["valor"], encoding="utf-8-sig")
    res["comparacion_inventario_actual"] = ind
    print(ind)

with open(f"{TAB}/o3_cifras_clave.json", "w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False, indent=2, default=str)
crit.to_csv(f"{TAB}/o3_criterios_exito.csv", index=False, encoding="utf-8-sig")
