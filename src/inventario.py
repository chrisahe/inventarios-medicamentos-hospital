"""Funciones del proyecto: clasificación ABC-VED y niveles de inventario.

Modelo analítico para la optimización de inventarios de medicamentos y
dispositivos médicos (Objetivos específicos 2 y 3).

Convenciones
------------
* Unidad de tiempo: mes (la fuente de consumo solo trae año y mes).
* Una "serie" es un producto (CODIGO_PROPIO) o un producto por servicio
  farmacéutico. La matriz de demanda tiene una fila por serie y una columna
  por período mensual (los meses sin registro valen 0).
* Todas las reglas que dependen de supuestos (niveles de servicio, tiempo de
  reposición, cortes ABC) están en las constantes del inicio para que puedan
  auditarse y variarse en el análisis de sensibilidad.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy.stats import norm, wilcoxon

# --------------------------------------------------------------------------
# Parámetros de diseño (supuestos documentados en docs/bitacora_decisiones.md)
# --------------------------------------------------------------------------
CORTES_ABC = (0.80, 0.95)          # A: hasta 80 % del valor, B: hasta 95 %, C: resto
MAPA_VEN_A_VED = {"V": "V", "E": "E", "N": "D"}   # N (no esencial) = D (deseable)

# Nivel de servicio objetivo por celda de la matriz ABC-VED.
# Criterio: más servicio donde la criticidad es mayor; menos donde el producto
# es costoso y no crítico (protege capital); más holgura en C porque
# mantener inventario cuesta poco.
NIVEL_SERVICIO = {
    ("A", "V"): 0.98, ("B", "V"): 0.98, ("C", "V"): 0.99,
    ("A", "E"): 0.93, ("B", "E"): 0.95, ("C", "E"): 0.97,
    ("A", "D"): 0.85, ("B", "D"): 0.90, ("C", "D"): 0.92,
}
NIVEL_SERVICIO_DEFECTO = 0.90

# Umbrales de los patrones de frecuencia del ACA 1 (24-29, 12-23, 4-11, 1-3 de
# 29 meses) expresados como fracción de meses con consumo, para poder aplicarlos
# a ventanas distintas de 29 meses.
UMBRALES_PATRON = (24 / 29, 12 / 29, 4 / 29)
PATRONES = ["Consumo sostenido", "Consumo frecuente",
            "Consumo intermitente", "Consumo esporádico"]

COLUMNAS_BIBLIA = ["CUM", "ATC", "BODEGA", "CODIGO_PROPIO", "DES_MEDICAMENTO",
                   "CONCEPTO", "ES_ACTIVO", "MOTIVO_INACTIVACION",
                   "COSTO_FINAL", "REGISTRO_SANITARIO", "VEN"]
COLUMNAS_BASE = ["CODIGO_PROPIO", "DES_MEDICAMENTO", "CANTIDAD", "BODEGA",
                 "DES_UF", "ANIO", "MES", "SERVICIO_FARMACEUTICO"]


# --------------------------------------------------------------------------
# Lectura y limpieza de fuentes
# --------------------------------------------------------------------------
def _codigo(serie: pd.Series) -> pd.Series:
    """Código como texto sin espacios ni '.0' (evita fallos de unión)."""
    return (serie.astype("string").str.strip()
            .str.replace(r"\.0$", "", regex=True))


def leer_biblia(ruta: str) -> pd.DataFrame:
    """Lee la base maestra de productos (Biblia) y la deja tipada y sin duplicados."""
    df = pd.read_excel(ruta, dtype=str)
    if df.shape[1] != len(COLUMNAS_BIBLIA):
        raise ValueError(f"Se esperaban {len(COLUMNAS_BIBLIA)} columnas y hay {df.shape[1]}")
    df.columns = COLUMNAS_BIBLIA          # renombre por posición (hay tildes dañadas)
    df = df.drop_duplicates()
    df["CODIGO_PROPIO"] = _codigo(df["CODIGO_PROPIO"])
    df["BODEGA"] = _codigo(df["BODEGA"])
    df = df.dropna(subset=["CODIGO_PROPIO"])
    df["COSTO_FINAL"] = pd.to_numeric(df["COSTO_FINAL"], errors="coerce")
    df["ES_ACTIVO"] = df["ES_ACTIVO"].str.strip().str.lower().eq("true")
    df["VEN"] = df["VEN"].str.strip().str.upper()
    return df.reset_index(drop=True)


def leer_novedades(ruta: str) -> pd.DataFrame:
    """Novedades vigentes por producto, una columna booleana por tipo de novedad.

    Las novedades se usan como atributo vigente a la fecha de corte del archivo y
    no como explicación temporal del consumo (ver ACA 1, Figura 18).
    """
    df = pd.read_excel(ruta, dtype=str)
    df.columns = ["CODIGO_PROPIO", "DES_MEDICAMENTO", "NOVEDAD", "FECHA_REPORTE"]
    df["CODIGO_PROPIO"] = _codigo(df["CODIGO_PROPIO"])
    nov = df["NOVEDAD"].str.strip().str.lower()
    flags = pd.DataFrame({
        "CODIGO_PROPIO": df["CODIGO_PROPIO"],
        "NOV_INVIMA_VENCIDO": nov.str.contains("invima", na=False),
        "NOV_SOLO_NECESIDAD": nov.str.contains("solo se solicita", na=False),
        "NOV_NO_SE_UTILIZA": nov.str.contains("ya no se utiliza", na=False),
        "NOV_DESABASTECIDO": nov.str.contains("desabastecido", na=False),
        "NOV_SIEMPRE_EXISTENCIAS": nov.str.contains("siempre deben", na=False),
    })
    return flags.groupby("CODIGO_PROPIO", as_index=False).max()


def leer_base_analitica(ruta: str) -> pd.DataFrame:
    """Base analítica exportada del cuaderno del Objetivo 1 (parquet o csv)."""
    df = pd.read_parquet(ruta) if ruta.endswith(".parquet") else pd.read_csv(ruta, dtype={"CODIGO_PROPIO": str})
    df = df.rename(columns={"AÑO": "ANIO", "Año": "ANIO", "Mes": "MES"})
    faltan = [c for c in COLUMNAS_BASE if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en la base analítica: {faltan}")
    df["CODIGO_PROPIO"] = _codigo(df["CODIGO_PROPIO"])
    df["BODEGA"] = _codigo(df["BODEGA"])
    df["PERIODO"] = (df["ANIO"].astype(int).astype(str) + "-"
                     + df["MES"].astype(int).astype(str).str.zfill(2))
    return df


def agregar_base(base: pd.DataFrame) -> pd.DataFrame:
    """Reduce los registros a producto x bodega x servicio x período (suma de CANTIDAD)."""
    return (base.groupby(["CODIGO_PROPIO", "BODEGA", "SERVICIO_FARMACEUTICO", "PERIODO"],
                         as_index=False)["CANTIDAD"].sum())


# --------------------------------------------------------------------------
# Costos (8.2.1)
# --------------------------------------------------------------------------
def tablas_costo(biblia: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Costo unitario por (código, bodega), por código con registros activos y por
    código con registros inactivos (último recurso). Se usa la mediana."""
    con_costo = biblia.dropna(subset=["COSTO_FINAL"])
    act = con_costo[con_costo["ES_ACTIVO"]]
    ina = con_costo[~con_costo["ES_ACTIVO"]]
    c_bod = act.groupby(["CODIGO_PROPIO", "BODEGA"], as_index=False)["COSTO_FINAL"].median()
    c_cod = act.groupby("CODIGO_PROPIO", as_index=False)["COSTO_FINAL"].median()
    c_ina = ina.groupby("CODIGO_PROPIO", as_index=False)["COSTO_FINAL"].median()
    return c_bod, c_cod, c_ina


def asignar_costo(agr: pd.DataFrame, biblia: pd.DataFrame) -> pd.DataFrame:
    """Asigna costo unitario a cada fila de consumo con la regla de cuatro niveles:
    1) costo activo del producto en la misma bodega, 2) mediana de costos activos
    del producto, 3) mediana de costos de registros inactivos, 4) sin costo."""
    c_bod, c_cod, c_ina = tablas_costo(biblia)
    out = agr.merge(c_bod.rename(columns={"COSTO_FINAL": "c1"}), on=["CODIGO_PROPIO", "BODEGA"], how="left")
    out = out.merge(c_cod.rename(columns={"COSTO_FINAL": "c2"}), on="CODIGO_PROPIO", how="left")
    out = out.merge(c_ina.rename(columns={"COSTO_FINAL": "c3"}), on="CODIGO_PROPIO", how="left")
    out["COSTO_UNITARIO"] = out["c1"].fillna(out["c2"]).fillna(out["c3"])
    out["FUENTE_COSTO"] = np.select(
        [out["c1"].notna(), out["c2"].notna(), out["c3"].notna()],
        ["1_bodega_activo", "2_codigo_activo", "3_codigo_inactivo"], default="4_sin_costo")
    out["VALOR"] = out["CANTIDAD"] * out["COSTO_UNITARIO"]
    return out.drop(columns=["c1", "c2", "c3"])


# --------------------------------------------------------------------------
# Criticidad VED (8.2.3)
# --------------------------------------------------------------------------
def tabla_ved(codigos: pd.Series, biblia: pd.DataFrame, novedades: pd.DataFrame | None = None) -> pd.DataFrame:
    """VED por producto a partir de la clasificación VEN institucional de la Biblia.

    N (no esencial) se homologa a D (deseable). Productos sin clasificación se
    asignan a E (criterio conservador) y quedan marcados para revisión del
    químico farmacéutico. La novedad 'Siempre deben haber existencias' eleva a V.
    """
    ven = (biblia.dropna(subset=["VEN"]).groupby("CODIGO_PROPIO")["VEN"]
           .agg(lambda s: s.mode().iat[0]))
    out = pd.DataFrame({"CODIGO_PROPIO": pd.Series(codigos.unique(), dtype="string")})
    out["VEN"] = out["CODIGO_PROPIO"].map(ven)
    out["VED"] = out["VEN"].map(MAPA_VEN_A_VED)
    out["FUENTE_VED"] = np.where(out["VED"].notna(), "biblia_VEN", "defecto_E")
    out["VED"] = out["VED"].fillna("E")
    if novedades is not None:
        siempre = set(novedades.loc[novedades["NOV_SIEMPRE_EXISTENCIAS"], "CODIGO_PROPIO"])
        m = out["CODIGO_PROPIO"].isin(siempre)
        out.loc[m, ["VED", "FUENTE_VED"]] = ["V", "novedad_siempre_existencias"]
    return out


# --------------------------------------------------------------------------
# Clasificación ABC (8.2.2)
# --------------------------------------------------------------------------
def clasificar_abc(valor: pd.Series, cortes=CORTES_ABC) -> pd.Series:
    """ABC por participación acumulada del valor de consumo.

    Un producto es A si el acumulado *previo* a él es menor que el primer corte
    (así el producto que cruza el 80 % queda en A). Valor <= 0 o nulo -> 'SIN_VALOR'.
    """
    v = valor.fillna(0.0)
    res = pd.Series("SIN_VALOR", index=v.index, dtype="object")
    pos = v[v > 0].sort_values(ascending=False)
    if pos.empty:
        return res
    previo = (pos.cumsum() - pos) / pos.sum()
    res.loc[pos.index] = np.select([previo < cortes[0], previo < cortes[1]], ["A", "B"], default="C")
    return res


def patron_frecuencia(frac_meses_activos: np.ndarray) -> np.ndarray:
    """Patrón descriptivo del ACA 1 a partir de la fracción de meses con consumo."""
    f = np.asarray(frac_meses_activos, dtype=float) + 1e-9
    return np.select([f >= UMBRALES_PATRON[0], f >= UMBRALES_PATRON[1], f >= UMBRALES_PATRON[2]],
                     [PATRONES[0], PATRONES[1], PATRONES[2]], default=PATRONES[3])


def matriz_demanda(agr: pd.DataFrame, claves: list[str], periodos: list[str] | None = None) -> pd.DataFrame:
    """Matriz serie x período con ceros en los meses sin consumo."""
    if periodos is None:
        periodos = sorted(agr["PERIODO"].unique())
    m = agr.pivot_table(index=claves, columns="PERIODO", values="CANTIDAD", aggfunc="sum", fill_value=0)
    return m.reindex(columns=periodos, fill_value=0)


def tabla_abc_ved(abc: pd.Series, ved: pd.Series, valor: pd.Series) -> dict[str, pd.DataFrame]:
    """Matriz 3x3 con número de productos, % de productos y % del valor."""
    d = pd.DataFrame({"ABC": abc, "VED": ved, "VALOR": valor.fillna(0)})
    d = d[d["ABC"].isin(list("ABC"))]
    n = pd.crosstab(d["ABC"], d["VED"]).reindex(index=list("ABC"), columns=list("VED"), fill_value=0)
    val = d.pivot_table(index="ABC", columns="VED", values="VALOR", aggfunc="sum", fill_value=0)
    val = val.reindex(index=list("ABC"), columns=list("VED"), fill_value=0)
    return {"productos": n, "pct_productos": (n / n.values.sum() * 100).round(2),
            "pct_valor": (val / val.values.sum() * 100).round(2)}


# --------------------------------------------------------------------------
# Políticas de inventario (Objetivo 3)
# --------------------------------------------------------------------------
def nivel_servicio_vector(abc: np.ndarray, ved: np.ndarray, tabla=NIVEL_SERVICIO,
                          delta: float = 0.0) -> np.ndarray:
    """Nivel de servicio objetivo por serie. `delta` desplaza todos los niveles
    (análisis de sensibilidad), acotado a [0.5, 0.999]."""
    sl = np.array([tabla.get((a, v), NIVEL_SERVICIO_DEFECTO) for a, v in zip(abc, ved)])
    return np.clip(sl + delta, 0.5, 0.999)


def politica_propuesta(X: np.ndarray, abc: np.ndarray, ved: np.ndarray,
                       L: float = 0.5, R: float = 0.5, delta_sl: float = 0.0) -> dict[str, np.ndarray]:
    """Mínimo (stock de seguridad), punto de reposición y máximo por serie.

    X : demanda mensual histórica (n series x W meses) de la ventana de estimación.
    L : tiempo de reposición en meses; R: período de revisión en meses; H = L + R.

    Regla por patrón de frecuencia (ver docs/protocolo_evaluacion.md):
      sostenido / frecuente -> modelo normal: máx = mu*H + z*sd*sqrt(H),
                               reposición = mu*L + z*sd*sqrt(L)
      intermitente          -> cuantil empírico del consumo mensual al nivel de
                               servicio, escalado por H (y por L para el mínimo)
      esporádico            -> V: máximo mensual de la ventana; E y D: bajo demanda.
    """
    n, W = X.shape
    H = L + R
    mu = X.mean(axis=1)
    sd = X.std(axis=1, ddof=1) if W > 1 else np.zeros(n)
    patron = patron_frecuencia((X > 0).sum(axis=1) / W)
    sl = nivel_servicio_vector(abc, ved, delta=delta_sl)
    z = norm.ppf(sl)

    # Normal
    max_n = mu * H + z * sd * np.sqrt(H)
    min_n = mu * L + z * sd * np.sqrt(L)
    # Cuantil empírico del consumo mensual (inversa de la distribución empírica)
    Xs = np.sort(X, axis=1)
    idx = np.clip(np.ceil(sl * W).astype(int) - 1, 0, W - 1)
    q = Xs[np.arange(n), idx]
    max_e, min_e = q * H, q * L
    # Esporádico: solo vitales mantienen existencia
    es_v = np.asarray(ved) == "V"
    max_s = np.where(es_v, Xs[:, -1], 0.0)
    min_s = np.where(es_v, Xs[:, -1] * L / H, 0.0)

    normal = np.isin(patron, PATRONES[:2])
    inter = patron == PATRONES[2]
    mx = np.where(normal, max_n, np.where(inter, max_e, max_s))
    mn = np.where(normal, min_n, np.where(inter, min_e, min_s))
    rop = np.ceil(np.maximum(mn, 0))                     # punto de reposición
    mx = np.ceil(np.maximum(mx, rop))                    # máximo (nivel objetivo)
    ss = np.ceil(np.maximum(rop - mu * L, 0))            # stock de seguridad = mínimo
    return {"minimo": ss, "punto_reposicion": rop, "maximo": mx, "patron": patron,
            "nivel_servicio": sl, "media_mensual": mu, "desv_mensual": sd}


def politica_ultimo_mes(Xhist: np.ndarray, L=0.5, R=0.5) -> dict[str, np.ndarray]:
    """Línea base B1: reponer lo consumido el mes anterior (práctica reactiva)."""
    u = Xhist[:, -1]
    return {"punto_reposicion": np.ceil(u * L), "maximo": np.ceil(u * (L + R))}


def politica_promedio_historico(Xhist: np.ndarray, L=0.5, R=0.5) -> dict[str, np.ndarray]:
    """Línea base B2: promedio de todo el histórico disponible, sin stock de seguridad."""
    mu = Xhist.mean(axis=1)
    return {"punto_reposicion": np.ceil(mu * L), "maximo": np.ceil(mu * (L + R))}


def politica_promedio_ventana(X: np.ndarray, L=0.5, R=0.5) -> dict[str, np.ndarray]:
    """Ablación A1: promedio de la ventana reciente sin stock de seguridad ni
    reglas por clase (aísla el aporte del stock de seguridad y de ABC-VED)."""
    mu = X.mean(axis=1)
    return {"punto_reposicion": np.ceil(mu * L), "maximo": np.ceil(mu * (L + R))}


# --------------------------------------------------------------------------
# Evaluación retrospectiva (protocolo de evaluación)
# --------------------------------------------------------------------------
def calibrar_k_b2(M: np.ndarray, costo: np.ndarray, inventario_objetivo: float, W: int = 12,
                  L: float = 0.5, R: float = 0.5) -> float:
    """Factor k tal que la línea base B2 (promedio histórico) multiplicada por k
    inmoviliza, en promedio mensual, el mismo capital que la propuesta. Permite
    comparar el servicio de ambas a capital equivalente (bisección)."""
    n, T = M.shape
    c = np.nan_to_num(costo, nan=0.0)
    origenes = range(W - 1, T - 1)
    mus = [M[:, : t + 1].mean(axis=1) for t in origenes]
    Ds = [M[:, t + 1] for t in origenes]

    def capital(k):
        tot = 0.0
        for mu, D in zip(mus, Ds):
            S = np.ceil(mu * (L + R) * k)
            tot += (c * (S + np.maximum(S - D, 0)) / 2).sum()
        return tot / len(mus)

    lo, hi = 0.0, 1.0
    while capital(hi) < inventario_objetivo and hi < 1e4:
        hi *= 2
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if capital(mid) < inventario_objetivo else (lo, mid)
    return (lo + hi) / 2


def backtest(M: np.ndarray, costo: np.ndarray, ved: np.ndarray, W: int = 12,
             L: float = 0.5, R: float = 0.5, delta_sl: float = 0.0,
             cortes=CORTES_ABC, k_b2: float | None = None) -> pd.DataFrame:
    """Evaluación de origen móvil (walk-forward) con H = L + R = 1 mes.

    Para cada mes de origen t se estiman los parámetros SOLO con datos hasta t
    (incluida la clasificación ABC, para evitar fuga de información) y se evalúan
    contra el consumo real del mes t+1. Devuelve una fila por origen x política x
    clase VED con las sumas necesarias para calcular las métricas.
    """
    assert abs(L + R - 1.0) < 1e-9, "El backtest asume horizonte H = L + R = 1 mes"
    n, T = M.shape
    c = np.nan_to_num(costo, nan=0.0)
    filas = []
    for t in range(W - 1, T - 1):
        win = M[:, t - W + 1: t + 1]
        hist = M[:, : t + 1]
        D = M[:, t + 1]
        valor = pd.Series(win.sum(axis=1) * c)
        abc = clasificar_abc(valor, cortes).to_numpy()
        pol = {
            "PROPUESTA": politica_propuesta(win, abc, ved, L, R, delta_sl),
            "B1_ultimo_mes": politica_ultimo_mes(hist, L, R),
            "B2_promedio_historico": politica_promedio_historico(hist, L, R),
            "A1_promedio_ventana": politica_promedio_ventana(win, L, R),
        }
        if k_b2 is not None:      # B2 escalada a capital equivalente al de la propuesta
            pol["B2_capital_equivalente"] = {"maximo": np.ceil(hist.mean(axis=1) * (L + R) * k_b2)}
        for nombre, p in pol.items():
            S = p["maximo"]
            cubierto = np.minimum(D, S)
            for clase in ["V", "E", "D", "TODAS"]:
                m = np.ones(n, bool) if clase == "TODAS" else (ved == clase)
                dem = m & (D > 0)
                filas.append({
                    "origen": t, "politica": nombre, "VED": clase,
                    "unid_demandadas": D[m].sum(), "unid_cubiertas": cubierto[m].sum(),
                    "series_con_demanda": dem.sum(), "series_cubiertas": (dem & (S >= D)).sum(),
                    "valor_demandado": (D * c)[m].sum(), "valor_cubierto": (cubierto * c)[m].sum(),
                    "valor_inventario_prom": (c * (S + np.maximum(S - D, 0)) / 2)[m].sum(),
                    "unid_inventario_prom": ((S + np.maximum(S - D, 0)) / 2)[m].sum(),
                })
    return pd.DataFrame(filas)


def metricas_backtest(res: pd.DataFrame) -> pd.DataFrame:
    """Métricas globales por política y clase VED a partir del resultado del backtest."""
    g = res.groupby(["politica", "VED"], as_index=False).sum(numeric_only=True)
    g["fill_rate_unidades_%"] = 100 * g["unid_cubiertas"] / g["unid_demandadas"]
    g["fill_rate_valor_%"] = 100 * g["valor_cubierto"] / g["valor_demandado"]
    g["disponibilidad_ante_demanda_%"] = 100 * g["series_cubiertas"] / g["series_con_demanda"]
    n_orig = res["origen"].nunique()
    g["inventario_prom_valor_mensual"] = g["valor_inventario_prom"] / n_orig
    return g[["politica", "VED", "fill_rate_unidades_%", "fill_rate_valor_%",
              "disponibilidad_ante_demanda_%", "inventario_prom_valor_mensual"]]


def serie_mensual_metrica(res: pd.DataFrame, clase: str = "TODAS") -> pd.DataFrame:
    """Fill rate (valor) e inventario promedio por origen y política (para IC y pruebas)."""
    r = res[res["VED"] == clase].copy()
    r["fill_valor"] = r["valor_cubierto"] / r["valor_demandado"]
    return r.pivot(index="origen", columns="politica", values=["fill_valor", "valor_inventario_prom"])


def ic_bootstrap(x: np.ndarray, B: int = 5000, alfa: float = 0.05, seed: int = 42) -> tuple[float, float, float]:
    """Media e intervalo percentil por bootstrap sobre los meses de origen."""
    x = np.asarray(x, dtype=float)
    rng = np.random.default_rng(seed)
    medias = rng.choice(x, size=(B, len(x)), replace=True).mean(axis=1)
    return x.mean(), np.quantile(medias, alfa / 2), np.quantile(medias, 1 - alfa / 2)


def prueba_wilcoxon(a: np.ndarray, b: np.ndarray) -> float:
    """p-valor de la prueba de rangos con signo (muestras pareadas por mes)."""
    d = np.asarray(a) - np.asarray(b)
    if np.allclose(d, 0):
        return 1.0
    return float(wilcoxon(d).pvalue)


# --------------------------------------------------------------------------
# Estabilidad temporal del consumo (ajuste al ACA 1)
# --------------------------------------------------------------------------
def descomposicion_caida(agr: pd.DataFrame, meta: pd.DataFrame, por: str,
                         periodos_a: list[str], periodos_b: list[str]) -> pd.DataFrame:
    """Cambio de consumo entre dos conjuntos de períodos equivalentes, agrupado
    por una variable de `meta` (CONCEPTO, SERVICIO_FARMACEUTICO, ...)."""
    d = agr.merge(meta, on="CODIGO_PROPIO", how="left") if por not in agr.columns else agr
    d[por] = d[por].fillna("Sin dato")
    a = d[d["PERIODO"].isin(periodos_a)].groupby(por)["CANTIDAD"].sum()
    b = d[d["PERIODO"].isin(periodos_b)].groupby(por)["CANTIDAD"].sum()
    t = pd.concat([a.rename("consumo_a"), b.rename("consumo_b")], axis=1).fillna(0)
    t["cambio_unidades"] = t["consumo_b"] - t["consumo_a"]
    t["cambio_%"] = np.where(t["consumo_a"] > 0, 100 * t["cambio_unidades"] / t["consumo_a"], np.nan)
    tot = t["cambio_unidades"].sum()
    t["aporte_al_cambio_total_%"] = 100 * t["cambio_unidades"] / tot if tot else np.nan
    return t.sort_values("cambio_unidades")


def concentracion_caida(agr: pd.DataFrame, periodos_a: list[str], periodos_b: list[str]) -> dict:
    """¿La caída fue generalizada o concentrada? Evalúa cuántos productos bajan y
    qué parte de la caída neta explican los que más bajan."""
    a = agr[agr["PERIODO"].isin(periodos_a)].groupby("CODIGO_PROPIO")["CANTIDAD"].sum()
    b = agr[agr["PERIODO"].isin(periodos_b)].groupby("CODIGO_PROPIO")["CANTIDAD"].sum()
    t = pd.concat([a.rename("a"), b.rename("b")], axis=1).fillna(0)
    t["d"] = t["b"] - t["a"]
    bajan, suben = t[t["d"] < 0], t[t["d"] > 0]
    neta = t["d"].sum()
    bajas_ord = bajan["d"].sort_values()
    k1, k10 = max(1, int(len(t) * 0.01)), max(1, int(len(t) * 0.10))
    return {
        "productos_total": len(t), "productos_que_bajan": len(bajan),
        "productos_que_suben": len(suben), "productos_sin_cambio": int((t["d"] == 0).sum()),
        "pct_productos_que_bajan": round(100 * len(bajan) / len(t), 2),
        "caida_neta_unidades": float(neta),
        "pct_caida_bruta_top1pct_productos": round(100 * bajas_ord.iloc[:k1].sum() / bajan["d"].sum(), 2) if len(bajan) else np.nan,
        "pct_caida_bruta_top10pct_productos": round(100 * bajas_ord.iloc[:k10].sum() / bajan["d"].sum(), 2) if len(bajan) else np.nan,
    }
