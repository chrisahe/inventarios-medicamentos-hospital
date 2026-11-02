# %% [markdown]
# # 01 · Diccionario de datos y ficha técnica de las fuentes
# Ajuste solicitado en la retroalimentación del ACA 1 (pasos 3 y 4).
# Genera, a partir de los archivos institucionales, el diccionario de datos
# (variable, tipo, descripción, % de nulos, valores únicos, rango/ejemplo) y la
# ficha técnica de cada fuente. No contiene datos de pacientes.

# %%
import os, sys
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else ".", "..", "src"))
sys.path.insert(0, "src")
try:
    sys.stdout.reconfigure(encoding="utf-8")   # consola de Windows
except Exception:
    pass
import inventario as inv

DATA_DIR = os.environ.get("DATA_DIR", "data")            # carpeta con los .xlsx
OUT = os.environ.get("OUT_DIR", "outputs/tablas")
os.makedirs(OUT, exist_ok=True)

# %% [markdown]
# ## 1. Descripciones de variables (definidas con base en las fuentes y el ACA 1)

# %%
DESC = {
 "Biblia": {
  "CUM": "Código Único de Medicamentos (INVIMA) o código del dispositivo médico.",
  "ATC": "Clasificación Anatómica, Terapéutica y Química (OMS) del medicamento.",
  "BODEGA": "Código de la bodega en la que el producto está parametrizado.",
  "CODIGO_PROPIO": "Código interno del producto en el sistema institucional (llave de unión).",
  "DES_MEDICAMENTO": "Descripción del medicamento o dispositivo médico.",
  "CONCEPTO": "Tipo de producto: Medicamentos POS, Medicamentos NO POS, Materiales e Insumos.",
  "ES_ACTIVO": "Indica si el producto está activo en la bodega (True/False).",
  "MOTIVO_INACTIVACION": "Motivo de inactivación ('0' = sin motivo; 'Invima vencido'; códigos 20, 22, 23, 24 sin descripción en la fuente, por confirmar).",
  "COSTO_FINAL": "Costo unitario final del producto (COP, por confirmar con el área financiera).",
  "REGISTRO_SANITARIO": "Registro sanitario INVIMA.",
  "VEN": "Clasificación institucional de criticidad: V vital, E esencial, N no esencial.",
 },
 "Bodegas": {
  "Bodega": "Código de la bodega (llave de unión con el consumo).",
  "Servicio": "Nombre del servicio o punto de almacenamiento.",
  "Servicio Farmaceutico": "Servicio farmacéutico al que pertenece la bodega.",
  "Sede": "Sede donde está ubicada la bodega.",
  "Se utiliza": "Indica si la bodega se usa para el análisis (Si / No / condicional).",
 },
 "Disponibilidad": [  # por posición; los encabezados de origen traen caracteres dañados
  ("CANTIDAD", "Unidades disponibles por lote en la fecha de corte."),
  ("CODIGO", "Código de 13 dígitos del producto-lote (no coincide con CODIGO_PROPIO)."),
  ("CODIGO_GENERICO", "Código genérico; coincide con CODIGO_PROPIO de la Biblia."),
  ("LOTE", "Número de lote."),
  ("FECHA_VENCIMIENTO", "Fecha de vencimiento del lote."),
  ("NOMBRE_COMERCIAL", "Nombre comercial."),
  ("PRINCIPIO_ACTIVO_CONCENTRACION", "Principio activo y concentración."),
  ("PRESENTACION_COMERCIAL", "Presentación comercial."),
  ("CLASIFICACION_RIESGO", "Clasificación de riesgo del dispositivo (I, IIA, IIB, III)."),
  ("CONDICION_ALMACENAMIENTO", "Condición de almacenamiento requerida."),
  ("LABORATORIO", "Laboratorio o fabricante."),
  ("REG_INVIMA", "Registro INVIMA."),
  ("FORMA_FARMACEUTICA", "Forma farmacéutica."),
  ("FOTOSENSIBILIDAD", "Indica si el producto es fotosensible."),
  ("ESTABILIDAD", "Indica si requiere control de estabilidad."),
  ("VENCIDO", "Indica si el lote está vencido (Si/No)."),
  ("STOCK_EN_MINIMOS", "Indicador del sistema: existencias en nivel mínimo (Si/No)."),
  ("STOCK_EN_MAXIMOS", "Indicador del sistema: existencias en nivel máximo (Si/No)."),
  ("BODEGA", "Nombre de la bodega donde está el lote."),
  ("GRUPO", "Grupo del producto (Medicamentos, Dispositivo médico, Materiales...)."),
  ("SUBGRUPO", "Subgrupo del producto."),
  ("CUM", "Código Único de Medicamentos."),
  ("ALERTA_STOCK", "Texto del sistema con existencias y niveles mínimo y máximo parametrizados."),
 ],
 "Novedades": {
  "CODIGO_PROPIO": "Código interno del producto.",
  "DES_MEDICAMENTO": "Descripción del producto.",
  "NOVEDAD": "Tipo de novedad: Invima vencido, Solo se solicita a necesidad, Ya no se utiliza, Desabastecido, Siempre deben haber existencias.",
  "FECHA DE REPORTE": "Fecha de reporte de la novedad ('No aplica' en la mayoría de los registros).",
 },
 "Base analítica (consumo)": {
  "CODIGO_PROPIO": "Código interno del producto.",
  "DES_MEDICAMENTO": "Descripción del producto.",
  "CANTIDAD": "Unidades consumidas (salidas) en el registro; solo valores positivos tras la depuración.",
  "BODEGA": "Código de la bodega de origen del movimiento.",
  "DES_UF": "Unidad funcional o asistencial donde se registra el consumo.",
  "ANIO": "Año del registro.",
  "MES": "Mes del registro (1-12).",
  "SERVICIO_FARMACEUTICO": "Servicio farmacéutico derivado de la bodega con la base maestra de Bodegas.",
  "PERIODO": "Variable derivada AAAA-MM.",
 },
}

# %% [markdown]
# ## 2. Perfilado de cada fuente

# %%
def perfil(df: pd.DataFrame, fuente: str, desc: dict) -> pd.DataFrame:
    filas = []
    for c in df.columns:
        s = df[c]
        num = pd.to_numeric(s, errors="coerce")
        es_num = pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s)
        if pd.api.types.is_datetime64_any_dtype(s):
            tipo, ej = "fecha", f"{s.min():%Y-%m-%d} a {s.max():%Y-%m-%d}"
        elif es_num:
            tipo, ej = "numérico", f"{num.min():,.0f} a {num.max():,.0f}"
        else:
            tipo = "texto/categórico"
            ej = "; ".join(map(str, s.dropna().astype(str).value_counts().index[:3]))[:80]
        filas.append({"Fuente": fuente, "Variable": c, "Tipo": tipo,
                      "Descripción": desc.get(c, ""), "% nulos": round(100 * s.isna().mean(), 2),
                      "Valores únicos": s.nunique(dropna=True), "Rango / ejemplos": ej})
    return pd.DataFrame(filas)

fuentes, fichas, dic = {}, [], []

b = pd.read_excel(f"{DATA_DIR}/Biblia (1).xlsx", dtype=str)
n_b_raw, n_dup = len(b), int(b.duplicated().sum())
b = inv.leer_biblia(f"{DATA_DIR}/Biblia (1).xlsx")
b_perf = b.copy(); b_perf["ES_ACTIVO"] = b_perf["ES_ACTIVO"].astype(str)
dic.append(perfil(b_perf, "Biblia", DESC["Biblia"]))
fichas.append(["Biblia (base maestra de productos)", "Biblia (1).xlsx",
               "Sistema institucional de inventarios / maestro de productos", n_b_raw, b.shape[1],
               f"{n_b_raw:,} filas; {n_dup:,} duplicados exactos y {n_b_raw-n_dup-len(b):,} fila sin código eliminados -> {len(b):,}; "
               f"{b.CODIGO_PROPIO.nunique():,} códigos propios únicos",
               "CODIGO_PROPIO (+ BODEGA para el costo)", "Costo unitario y clasificación VEN"])

bo = pd.read_excel(f"{DATA_DIR}/bodegas.xlsx")
dic.append(perfil(bo, "Bodegas", DESC["Bodegas"]))
fichas.append(["Bodegas (maestro de bodegas)", "bodegas.xlsx", "Maestro institucional de bodegas",
               len(bo), bo.shape[1], f"{len(bo)} bodegas; {bo['Servicio Farmaceutico'].nunique()} servicios farmacéuticos",
               "Bodega", "Asignar servicio farmacéutico y depurar el universo"])

d = pd.read_excel(f"{DATA_DIR}/Disponibilidad.xlsx")
d.columns = [n for n, _ in DESC["Disponibilidad"]]
dic.append(perfil(d, "Disponibilidad", dict(DESC["Disponibilidad"])))
fichas.append(["Disponibilidad (inventario a fecha de corte)", "Disponibilidad.xlsx",
               "Sistema institucional de inventarios", len(d), d.shape[1],
               f"{len(d):,} filas producto-lote; {d.CODIGO_GENERICO.nunique():,} productos; fecha de corte NO registrada en el archivo",
               "CODIGO_GENERICO = CODIGO_PROPIO", "Línea base de inventario actual"])

nv = pd.read_excel(f"{DATA_DIR}/Novedades.xlsx", dtype=str)
dic.append(perfil(nv, "Novedades", DESC["Novedades"]))
fichas.append(["Novedades (estado de abastecimiento)", "Novedades.xlsx",
               "Registro institucional de novedades de abastecimiento", len(nv), nv.shape[1],
               f"{len(nv):,} filas; {nv.iloc[:,0].nunique():,} productos; 95,5 % sin fecha de reporte válida",
               "CODIGO_PROPIO", "Reglas de negocio vigentes (no explicación temporal)"])

# Base analítica: solo si el archivo exportado del cuaderno del Objetivo 1 existe
ruta_base = os.environ.get("BASE_ANALITICA", f"{DATA_DIR}/base_analitica.parquet")
if os.path.exists(ruta_base):
    base = inv.leer_base_analitica(ruta_base)
    dic.append(perfil(base, "Base analítica (consumo)", DESC["Base analítica (consumo)"]))
    fichas.append(["Base analítica de consumo", os.path.basename(ruta_base),
                   "Salidas mensuales del sistema institucional (2024-01 a 2026-05)", len(base), base.shape[1],
                   f"{len(base):,} registros depurados", "CODIGO_PROPIO, BODEGA", "Consumo mensual por producto y servicio"])
else:
    print("Aviso: no se encontró la base analítica; se documentan solo las fuentes maestras.")

diccionario = pd.concat(dic, ignore_index=True)
ficha = pd.DataFrame(fichas, columns=["Fuente", "Archivo", "Origen", "Filas (brutas)", "Columnas",
                                      "Contenido y volumen", "Llave de unión", "Uso en el proyecto"])

# %%
diccionario.to_csv(f"{OUT}/diccionario_datos.csv", index=False, encoding="utf-8-sig")
ficha.to_csv(f"{OUT}/ficha_tecnica_fuentes.csv", index=False, encoding="utf-8-sig")
with pd.ExcelWriter(f"{OUT}/diccionario_y_ficha.xlsx") as w:
    ficha.to_excel(w, sheet_name="Ficha técnica", index=False)
    diccionario.to_excel(w, sheet_name="Diccionario", index=False)
print(ficha[["Fuente", "Contenido y volumen"]].to_string(index=False))
print(diccionario.groupby("Fuente").size())
