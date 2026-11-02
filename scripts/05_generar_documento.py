"""Genera el contenido nuevo del documento de grado (ajustes del ACA 1 + Secciones 8.2 y 8.3) en Word.

Lee las tablas y figuras que producen los scripts 01-04 (carpeta `outputs/`) y escribe cada cifra
desde ellas, de modo que el texto no contiene números tecleados a mano. Lo que aún no se ha
calculado queda como un marcador amarillo [PENDIENTE: ...] con el script que lo produce.

Uso:  python scripts/05_generar_documento.py [salida.docx]
Variables de entorno: OUT_TAB, OUT_FIG (carpetas de resultados), REPO_URL, COLAB_URL.
"""
import json, os, re, sys
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

TAB = os.environ.get("OUT_TAB", "outputs/tablas")
FIG = os.environ.get("OUT_FIG", "outputs/figuras")
REPO = os.environ.get("REPO_URL", "[URL del repositorio de GitHub]")
COLAB = os.environ.get("COLAB_URL", "[URL del cuaderno de Google Colab]")
SALIDA = sys.argv[1] if len(sys.argv) > 1 else "ACA2_contenido_nuevo.docx"
TAB_N, FIG_N = [28], [42]          # la numeración continúa la del ACA 1 (última: Tabla 27, Figura 41)

doc = Document()
st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.3
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.5)


# ------------------------------------------------------------------ utilidades
def es(x, d=0):
    """Formato numérico español: 1.234.567,89"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "n. d."
    return f"{x:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def pesos(x):
    return "$" + es(x)


def leer(nombre, **kw):
    ruta = os.path.join(TAB, nombre)
    return pd.read_csv(ruta, **kw) if os.path.exists(ruta) else None


def cifras(nombre):
    ruta = os.path.join(TAB, nombre)
    return json.load(open(ruta, encoding="utf-8")) if os.path.exists(ruta) else {}


def h(texto, nivel):
    doc.add_heading(texto, level=nivel)


def p(texto, negrita_inicio=None):
    par = doc.add_paragraph()
    if negrita_inicio:
        par.add_run(negrita_inicio).bold = True
    par.add_run(texto)
    par.paragraph_format.first_line_indent = Cm(0)
    return par


def pendiente(texto):
    par = doc.add_paragraph()
    r = par.add_run(f"[PENDIENTE: {texto}]")
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return par


def bullets(items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


def _borde(celda, **kw):
    tcPr = celda._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for lado, v in kw.items():
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), v); e.set(qn("w:sz"), "6"); e.set(qn("w:color"), "000000")
        b.append(e)
    tcPr.append(b)


def tabla(df, titulo, nota, anchos=None, fuente=9):
    """Tabla en estilo APA 7: número en negrita, título en cursiva, solo líneas horizontales."""
    n = TAB_N[0]; TAB_N[0] += 1
    c = doc.add_paragraph(); c.paragraph_format.keep_with_next = True
    c.add_run(f"Tabla {n}").bold = True
    t = doc.add_paragraph(); t.paragraph_format.keep_with_next = True
    t.add_run(titulo).italic = True
    tb = doc.add_table(rows=1, cols=len(df.columns))
    for j, col in enumerate(df.columns):
        cel = tb.rows[0].cells[j]; cel.text = ""
        r = cel.paragraphs[0].add_run(str(col)); r.bold = True; r.font.size = Pt(fuente)
        _borde(cel, top="single", bottom="single")
    for _, fila in df.iterrows():
        cels = tb.add_row().cells
        for j, v in enumerate(fila):
            cels[j].text = ""
            r = cels[j].paragraphs[0].add_run("" if (isinstance(v, float) and pd.isna(v)) else str(v))
            r.font.size = Pt(fuente)
    for cel in tb.rows[-1].cells:
        _borde(cel, bottom="single")
    if anchos:
        for fila in tb.rows:
            for j, w in enumerate(anchos):
                fila.cells[j].width = Cm(w)
    q = doc.add_paragraph()
    q.add_run("Nota. ").italic = True
    q.add_run(nota)
    for r in q.runs:
        r.font.size = Pt(10)
    return n


def figura(archivo, titulo, nota, ancho=14):
    ruta = os.path.join(FIG, archivo)
    if not os.path.exists(ruta):
        return pendiente(f"figura {archivo} (se genera con los scripts)")
    n = FIG_N[0]; FIG_N[0] += 1
    c = doc.add_paragraph(); c.paragraph_format.keep_with_next = True
    c.add_run(f"Figura {n}").bold = True
    t = doc.add_paragraph(); t.paragraph_format.keep_with_next = True
    t.add_run(titulo).italic = True
    doc.add_picture(ruta, width=Cm(ancho))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    q = doc.add_paragraph()
    q.add_run("Nota. ").italic = True
    q.add_run(nota)
    for r in q.runs:
        r.font.size = Pt(10)
    return n


def df_es(df, dec=None, cols_pct=()):
    """Copia con números en formato español."""
    d = df.copy()
    for c in d.columns:
        if pd.api.types.is_numeric_dtype(d[c]):
            k = (dec or {}).get(c, 2 if c in cols_pct else 0)
            d[c] = d[c].map(lambda v: es(v, k))
    return d


# ================================================================ PARTE A · AJUSTES AL ACA 1
h("PARTE A. Ajustes al ACA 1", 1)
p("Esta parte reúne el contenido que responde a cada observación de la retroalimentación del ACA 1. "
  "Se indica dónde insertarlo en el documento.")

# ---- A1 Disponibilidad de datos y código
h("Disponibilidad de datos y código (insertar antes de la Sección 8)", 2)
p("El código desarrollado en el proyecto es público y puede ejecutarse de forma independiente. El repositorio de GitHub "
  f"({REPO}) contiene las funciones, los scripts numerados de cada fase, los cuadernos para Google Colab ({COLAB}), el archivo "
  "de requerimientos, el protocolo de evaluación y la bitácora de decisiones. Los diagramas de flujo de esta sección "
  "reemplazan las capturas de código del documento anterior; el código completo se encuentra en el repositorio.")
p("Los datos son institucionales y no se publican, por tratarse de información interna del hospital. Antes de cualquier "
  "procesamiento se retiraron los campos de pacientes y ventas. La Ficha técnica (Tabla 28) y el Diccionario de datos "
  "(Tabla 29) describen cada fuente. Para verificar que el código se ejecuta, el repositorio incluye un generador de datos "
  "sintéticos; los resultados obtenidos con esos datos no representan al hospital.")
p("Declaración de uso de inteligencia artificial: se utilizó el asistente Claude (Anthropic) como apoyo para estructurar el código "
  "y los borradores de texto. Los autores verificaron las cifras contra la ejecución del código, las referencias contra sus fuentes "
  "originales y responden por cada decisión registrada en la bitácora (Anexo B).")

# ---- A2 Ficha técnica y diccionario
h("Ficha técnica de las fuentes y diccionario de datos (Anexo A)", 2)
ft = leer("ficha_tecnica_fuentes.csv")
dd = leer("diccionario_datos.csv")
if ft is not None:
    f = ft[["Fuente", "Archivo", "Filas (brutas)", "Columnas", "Contenido y volumen", "Uso en el proyecto"]].copy()
    f["Filas (brutas)"] = f["Filas (brutas)"].map(es)
    tabla(f, "Ficha técnica de las fuentes de información",
          "Las filas brutas corresponden al archivo original. La fecha de corte de Disponibilidad no está registrada en el archivo "
          "y debe confirmarse con la institución. Fuente: elaboración propia a partir de las bases institucionales.",
          fuente=8)
if dd is not None:
    d2 = dd.copy()
    d2["% nulos"] = d2["% nulos"].map(lambda v: es(v, 2)); d2["Valores únicos"] = d2["Valores únicos"].map(es)
    tabla(d2[["Fuente", "Variable", "Tipo", "Descripción", "% nulos", "Valores únicos"]],
          "Diccionario de datos de las fuentes maestras",
          "El porcentaje de nulos se calculó sobre cada fuente tras eliminar los duplicados exactos de la Biblia. "
          "Fuente: elaboración propia.", fuente=7.5)
    if not (dd["Fuente"] == "Base analítica (consumo)").any():
        pendiente("agregar las filas de la base analítica de consumo: ejecutar scripts/01_diccionario_y_ficha.py con data/base_analitica.parquet")

# ---- A3 Costo unitario
h("8.1.13 Análisis descriptivo del costo unitario (insertar al final del Objetivo 1)", 2)
cob = leer("costo_cobertura.csv"); desc = leer("costo_descriptivo.csv", index_col=0)
var = leer("costo_variacion_entre_bodegas.csv", index_col=0)["valor"]; ext = leer("costo_extremos.csv", index_col=0)["valor"]
top = leer("costo_top15.csv", dtype={"CODIGO_PROPIO": str})
if cob is not None and desc is not None:
    tot = int(cob["Productos"].sum()); sin = int(cob.loc[cob["Fuente del costo"] == "sin_costo", "Productos"].sum())
    act = cob.set_index("Fuente del costo")
    p(f"El Objetivo 2 requiere valorar el consumo, por lo que se caracterizó el costo unitario registrado en la base maestra de "
      f"productos (Biblia). Tras eliminar los registros duplicados, la Biblia quedó con {es(tot)} productos únicos. Para "
      f"{es(tot - sin)} de ellos ({es(100 * (tot - sin) / tot, 2)} %) hay un costo registrado: {es(act.loc['codigo_activo', 'Productos'])} "
      f"({es(act.loc['codigo_activo', '%'], 2)} %) lo tienen en registros activos y {es(act.loc['codigo_inactivo', 'Productos'])} "
      f"({es(act.loc['codigo_inactivo', '%'], 2)} %) solo en registros inactivos. Los {es(sin)} productos restantes ({es(act.loc['sin_costo', '%'], 2)} %) "
      "no tienen costo y no pueden entrar al análisis ABC.")
    t = desc.loc[["Todos", "Materiales e Insumos", "Medicamentos NO POS", "Medicamentos POS", "V", "E", "N"],
                 ["n", "media", "mediana", "p90", "p99", "máximo", "CV_%"]].reset_index().rename(columns={"index": "Grupo"})
    t["Grupo"] = t["Grupo"].replace({"V": "Vital (V)", "E": "Esencial (E)", "N": "No esencial (N)"})
    t.columns = ["Grupo", "n", "Media", "Mediana", "P90", "P99", "Máximo", "CV (%)"]
    tabla(df_es(t, {"CV (%)": 1}), "Estadísticos descriptivos del costo unitario (COP) por tipo de producto y criticidad",
          "El costo corresponde a la mediana de los registros activos de cada producto o, en su defecto, de los inactivos. "
          "CV: coeficiente de variación. Fuente: elaboración propia a partir de la Biblia.", fuente=8.5)
    todos = desc.loc["Todos"]
    p(f"La distribución del costo es muy asimétrica: el promedio ({pesos(todos['media'])}) es {es(todos['media'] / todos['mediana'], 1)} veces la mediana "
      f"({pesos(todos['mediana'])}) y el coeficiente de variación alcanza {es(todos['CV_%'], 0)} %. El costo máximo es {pesos(todos['máximo'])}. "
      f"En escala logarítmica ningún producto supera el límite de valores extremos de Tukey ({es(ext['productos_extremos'])} productos), por lo que no se identificaron "
      "valores que sugieran errores de registro. La mediana de los Materiales e Insumos "
      f"({pesos(desc.loc['Materiales e Insumos', 'mediana'])}) es mayor que la de los Medicamentos NO POS ({pesos(desc.loc['Medicamentos NO POS', 'mediana'])}) y la de los Medicamentos POS "
      f"({pesos(desc.loc['Medicamentos POS', 'mediana'])}).")
    figura("costo_distribucion.png", "Distribución del costo unitario de los productos",
           "Escala logarítmica base 10. Los diagramas de caja no muestran valores atípicos individuales. Fuente: elaboración propia.")
    p(f"El costo de un mismo producto también varía entre bodegas: {es(var['productos_con_mas_de_un_costo_activo'])} de {es(var['productos_con_costo_activo'])} productos con costo activo "
      f"tienen más de un costo, con una razón entre el mayor y el menor de {es(var['mediana_razon_max_min'], 2)} en la mediana y de {es(var['p90_razon_max_min'], 2)} en el percentil 90. "
      "Por esta razón, el Objetivo 2 asigna el costo con una regla jerárquica (Sección 8.2.2) y evalúa en qué medida la clasificación depende de ella.")
    if top is not None:
        tt = top.head(5)[["DES_MEDICAMENTO", "CONCEPTO", "VEN", "COSTO"]].copy()
        tt["COSTO"] = tt["COSTO"].map(pesos); tt.columns = ["Producto", "Tipo", "VEN", "Costo unitario"]
        tabla(tt, "Cinco productos de mayor costo unitario", "Fuente: elaboración propia a partir de la Biblia.", fuente=8.5)

# ---- A4 Estabilidad temporal
h("8.1.14 Estabilidad temporal del consumo (insertar al final del Objetivo 1)", 2)
est = cifras("est_cifras_clave.json")
if est:
    for k, lab in [("concentracion_ene_may_2024_vs_2025", "enero–mayo de 2024 y de 2025"),
                   ("concentracion_ene_may_2025_vs_2026", "enero–mayo de 2025 y de 2026")]:
        c = est[k]
        p(f"Entre {lab}, el consumo disminuyó en {es(c['productos_que_bajan'])} de {es(c['productos_total'])} productos "
          f"({es(c['pct_productos_que_bajan'], 2)} %) y aumentó en {es(c['productos_que_suben'])}. Los productos con mayor caída (1 % del total) explicaron "
          f"el {es(c['pct_caida_bruta_top1pct_productos'], 2)} % de la caída bruta y el 10 % con mayor caída, el {es(c['pct_caida_bruta_top10pct_productos'], 2)} %.")
    c1, c2e = est["concentracion_ene_may_2024_vs_2025"], est["concentracion_ene_may_2025_vs_2026"]
    cc_a, cc_b = leer("est_caida_concepto_2024_2025.csv", index_col=0), leer("est_caida_concepto_2025_2026.csv", index_col=0)
    sv_a, sv_b = leer("est_caida_servicio_2024_2025.csv", index_col=0), leer("est_caida_servicio_2025_2026.csv", index_col=0)
    ix = leer("est_indice_por_concepto.csv", index_col=0)
    p(f"Interpretación. La caída fue amplia en número de productos y concentrada en volumen. Amplia, porque más de la mitad de los productos disminuyó en cada "
      f"comparación ({es(c1['pct_productos_que_bajan'], 1)} % y {es(c2e['pct_productos_que_bajan'], 1)} %), aunque una parte importante aumentó "
      f"({es(c1['productos_que_suben'])} y {es(c2e['productos_que_suben'])} productos). Concentrada, porque una pequeña fracción de productos explica casi toda la caída en unidades "
      f"(el 1 % con mayor caída explica el {es(c1['pct_caida_bruta_top1pct_productos'], 1)} % y el {es(c2e['pct_caida_bruta_top1pct_productos'], 1)} % de la caída bruta); este resultado es "
      "esperable en una distribución de consumo muy asimétrica, en la que pocos productos de alto volumen determinan el total.")
    if cc_a is not None and cc_b is not None:
        gr = ["Materiales e Insumos", "Medicamentos POS", "Medicamentos NO POS"]
        p("Por tipo de producto, la disminución se presentó en los tres grupos principales en ambos períodos: " +
          "; ".join(f"{g}, {es(cc_a.loc[g, 'cambio_%'], 1)} % entre 2024 y 2025 y {es(cc_b.loc[g, 'cambio_%'], 1)} % entre 2025 y 2026" for g in gr) +
          f". Los Materiales e Insumos aportan el {es(cc_a.loc['Materiales e Insumos', 'aporte_al_cambio_total_%'], 1)} % de la caída neta del primer período y los Medicamentos POS "
          f"el {es(cc_b.loc['Medicamentos POS', 'aporte_al_cambio_total_%'], 1)} % de la del segundo, de modo que el predominio cambia de un tipo de producto a otro.")
    if sv_a is not None and "Almacen" in sv_a.index:
        ha_a = sv_a.loc[["Hospitalizacion", "Almacen"], ["consumo_a", "consumo_b"]].sum()
        p(f"Por servicio farmacéutico la comparación no es directa. Almacén no registra consumo en enero–mayo de 2024 y pasa a {es(sv_a.loc['Almacen', 'consumo_b'])} unidades en enero–mayo de 2025, "
          f"mientras que Hospitalización cae {es(-sv_a.loc['Hospitalizacion', 'cambio_%'], 1)} %. Al sumar ambos servicios, el cambio es de {es(ha_a['consumo_a'])} a {es(ha_a['consumo_b'])} unidades "
          f"({es(100 * (ha_a['consumo_b'] / ha_a['consumo_a'] - 1), 1)} %). Esto sugiere un cambio en la forma de registrar el consumo (parte de las salidas de los servicios satélite se registraría "
          "en Almacén desde 2025), hipótesis que debe confirmarse con el servicio farmacéutico; por ahora, las comparaciones por servicio entre años se interpretan con cautela y los niveles finales "
          "usan solo la ventana de los últimos 12 meses.")
    if ix is not None:
        m_ = ix["Materiales e Insumos"]
        p(f"Además, el índice mensual muestra descensos de un solo mes que no parecen cambios de demanda: octubre y noviembre de 2024 (Materiales e Insumos en {es(m_.loc['2024-10'], 1)} y "
          f"{es(m_.loc['2024-11'], 1)}, frente a un promedio de 2024 igual a 100) y diciembre de 2025 ({es(m_.loc['2025-12'], 1)}). Estos meses deben verificarse con el área de "
          "sistemas por si corresponden a registros incompletos de la extracción y no a una disminución real del consumo.")
    figura("est_indice_por_concepto.png", "Consumo mensual por tipo de producto (índice, promedio de 2024 = 100)", "Fuente: elaboración propia.")
else:
    pendiente("ejecutar scripts/02b_estabilidad_temporal.py con la base analítica para obtener los resultados de esta sección")
p("Implicación para el modelado: como el consumo de enero a mayo disminuyó entre 2024 y 2026, usar el promedio histórico completo "
  "como referencia de la demanda puede sobredimensionar el inventario. Por ello, los niveles del Objetivo 3 se estiman con una ventana móvil "
  "de los últimos 12 meses y el protocolo de evaluación incluye el promedio histórico como línea base (B2).")

# ---- A5 EDA → decisiones
h("8.1.15 Del análisis exploratorio a las decisiones de inventario (insertar al final del Objetivo 1)", 2)
dec = pd.DataFrame([
    ["Consumo sostenido", "893 (7,51 %)", "87,17 %", "Demanda continua; el CV es bajo o moderado",
     "Revisión periódica con modelo normal y stock de seguridad según el nivel de servicio de su celda ABC-VED"],
    ["Consumo frecuente", "1.024 (8,61 %)", "7,95 %", "Presencia en 12–23 meses; el CV puede superar 300 % (p. ej., clozapina 25 mg)",
     "Mismo modelo normal con ventana de 12 meses; el nivel de servicio por clase evita sobredimensionar los de baja criticidad"],
    ["Consumo intermitente", "2.482 (20,88 %)", "4,04 %", "Meses sin consumo; el supuesto normal no aplica",
     "Cuantil empírico del consumo mensual al nivel de servicio objetivo"],
    ["Consumo esporádico", "7.488 (62,99 %)", "0,84 %", "Consumo en 1–3 meses; CV de 400–600 %",
     "Productos vitales: existencia igual al mayor consumo mensual de la ventana. Esenciales y deseables: bajo demanda, sin stock de seguridad"],
], columns=["Patrón de frecuencia", "Productos", "% del consumo", "Característica", "Decisión en el modelo de inventario"])
tabla(dec, "Cómo los patrones de rotación alimentan el modelo de inventario",
      "Productos y participación del consumo tomados de las Tablas 23 y 24. Los productos esporádicos agrupan casi tres de cada cinco productos "
      "pero menos del 1 % del consumo: mantener existencia de ellos inmoviliza capital con muy poco beneficio en servicio, salvo que sean vitales. "
      "Fuente: elaboración propia.", fuente=8.5)
p("Las novedades institucionales se usan como reglas vigentes a la fecha del archivo (no como explicación temporal del consumo): "
  "'Ya no se utiliza' y 'Solo se solicita a necesidad' → sin niveles de inventario (bajo demanda); 'Siempre deben haber existencias' → clase V con mínimo de una unidad; "
  "'Invima vencido' → se marca para verificación antes de comprar; 'Desabastecido' → se marca como riesgo de abastecimiento.")

# ---- A6 Protocolo de evaluación
h("7.9 Protocolo de evaluación (insertar al final de la Metodología)", 2)
lb = leer("linea_base_inventario_actual.csv", index_col=0)
if lb is not None:
    lb = lb["valor"]
    p("Aunque el proyecto no entrena un modelo de aprendizaje automático, el éxito de la propuesta de inventario debe medirse contra la práctica actual. "
      f"El sistema institucional no tiene niveles mínimo y máximo definidos para la mayoría de los productos: en el archivo de disponibilidad, "
      f"el {es(lb['pct_producto_bodega_sin_min_max (0/0)'], 2)} % de las {es(lb['combinaciones_producto_bodega'])} combinaciones producto-bodega tiene mínimo y máximo en cero. "
      f"El inventario registrado asciende a {pesos(lb['valor_inventario_cop (lotes con costo)'])}, de los cuales {pesos(lb['valor_lotes_vencidos_cop'])} "
      f"({es(lb['pct_valor_vencido_sobre_inventario'], 2)} %) corresponden a {es(lb['lotes_vencidos'])} lotes vencidos. Por ello la línea base es el método empírico que se deduce de esa práctica.")
tabla(pd.DataFrame([
    ["B1", "Nivel máximo = consumo del último mes", "Reposición reactiva de lo consumido"],
    ["B2", "Nivel máximo = promedio de todo el histórico, sin stock de seguridad", "Promedio histórico simple"],
    ["A1 (ablación)", "Promedio de los últimos 12 meses, sin stock de seguridad ni reglas ABC-VED", "Aísla el aporte del stock de seguridad y de la clasificación"],
    ["B2k (control)", "B2 multiplicada por un factor para igualar el capital promedio de la propuesta", "Compara el servicio a capital equivalente"],
], columns=["Código", "Regla", "Qué representa"]), "Líneas base de la evaluación",
    "Pendiente de validar con la jefatura del servicio farmacéutico cómo se decide hoy la compra. Fuente: elaboración propia.")
tabla(pd.DataFrame([
    ["Fill rate (unidades y valor)", "Σ mín(D, S) / Σ D", "Proporción de la demanda que el nivel máximo S habría podido atender"],
    ["Disponibilidad ante demanda", "Series con D > 0 y S ≥ D / series con D > 0", "Proporción de productos-mes sin faltante"],
    ["Inventario promedio teórico (valor)", "Σ costo × [S + máx(S − D, 0)] / 2", "Capital inmovilizado teórico por mes"],
], columns=["Métrica", "Fórmula", "Lectura"]), "Métricas de la evaluación retrospectiva",
    "D: consumo real del mes siguiente; S: nivel máximo estimado con datos hasta el mes de origen. Fuente: elaboración propia.")
tabla(pd.DataFrame([
    ["C1", "Disponibilidad ante demanda ≥ 95 % en productos vitales"],
    ["C2", "Fill rate en valor de la propuesta mayor que el de B1 y el de B2"],
    ["C3", "A capital equivalente, fill rate de la propuesta mayor que el de B2k"],
], columns=["Criterio", "Condición"]), "Criterios de éxito del Objetivo 3",
    "Se definen antes de calcular los resultados y no se ajustan después. Los umbrales son propuestas de los autores sujetas a validación. Fuente: elaboración propia.")
p("La evaluación es de origen móvil: para cada mes de origen se estiman los parámetros solo con datos hasta ese mes (incluida la clasificación ABC) y se comparan con "
  "el consumo del mes siguiente (Hyndman & Athanasopoulos, 2021). Las novedades no se aplican hacia atrás para evitar fuga de información. "
  "Se reportan intervalos de confianza del 95 % por bootstrap sobre los meses y la prueba de Wilcoxon pareada. El protocolo completo está en el repositorio.")
figura("flujo_04_evaluacion.png", "Protocolo de evaluación retrospectiva de origen móvil", "Fuente: elaboración propia.", ancho=11)

# ---- A7 Diagramas del Objetivo 1
h("Diagramas de flujo que reemplazan las capturas de código del Objetivo 1", 2)
figura("flujo_01_etl_objetivo1.png", "Flujo de extracción, depuración y construcción de la base analítica (reemplaza las Figuras 1 a 23)",
       "El código ejecutable se encuentra en el repositorio. Fuente: elaboración propia.", ancho=11)
figura("flujo_05_arquitectura.png", "Arquitectura de la solución", "Fuente: elaboración propia.", ancho=11)

# ================================================================ PARTE B · OBJETIVO 2
doc.add_page_break()
h("PARTE B. Desarrollo del proyecto (continuación de la Sección 8)", 1)
p("Nota de numeración: el Objetivo 1 pasa a ser la Sección 8.1 (sus apartados actuales 8.1 a 8.12 se renumeran 8.1.1 a 8.1.12), "
  "el Objetivo 2 es la Sección 8.2 y el Objetivo 3 es la Sección 8.3.")
h("8.2 Desarrollo del Objetivo Específico 2", 2)
p("Clasificar los medicamentos y dispositivos médicos según su impacto económico, criticidad clínica y comportamiento de consumo, "
  "mediante metodologías de priorización de inventarios.", "Objetivo específico 2. ")

c2 = cifras("o2_cifras_clave.json")
h("8.2.1 Diseño metodológico", 3)
p("La clasificación combina dos criterios: el impacto económico (ABC), calculado como el valor del consumo, y la criticidad clínica (VED), tomada de la clasificación "
  "institucional de la Biblia. El cruce de ambos produce la matriz ABC-VED, que diferencia el nivel de control y de servicio que requiere cada producto "
  "(Devnani et al., 2010). El comportamiento del consumo definido en el Objetivo 1 (patrones de frecuencia) se cruza con la clasificación para validarla.")
figura("flujo_02_clasificacion_abc_ved.png", "Flujo de clasificación ABC-VED", "Fuente: elaboración propia.", ancho=11)
tabla(pd.DataFrame([
    ["Valor de consumo", "Cantidad × costo unitario vigente", "La Biblia no tiene historia de precios; un precio único evita que los cambios de precio alteren el ranking"],
    ["Ventana del ABC", "Últimos 12 meses (jun-2025 a may-2026)", "El consumo disminuyó en 2025-2026; la ventana reciente refleja la demanda vigente. Se evalúa también con 29 meses"],
    ["Cortes ABC", "A hasta 80 %, B hasta 95 %, C el resto del valor acumulado", "Convención más usada; se evalúa 70/90 y 75/95"],
    ["VED", "V vital, E esencial, N→D deseable, según VEN de la Biblia", "Criterio clínico institucional existente"],
    ["Sin clasificación VEN", "Se asigna E y se marca para revisión", "Criterio conservador: no subestimar la criticidad"],
], columns=["Parámetro", "Valor", "Justificación"]), "Parámetros de diseño de la clasificación",
    "Las decisiones están registradas en la bitácora (Anexo B). Fuente: elaboración propia.", fuente=8.5)

h("8.2.2 Preparación de variables financieras y clínicas", 3)
p("Costo unitario. Para cada registro de consumo se asigna el costo con una regla jerárquica: (1) mediana de los costos activos del producto en la misma bodega; "
  "(2) si no existe, mediana de los costos activos del producto en cualquier bodega; (3) si no existe, mediana de los costos de registros inactivos; "
  "(4) si el producto no tiene costo, queda sin costo y se excluye del ABC, informando su peso en unidades.")
cc = leer("o2_cobertura_costo.csv", index_col=0)
if cc is not None:
    t = cc.reset_index(); t.columns = ["Fuente del costo", "Filas", "Unidades", "% filas", "% unidades"]
    t["Fuente del costo"] = t["Fuente del costo"].replace({"1_bodega_activo": "1. Activo, misma bodega", "2_codigo_activo": "2. Activo, otra bodega",
                                                           "3_codigo_inactivo": "3. Registro inactivo", "4_sin_costo": "4. Sin costo"})
    tabla(df_es(t, {"% filas": 2, "% unidades": 2}), "Cobertura de la asignación de costo sobre el consumo",
          "Filas: combinaciones producto-bodega-servicio-período. Fuente: elaboración propia.")
    if c2:
        p(f"{es(c2['productos_sin_costo'])} de {es(c2['productos_total'])} productos del consumo no tienen costo en la Biblia.")
else:
    pendiente("cobertura de costo sobre el consumo (scripts/03_objetivo2_clasificacion_abc_ved.py)")
p("Criticidad clínica. La Biblia incluye una clasificación institucional VEN (vital, esencial, no esencial) para cada producto, sin productos con dos clasificaciones distintas. "
  "Para ajustarla al método VED se homologa N (no esencial) a D (deseable).")
tabla(pd.DataFrame([
    ["V (vital)", "Su ausencia pone en riesgo la vida del paciente o puede causar un daño grave, por lo que no puede faltar"],
    ["E (esencial)", "Se usa en el tratamiento de enfermedades importantes pero menos graves; su falta afecta la calidad de la atención sin riesgo vital inmediato"],
    ["D (deseable)", "Tiene alternativas o se usa en situaciones menores; su falta no compromete la continuidad de la atención"],
], columns=["Clase", "Criterio"]), "Criterios de criticidad clínica (VED)",
    "Definiciones propias basadas en el enfoque ABC-VED de Devnani et al. (2010). La asignación a cada producto es la de la Biblia. Fuente: elaboración propia.")
vt, vc = leer("ven_total.csv"), leer("ven_por_concepto_pct.csv", index_col=0)
if vt is not None:
    v = vt.set_index("VEN")
    p(f"En la Biblia, {es(v.loc['E', '%'], 2)} % de los productos son esenciales, {es(v.loc['V', '%'], 2)} % vitales y {es(v.loc['N', '%'], 2)} % no esenciales. "
      f"Los vitales representan el {es(vc.loc['Materiales e Insumos', 'V'], 2)} % de los Materiales e Insumos, pero solo el {es(vc.loc['Medicamentos POS', 'V'], 2)} % de los Medicamentos POS "
      f"y el {es(vc.loc['Medicamentos NO POS', 'V'], 2)} % de los Medicamentos NO POS. Esta distribución, muy concentrada en E y diferente entre tipos de producto, debe validarse "
      "con el químico farmacéutico (ver casos para revisión en la Sección 8.2.6).")

h("8.2.3 Clasificación ABC", 3)
a12 = leer("o2_abc_resumen_12m.csv", index_col=0)
if a12 is not None:
    t = a12.reset_index(); t.columns = ["Clase", "Productos", "Valor de consumo (COP)", "% productos", "% valor"]
    tabla(df_es(t, {"% productos": 2, "% valor": 2}), "Clasificación ABC por valor de consumo de los últimos 12 meses",
          f"Ventana: {c2['ventana_abc'][0]} a {c2['ventana_abc'][1]}. Se excluyen {es(c2['productos_sin_valor_12m'])} productos sin valor en la ventana. Fuente: elaboración propia." if c2 else "Fuente: elaboración propia.")
    A, B_, C = a12.loc["A"], a12.loc["B"], a12.loc["C"]
    p(f"La clase A reúne {es(A['productos'])} productos ({es(A['pct_productos'], 2)} %) que concentran el {es(A['pct_valor'], 2)} % del valor; la clase B, {es(B_['productos'])} productos "
      f"({es(B_['pct_productos'], 2)} %) con el {es(B_['pct_valor'], 2)} %; y la clase C, {es(C['productos'])} productos ({es(C['pct_productos'], 2)} %) con el {es(C['pct_valor'], 2)} %.")
    figura("o2_pareto.png", "Curva de Pareto del valor de consumo", "Las líneas punteadas marcan los cortes de 80 % y 95 %. Fuente: elaboración propia.", ancho=10)
else:
    pendiente("resultados del ABC: ejecutar scripts/03_objetivo2_clasificacion_abc_ved.py")

h("8.2.4 Clasificación VED", 3)
vr = leer("o2_ved_resumen.csv", index_col=0)
if vr is not None:
    t = vr.reset_index(); t.columns = ["Clase", "Productos", "Valor de consumo (COP)", "% productos", "% valor"]
    tabla(df_es(t, {"% productos": 2, "% valor": 2}), "Clasificación VED de los productos con consumo",
          "Fuente: elaboración propia a partir de la Biblia." )
    fv = c2.get("ved_fuente", {})
    if fv:
        p("Origen de la clasificación: " + "; ".join(f"{k}: {es(v)} productos" for k, v in fv.items()) + ".")
else:
    pendiente("resultados del VED (script 03)")

h("8.2.5 Integración en la matriz ABC-VED", 3)
mp, mv = leer("o2_matriz_productos.csv", index_col=0), leer("o2_matriz_pct_valor.csv", index_col=0)
if mp is not None:
    t = mp.reset_index().rename(columns={"ABC": "ABC \\ VED"}); tabla(df_es(t), "Número de productos por celda de la matriz ABC-VED", "Fuente: elaboración propia.")
    t = mv.reset_index().rename(columns={"ABC": "ABC \\ VED"}); tabla(df_es(t, {"V": 2, "E": 2, "D": 2}), "Porcentaje del valor de consumo por celda de la matriz ABC-VED", "Fuente: elaboración propia.")
    figura("o2_matriz_abc_ved.png", "Matriz ABC-VED", "Fuente: elaboración propia.", ancho=13)
else:
    pendiente("matriz ABC-VED (script 03)")
tabla(pd.DataFrame([
    ["AV, BV", "98 %", "Control estricto y seguimiento permanente; no debe faltar"],
    ["CV", "99 %", "Existencia amplia: mantener es barato y la criticidad es alta"],
    ["AE", "93 %", "Control estricto del capital con servicio alto"],
    ["BE, CE", "95 %, 97 %", "Control intermedio y revisión periódica"],
    ["AD", "85 %", "Control estricto del capital; mínimo stock de seguridad"],
    ["BD, CD", "90 %, 92 %", "Reposición simple"],
], columns=["Celda", "Nivel de servicio objetivo", "Criterio de gestión"]), "Criterio de gestión por celda de la matriz ABC-VED",
    "Los niveles de servicio son supuestos de diseño, sujetos a validación con el servicio farmacéutico, y alimentan el Objetivo 3. Fuente: elaboración propia.", fuente=8.5)

h("8.2.6 Validación de la clasificación y hallazgos", 3)
if c2:
    sens = c2["val_sensibilidad_acuerdo_pct"]; ent = c2["val_acuerdo_entre_anios_pct"]
    p(f"Estabilidad. Al comparar la clasificación con 12 meses y con los 29 meses de datos, {es(c2['val_acuerdo_12m_vs_29m_pct'], 2)} % de los productos conservan su clase "
      f"(correlación de Spearman entre los valores: {es(c2['val_spearman_valor_12m_vs_29m'], 3)}). Entre años consecutivos, el acuerdo es " +
      "; ".join(f"{es(v, 2)} % ({k})" for k, v in ent.items()) + ".")
    p(f"Sensibilidad. Con el costo mínimo activo, {es(sens['costo_minimo_activo'], 2)} % de los productos conservan la clase; con el costo máximo, {es(sens['costo_maximo_activo'], 2)} %. "
      f"Con cortes 70/90 el acuerdo es {es(sens['cortes_70_90'], 2)} % y con 75/95, {es(sens['cortes_75_95'], 2)} %.")
    tr = leer("o2_val_transicion_29m_a_12m.csv", index_col=0)
    if tr is not None:
        tabla(df_es(tr.reset_index().rename(columns={"ABC29": "29 meses \\ 12 meses"})), "Cambio de clase entre la clasificación de 29 meses y la de 12 meses", "Fuente: elaboración propia.")
    cr = leer("o2_val_patron_por_abc.csv", index_col=0)
    if cr is not None:
        tabla(df_es(cr.reset_index().rename(columns={"PATRON": "Patrón de frecuencia"})), "Clase ABC según el patrón de frecuencia del Objetivo 1", "Fuente: elaboración propia.")
    p(f"Casos para revisión clínica. {es(c2['casos_revision_AD_o_CV'])} productos están en clase A con criticidad deseable, o en clase C con criticidad vital; "
      "su lista (o2_val_casos_para_revision.csv) se entrega al servicio farmacéutico para confirmar o ajustar la clasificación.")
    t20 = leer("o2_val_top20_valor.csv", dtype={"CODIGO_PROPIO": str})
    if t20 is not None and mp is not None and vr is not None:
        nE, nV = int((t20["VED"] == "E").sum()), int((t20["VED"] == "V").sum())
        nS = int((t20["PATRON"] == "Consumo sostenido").sum())
        tt = t20.head(10)[["DES_MEDICAMENTO", "VALOR_12M", "PCT_VALOR_12M", "ABC12", "VED"]].copy()
        tt["VALOR_12M"] = tt["VALOR_12M"].map(pesos); tt["PCT_VALOR_12M"] = tt["PCT_VALOR_12M"].map(lambda v: es(v, 2))
        tt.columns = ["Producto", "Valor 12 meses", "% del valor", "ABC", "VED"]
        tabla(tt, "Diez productos de mayor valor de consumo (últimos 12 meses)",
              "La lista de los 20 productos está en o2_val_top20_valor.csv. Fuente: elaboración propia.", fuente=8)
        ved_e = vr.loc["E"]
        p(f"Hallazgos. (1) Los 20 productos de mayor valor concentran el {es(t20['PCT_VALOR_12M'].sum(), 1)} % del valor de consumo y todos son clase A; {nE} son esenciales, {nV} vital y "
          f"{nS} tienen consumo sostenido. Entre ellos hay varios medicamentos oncológicos y biológicos de alto costo, junto con insumos de alto volumen. Como el consumo de estos productos "
          "es continuo y el costo de un exceso o de un faltante es alto, justifican revisión periódica estricta. "
          f"(2) La clasificación VED de la Biblia está concentrada en la clase esencial ({es(ved_e['pct_productos'], 2)} % de los productos con consumo y {es(ved_e['pct_valor'], 2)} % del valor), "
          f"de modo que la matriz discrimina poco dentro de E y casi todo el valor se ubica en la celda AE ({es(mv.loc['A', 'E'], 2)} %). "
          f"(3) Hay {es(int(mp.loc['A', 'D']))} productos en la celda AD (clase A y deseables) y {es(int(mp.loc['C', 'V']))} en la celda CV (clase C y vitales); los de la celda CV son de bajo valor y criticidad alta "
          "(p. ej., tubos endotraqueales y sondas), por lo que la regla de gestión mantiene un nivel de servicio de 99 % sin comprometer capital. "
          f"(4) De {es(c2['productos_total'])} productos con consumo en el período, {es(c2['productos_sin_valor_12m'])} no tienen valor en los últimos 12 meses (sin consumo reciente o sin costo); "
          "la gestión de este grupo se resuelve en el Objetivo 3 con la regla para productos de consumo esporádico.")
else:
    pendiente("validación de la clasificación (script 03)")

h("8.2.7 Cierre del Objetivo 2", 3)
if c2 and a12 is not None and vr is not None:
    A = a12.loc["A"]
    p(f"Se clasificaron los productos del hospital según tres criterios. Impacto económico: con el valor de consumo de los últimos 12 meses ({c2['ventana_abc'][0]} a {c2['ventana_abc'][1]}), "
      f"{es(A['productos'])} productos de clase A concentran el {es(A['pct_valor'], 2)} % del valor. Criticidad clínica: los {es(c2['productos_total'])} productos con consumo se agruparon en vitales, "
      f"esenciales y deseables con la clasificación institucional de la Biblia. Comportamiento del consumo: la clasificación se cruzó con el patrón de frecuencia del Objetivo 1. "
      f"La matriz ABC-VED resultante define el nivel de servicio de cada celda, que alimenta el Objetivo 3.")
    p(f"La clasificación es estable (el {es(c2['val_acuerdo_12m_vs_29m_pct'], 1)} % de los productos conserva su clase al usar 29 meses en lugar de 12) y poco sensible a la elección del costo "
      f"(acuerdo de {es(c2['val_sensibilidad_acuerdo_pct']['costo_maximo_activo'], 1)} % o más). Es más sensible a los cortes de la curva (acuerdo de {es(c2['val_sensibilidad_acuerdo_pct']['cortes_70_90'], 1)} % con cortes 70/90) "
      "y cambia entre años consecutivos, lo que justifica recalcularla periódicamente.")
    p("Limitaciones. El costo es un valor vigente sin historia de precios; la criticidad VED procede de la clasificación institucional, concentrada en la clase esencial, y aún no ha sido validada con "
      f"el químico farmacéutico; y {es(c2['casos_revision_AD_o_CV'])} productos se entregan como casos para revisión clínica. Con estas salvedades, el Objetivo 2 se considera cumplido: los productos "
      "quedaron clasificados por impacto económico, criticidad clínica y comportamiento del consumo, con evidencia de validación y trazabilidad reproducible en el repositorio.")
else:
    pendiente("redactar el cumplimiento del Objetivo 2 con los resultados de las secciones 8.2.3 a 8.2.6")

# ================================================================ PARTE C · OBJETIVO 3
h("8.3 Desarrollo del Objetivo Específico 3", 2)
p("Estimar niveles de inventario mínimo, máximo y de reposición para los medicamentos y dispositivos médicos, a partir del comportamiento histórico del consumo y la variabilidad de la demanda.",
  "Objetivo específico 3. ")
c3 = cifras("o3_cifras_clave.json")
h("8.3.1 Diseño metodológico", 3)
p("El modelo estima, para cada producto y servicio farmacéutico, tres niveles. Con μ y σ como la media y la desviación estándar del consumo mensual en la ventana de estimación, "
  "L como tiempo de reposición y R como período de revisión (ambos en meses), H = L + R, y z el cuantil normal del nivel de servicio objetivo:")
bullets(["Punto de reposición (ROP) = μ·L + z·σ·√L",
         "Máximo (S) = μ·H + z·σ·√H",
         "Mínimo = stock de seguridad = ROP − μ·L"])
p("Esta formulación corresponde a un sistema de revisión periódica con nivel objetivo (Silver et al., 1998) y precisa las fórmulas simplificadas del marco teórico. "
  "El método de cálculo depende del patrón de frecuencia del Objetivo 1, porque el supuesto de normalidad no es adecuado cuando el consumo es intermitente (Syntetos & Boylan, 2005):")
tabla(pd.DataFrame([
    ["Sostenido, frecuente", "Modelo normal (fórmulas anteriores)", "Demanda continua"],
    ["Intermitente", "Cuantil empírico del consumo mensual al nivel de servicio, escalado por H y por L", "Muchos meses en cero"],
    ["Esporádico, vital", "Máximo consumo mensual de la ventana", "No puede faltar aunque rote poco"],
    ["Esporádico, esencial o deseable", "Bajo demanda (sin niveles)", "Baja participación en el consumo (menos del 1 %)"],
], columns=["Patrón", "Método", "Razón"]), "Método de estimación según el patrón de frecuencia",
    "Fuente: elaboración propia.", fuente=8.5)
p("Supuestos: tiempo de reposición L = 0,5 mes y período de revisión R = 0,5 mes, porque los datos disponibles no incluyen tiempos de entrega; el efecto de este supuesto se evalúa "
  "en el análisis de sensibilidad (L = 0,25 y 0,75). La ventana de estimación es de 12 meses para reflejar la demanda reciente. El nivel de servicio depende de la celda ABC-VED (Tabla de la Sección 8.2.5).")
figura("flujo_03_niveles_inventario.png", "Flujo de estimación de niveles de inventario", "Fuente: elaboración propia.", ancho=11)

h("8.3.2 Evaluación retrospectiva", 3)
bm = leer("o3_backtest_metricas.csv")
if bm is not None:
    p(f"Se evaluaron {es(c3.get('origenes_evaluados'))} meses de origen sobre {es(c3.get('series_backtest'))} productos, con el protocolo de la Sección 7.9.")
    orden = ["PROPUESTA", "B1_ultimo_mes", "B2_promedio_historico", "A1_promedio_ventana", "B2_capital_equivalente"]
    nombres = {"PROPUESTA": "Propuesta", "B1_ultimo_mes": "B1 último mes", "B2_promedio_historico": "B2 promedio histórico",
               "A1_promedio_ventana": "A1 promedio 12 meses", "B2_capital_equivalente": "B2k capital equivalente"}
    g = bm[bm["VED"] == "TODAS"].set_index("politica").reindex(orden)
    t = pd.DataFrame({"Política": [nombres[i] for i in g.index], "Fill rate unidades (%)": g["fill_rate_unidades_%"].round(2).values,
                      "Fill rate valor (%)": g["fill_rate_valor_%"].round(2).values,
                      "Disponibilidad ante demanda (%)": g["disponibilidad_ante_demanda_%"].round(2).values,
                      "Inventario promedio (millones COP)": (g["inventario_prom_valor_mensual"] / 1e6).round(1).values})
    tabla(df_es(t, {c: 2 for c in t.columns[1:4]} | {t.columns[4]: 1}), "Resultados de la evaluación retrospectiva (todos los productos)",
          "Fill rate y disponibilidad calculados sobre el consumo real del mes siguiente. Inventario promedio teórico mensual. Fuente: elaboración propia.", fuente=8.5)
    gv = bm[bm["VED"] != "TODAS"].pivot(index="politica", columns="VED", values="disponibilidad_ante_demanda_%").reindex(orden)
    t = gv.round(2).reset_index(); t["politica"] = t["politica"].map(nombres); t.columns = ["Política", "D (%)", "E (%)", "V (%)"]
    tabla(df_es(t[["Política", "V (%)", "E (%)", "D (%)"]], {"V (%)": 2, "E (%)": 2, "D (%)": 2}), "Disponibilidad ante demanda por criticidad", "Fuente: elaboración propia.", fuente=8.5)
    ic, pr = leer("o3_backtest_ic95.csv"), leer("o3_backtest_pruebas_pareadas.csv")
    if ic is not None:
        t = ic[["politica", "fill_rate_valor_%", "IC95_inf", "IC95_sup"]].copy(); t["politica"] = t["politica"].map(nombres)
        t.columns = ["Política", "Fill rate valor (%)", "IC 95 % inferior", "IC 95 % superior"]
        tabla(df_es(t, {c: 2 for c in t.columns[1:]}), "Fill rate en valor con intervalo de confianza del 95 %", "Bootstrap sobre los meses de origen; corresponde al promedio de los fill rate mensuales, por lo que difiere levemente del cociente de sumas de la tabla anterior. Fuente: elaboración propia.", fuente=8.5)
    if pr is not None:
        t = pr[["comparacion", "meses_propuesta_mayor_fill", "p_wilcoxon_fill", "meses_evaluados"]].copy()
        t.columns = ["Comparación", "Meses con mayor fill rate", "p (Wilcoxon)", "Meses evaluados"]
        t["p (Wilcoxon)"] = t["p (Wilcoxon)"].map(lambda v: "< 0,0001" if v < 1e-4 else es(v, 4))
        tabla(df_es(t), "Comparación pareada de la propuesta con las líneas base",
              "Los 17 meses de origen no son independientes (las ventanas de estimación se superponen), por lo que el valor p debe leerse como indicativo. Fuente: elaboración propia.", fuente=8.5)
    figura("o3_servicio_vs_capital.png", "Nivel de servicio frente a capital inmovilizado", "Fuente: elaboración propia.", ancho=10)
    cr = leer("o3_criterios_exito.csv")
    if cr is not None:
        cr["Cumple"] = cr["Cumple"].map({True: "Sí", False: "No"})
        tabla(cr, "Verificación de los criterios de éxito", "Criterios definidos en el protocolo de evaluación (Sección 7.9). Fuente: elaboración propia.", fuente=8.5)
        cap = c3.get("capital_propuesta_vs_B2")
        if cap:
            p(f"La propuesta mantiene un inventario promedio teórico {es(cap, 2)} veces el de la línea base B2. Este capital adicional se tratará en el Objetivo 4 con escenarios de restricción presupuestal.")
else:
    pendiente("resultados de la evaluación: ejecutar scripts/04_objetivo3_niveles_inventario.py")

h("8.3.3 Análisis de sensibilidad", 3)
se = leer("o3_sensibilidad.csv")
if se is not None:
    tt = se[se["VED"] == "TODAS"][["escenario", "fill_rate_valor_%", "disponibilidad_%", "inventario_prom_millones"]].copy()
    tt.columns = ["Escenario", "Fill rate valor (%)", "Disponibilidad (%)", "Inventario promedio (millones COP)"]
    tabla(df_es(tt, {"Fill rate valor (%)": 2, "Disponibilidad (%)": 2, "Inventario promedio (millones COP)": 1}), "Sensibilidad de la propuesta a la ventana, el tiempo de reposición y el nivel de servicio", "Fuente: elaboración propia.", fuente=8.5)
else:
    pendiente("sensibilidad (script 04)")

h("8.3.4 Niveles finales por servicio farmacéutico", 3)
rp = leer("o3_resumen_por_patron.csv", index_col=0)
if rp is not None:
    t = rp.reset_index(); t.columns = ["Patrón", "Series producto-servicio", "Con nivel definido", "Suma de máximos (unidades)"]
    tabla(df_es(t), "Niveles de inventario estimados por patrón de frecuencia", "Se aplican las reglas vigentes de Novedades. La tabla completa está en niveles_inventario_por_producto_servicio.csv. Fuente: elaboración propia.", fuente=8.5)
ci = leer("o3_comparacion_inventario_actual.csv", index_col=0)
if ci is not None:
    ci = ci["valor"]
    p(f"Comparación indicativa con el inventario actual: de {es(float(ci['productos_comparados']))} productos comparados, {es(float(ci['productos_con_stock_sobre_maximo']))} ({es(float(ci['pct_productos_sobre_maximo']), 2)} %) "
      f"tienen existencias superiores al máximo propuesto, con un valor de exceso de {pesos(float(ci['valor_exceso_cop']))}, y {es(float(ci['productos_bajo_punto_reposicion']))} "
      f"({es(float(ci['pct_productos_bajo_punto_reposicion']), 2)} %) están por debajo del punto de reposición. La fecha de corte de la disponibilidad no está registrada, por lo que esta comparación es referencial.")
h("8.3.5 Cierre del Objetivo 3", 3)
cr3 = leer("o3_criterios_exito.csv")
dv = leer("o3_diagnostico_vitales.csv", index_col=0)
if cr3 is not None and c3:
    ok = cr3.set_index("Criterio")["Cumple"].astype(bool)
    res_ = cr3.set_index("Criterio")["Resultado"]; umb_ = cr3.set_index("Criterio")["Umbral"].astype(str)
    p(f"Criterios de éxito. Se cumplieron {' y '.join(k for k in ok.index if ok[k])}"
      + (f" y no se cumplió {' ni '.join(k for k in ok.index if not ok[k])}" if (~ok).any() else "") +
      f". La propuesta atiende el {es(float(res_['C2']), 2)} % del valor demandado frente a {es(float(umb_['C2'].replace('>', '')), 2)} % de la mejor línea base "
      f"(B2), y supera a B2 escalada a igual capital ({es(float(umb_['C3'].replace('>', '')), 2)} %). La disponibilidad de los productos vitales es de {es(float(res_['C1']), 2)} %, "
      "por debajo del umbral de 95 % definido a priori. Siguiendo el protocolo, el umbral no se modificó después de ver el resultado y se analizó la causa.")
    if dv is not None:
        t = dv[["series_con_demanda", "faltantes", "disponibilidad_%", "%_de_los_faltantes_vitales", "faltantes_con_nivel_cero"]].reset_index()
        t.columns = ["Patrón", "Productos-mes con demanda", "Faltantes", "Disponibilidad (%)", "% de los faltantes", "Faltantes con nivel cero"]
        tabla(df_es(t, {"Disponibilidad (%)": 2, "% de los faltantes": 2}), "Disponibilidad de los productos vitales por patrón de frecuencia",
              "Suma de los 17 meses de origen. Faltante: demanda del mes siguiente mayor que el nivel máximo. Fuente: elaboración propia.", fuente=8.5)
        sos, esp = dv.loc["Consumo sostenido"], dv.loc["Consumo esporádico"]
        sens = leer("o3_sensibilidad.csv")
        sv = sens[(sens["VED"] == "V")].set_index("escenario")["disponibilidad_%"] if sens is not None else None
        p(f"Causa. Los vitales de consumo sostenido alcanzan {es(sos['disponibilidad_%'], 2)} %, pero los vitales esporádicos solo {es(esp['disponibilidad_%'], 2)} % y aportan el "
          f"{es(esp['%_de_los_faltantes_vitales'], 1)} % de los faltantes vitales; en {es(int(esp['faltantes_con_nivel_cero']))} de sus {es(int(esp['faltantes']))} faltantes el nivel era cero porque el producto "
          "no tuvo consumo en los 12 meses previos. Son demandas que el historial no permite anticipar con un modelo estadístico, de modo que la regla de máximo mensual de la ventana es insuficiente para ellos.")
        if sv is not None:
            p(f"Alternativas evaluadas. Ampliar la ventana a 24 meses sube la disponibilidad de los vitales a {es(sv['ventana=24'], 2)} % y subir tres puntos el nivel de servicio, a {es(sv['Δ nivel servicio=+0.03'], 2)} %, "
              "pero ninguna alternativa alcanza 95 % y ambas aumentan el capital inmovilizado. La cobertura de los vitales esporádicos requeriría una regla de negocio (existencia mínima institucional por "
              "criticidad) que no se deriva del historial; se propone evaluarla en el Objetivo 4 junto con el costo de mantenerla. El tiempo de reposición no modifica el nivel máximo porque H = L + R se "
              "mantiene en un mes; solo desplaza el punto de reposición y el mínimo.")
    p(f"Conclusión. Los niveles mínimo, de reposición y máximo se estimaron para {es(c3['series_con_nivel'])} de {es(c3['series_finales'])} combinaciones producto-servicio (el resto corresponde a consumo "
      "esporádico no vital o sin consumo reciente, que se gestionan bajo demanda) y se evaluaron con un protocolo definido de antemano. La propuesta supera a las líneas base en servicio, incluso a igual capital, pero "
      f"requiere {es(c3['capital_propuesta_vs_B2'], 2)} veces el inventario promedio de B2. Como el criterio C1 no se cumple, el Objetivo 3 se considera cumplido en la estimación y la evaluación de los niveles, "
      "y parcialmente cumplido en la meta de disponibilidad de vitales; las acciones para cerrar esa brecha se retoman en el Objetivo 4. Las decisiones pendientes de validación con el servicio farmacéutico "
      "(niveles de servicio, tiempo de reposición, forma de compra actual) figuran en la bitácora (Anexo B).")
else:
    pendiente("redactar el cumplimiento del Objetivo 3 a partir de la verificación de los criterios C1 a C3")

# ================================================================ ANEXOS
doc.add_page_break()
h("Anexo B. Bitácora de decisiones", 1)
ruta_b = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "docs", "bitacora_decisiones.md")
if os.path.exists(ruta_b):
    filas = [[c.strip() for c in l.strip().strip("|").split("|")] for l in open(ruta_b, encoding="utf-8")
             if l.startswith("| D-")]
    tabla(pd.DataFrame(filas, columns=["ID", "Fase", "Decisión", "Alternativa considerada", "Justificación / evidencia", "Estado"]),
          "Bitácora de decisiones de depuración y modelado", "Fuente: elaboración propia.", fuente=7)

h("Referencias nuevas y corregidas (APA 7)", 1)
p("Reemplazan o se agregan a la lista del ACA 1. Orden alfabético y sangría francesa. Verificar cada DOI antes de entregar.")
refs = [
    "Hyndman, R. J., & Athanasopoulos, G. (2021). Forecasting: Principles and practice (3.ª ed.). OTexts. https://otexts.com/fpp3/",
    "Kabera, J. C., & Mukanyangezi, M. F. (2024). Influence of inventory management practices on the availability of emergency obstetric drugs in Rwandan public hospitals: A case of Rwanda Southern Province. BMC Health Services Research, 24, Artículo 14. https://doi.org/10.1186/s12913-023-10459-x",
    "Rehman, A. U., Mian, S. H., Usmani, Y. S., Abidi, M. H., & Mohammed, M. K. (2023). Modelling and analysis of hospital inventory policies during COVID-19 pandemic. Processes, 11(4), 1062. https://doi.org/10.3390/pr11041062",
    "Syntetos, A. A., & Boylan, J. E. (2005). The accuracy of intermittent demand estimates. International Journal of Forecasting, 21(2), 303–314. https://doi.org/10.1016/j.ijforecast.2004.10.001",
]
for r in refs:
    par = doc.add_paragraph(r)
    par.paragraph_format.left_indent = Cm(1.27); par.paragraph_format.first_line_indent = Cm(-1.27)

doc.save(SALIDA)
print("Documento generado:", SALIDA)
