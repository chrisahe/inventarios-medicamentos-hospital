# %% [markdown]
# # 00c · Diagramas de flujo por paso del Objetivo 1
# Reemplazan, una a una, las capturas de código y de resultados de Colab del ACA 1 (Figuras 1-24, 27, 28, 30-34 y 36-40).
# Cada diagrama resume el procedimiento del código correspondiente (disponible en el repositorio).
# Salida: `outputs/figuras/pasos/fig_NN.png`. El documento (06_integrar_documento.py) lee `PASOS` para sustituir las imágenes.

# %%
import os, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

FIG = os.path.join(os.environ.get("OUT_FIG", "outputs/figuras"), "pasos")
os.makedirs(FIG, exist_ok=True)
COLORES = {"e": "#dbe8f5", "p": "#eef3e2", "d": "#fbe9c8", "s": "#e3d9f0"}   # entrada, proceso, decisión, salida


def flujo(nombre, nodos, ancho=5.4, sep=0.3, chars=48):
    lineas = [textwrap.wrap(t, chars) for _, t in nodos]
    altos = [0.3 + 0.19 * len(l) for l in lineas]
    alto = sum(altos) + sep * (len(nodos) - 1) + 0.3
    fig, ax = plt.subplots(figsize=(ancho, alto))
    ax.set_xlim(0, 10); ax.set_ylim(0, alto); ax.axis("off")
    y_top, ys = alto - 0.15, []
    for (tipo, _), l, h in zip(nodos, lineas, altos):
        y = y_top - h / 2
        ax.add_patch(FancyBboxPatch((0.9, y - h / 2), 8.2, h, boxstyle="round,pad=0.02,rounding_size=0.12",
                                    fc=COLORES[tipo], ec="#555555", lw=1))
        ax.text(5, y, chr(10).join(l), ha="center", va="center", fontsize=8.6, linespacing=1.25)
        ys.append((y, h)); y_top -= h + sep
    for (y1, h1), (y2, h2) in zip(ys[:-1], ys[1:]):
        ax.add_patch(FancyArrowPatch((5, y1 - h1 / 2), (5, y2 + h2 / 2), arrowstyle="-|>", mutation_scale=11, color="#555555"))
    fig.tight_layout(pad=0.2); fig.savefig(f"{FIG}/{nombre}.png", dpi=170); plt.close(fig)


# Figura -> (título del rótulo, nodos). Los números que aparecen provienen del propio texto del ACA 1.
PASOS = {
    1: ("Diagrama de flujo de la carga de las bases de consumo de 2024",
        [("e", "Archivo comprimido 2024.zip (12 archivos mensuales .xlsx)"), ("p", "Cargar el archivo en Google Colab (files.upload)"),
         ("s", "Archivo disponible en el entorno de trabajo")]),
    2: ("Diagrama de flujo de la identificación de los archivos mensuales de 2024",
        [("e", "Archivo 2024.zip"), ("p", "Leer el contenido del ZIP (zipfile)"), ("p", "Filtrar los nombres que terminan en .xlsx"),
         ("s", "Doce archivos mensuales identificados (enero a diciembre)")]),
    3: ("Diagrama de flujo de la verificación de la estructura de las bases mensuales de 2024",
        [("e", "Doce archivos mensuales de 2024"), ("p", "Leer las primeras filas de cada archivo"),
         ("p", "Registrar el número y los nombres de las columnas"), ("s", "Tabla de estructura por archivo")]),
    4: ("Diagrama de flujo del conteo de registros de las bases mensuales de 2024",
        [("e", "Doce archivos mensuales de 2024"), ("p", "Leer cada archivo completo"), ("p", "Contar las filas de cada archivo"),
         ("s", "Tabla de registros por archivo")]),
    5: ("Diagrama de flujo de la organización cronológica y el volumen de registros de 2024",
        [("e", "Tabla de registros por archivo"), ("p", "Identificar el mes por el nombre del archivo"),
         ("p", "Asignar el número de mes y ordenar cronológicamente"), ("s", "Total de registros de 2024")]),
    6: ("Diagrama de flujo de la validación de calidad de las bases de consumo de 2024",
        [("e", "Doce archivos mensuales de 2024"), ("p", "Normalizar nombres de columnas, mes y cantidad numérica"),
         ("p", "Contar nulos en código, medicamento, cantidad, bodega, unidad funcional, mes y año"),
         ("p", "Contar cantidades iguales a cero o negativas y años o meses fuera de rango"), ("s", "Tabla de hallazgos por archivo")]),
    7: ("Diagrama de flujo de la consolidación de las variables de caracterización de 2024",
        [("e", "Doce archivos mensuales de 2024"), ("p", "Leer cada archivo"),
         ("p", "Conservar CODIGO_PROPIO, DES_MEDICAMENTO, BODEGA y DES_UF"), ("p", "Concatenar los doce meses"),
         ("s", "Universo de consumo de 2024")]),
    8: ("Diagrama de flujo de la caracterización global del universo de consumo de 2024",
        [("e", "Universo de consumo de 2024"), ("p", "Contar valores únicos de código, descripción, bodega y unidad funcional"),
         ("s", "Indicadores globales del universo")]),
    9: ("Diagrama de flujo de la identificación de productos asociados a más de una bodega",
        [("e", "Universo de consumo de 2024"), ("p", "Agrupar por código de producto"), ("p", "Contar las bodegas distintas de cada producto"),
         ("d", "¿Más de una bodega?"), ("s", "Número de productos presentes en varias bodegas")]),
    10: ("Diagrama de flujo de la consolidación de la base histórica de consumo",
         [("e", "Archivos ZIP de 2024, 2025 y enero–mayo de 2026"), ("p", "Leer cada archivo mensual y normalizar los nombres de columnas"),
          ("p", "Conservar las siete variables definidas para el análisis"), ("p", "Unir todos los meses en una sola base"),
          ("s", "Base histórica de consumo")]),
    11: ("Diagrama de flujo del resultado de la consolidación de la base histórica",
         [("e", "Bases mensuales de 2024, 2025 y enero–mayo de 2026"), ("p", "Integrar en una sola base (concatenación)"),
          ("s", "Total de registros de la base histórica")]),
    12: ("Diagrama de flujo de la verificación de los períodos mensuales de la base histórica",
         [("e", "Base histórica de consumo"), ("p", "Agrupar por año y mes"), ("p", "Contar los registros de cada período"),
          ("d", "¿Cobertura completa: 12 meses en 2024, 12 en 2025 y 5 en 2026?"), ("s", "Cobertura temporal verificada")]),
    13: ("Diagrama de flujo de la creación de la variable temporal PERIODO",
         [("e", "Base histórica (año y mes)"), ("p", "Convertir el nombre del mes en número"),
          ("p", "Construir la fecha del primer día de cada mes"), ("s", "Variable PERIODO (tipo fecha)")]),
    14: ("Diagrama de flujo de la identificación de cantidades nulas o negativas",
         [("e", "Base histórica consolidada"), ("p", "Filtrar los registros con CANTIDAD menor o igual a cero"),
          ("p", "Listar producto, bodega, unidad funcional, mes y año"), ("s", "Registros marcados para revisión")]),
    15: ("Diagrama de flujo de la lectura de la base maestra de Bodegas",
         [("e", "Archivo bodegas.xlsx"), ("p", "Leer el archivo con pandas"),
          ("p", "Revisar las columnas: Bodega, Servicio, Servicio Farmacéutico, Sede y Se utiliza"), ("s", "Base maestra de 37 bodegas")]),
    16: ("Diagrama de flujo de la revisión de los productos de la bodega 829",
         [("e", "Base histórica consolidada"), ("p", "Filtrar los registros de la bodega 829"), ("p", "Listar los productos distintos"),
          ("p", "Contrastar con la condición especial de la base de Bodegas"), ("s", "Productos de la bodega 829 identificados")]),
    17: ("Diagrama de flujo de la identificación de productos cuya descripción inicia con “NPT”",
         [("e", "Base histórica consolidada"), ("p", "Normalizar la descripción (texto, sin espacios, mayúsculas)"),
          ("d", "¿La descripción inicia con NPT?"), ("s", "Registros de productos NPT identificados")]),
    18: ("Diagrama de flujo de la validación temporal de la base de Novedades",
         [("e", "Base de Novedades (fecha de reporte)"), ("p", "Filtrar fechas entre el 1 de enero de 2024 y el 31 de mayo de 2026"),
          ("d", "¿Hay registros dentro del período?"), ("s", "Ninguno: Novedades no se usa como variable temporal")]),
    19: ("Diagrama de flujo de la identificación de los productos permitidos en la bodega 829",
         [("e", "Base histórica consolidada"), ("p", "Normalizar la descripción del producto"),
          ("p", "Marcar bodegas con Se utiliza = No, productos NPT y registros de la bodega 829"),
          ("d", "¿Bolsa EVA de nutrición parenteral de 500, 1000 o 3000 ml?"), ("s", "Criterio de exclusión definido")]),
    20: ("Diagrama de flujo de la depuración definitiva de la base histórica de consumo",
         [("e", "Base histórica: 11.625.461 registros"),
          ("p", "Excluir bodegas con Se utiliza = No, productos NPT y productos no permitidos de la bodega 829"),
          ("s", "Base depurada: 11.384.908 registros (240.553 excluidos)")]),
    21: ("Diagrama de flujo de la validación de cantidades no válidas",
         [("e", "Base depurada"), ("p", "Contar los registros con cantidad negativa o igual a cero"),
          ("s", "68 registros negativos y 1 registro igual a cero")]),
    22: ("Diagrama de flujo de la identificación del oxígeno medicinal",
         [("e", "Base depurada"), ("p", "Filtrar el código 100020287 (oxígeno medicinal)"),
          ("p", "Sumar los registros y la cantidad por año"), ("s", "Oxígeno medicinal identificado (gestión de Biomédica)")]),
    23: ("Diagrama de flujo de la verificación de la exclusión del oxígeno medicinal",
         [("e", "Base depurada"), ("p", "Excluir el código 100020287"), ("d", "¿Quedan registros del código?"),
          ("s", "Ninguno: base analítica final de 11.341.686 registros")]),
    24: ("Diagrama de flujo de la consolidación del consumo mensual de la base analítica final",
         [("e", "Base analítica final"), ("p", "Agrupar por PERIODO"), ("p", "Sumar la cantidad consumida"),
          ("p", "Ordenar cronológicamente"), ("s", "Consumo mensual de los 29 períodos")]),
    27: ("Diagrama de flujo del cálculo del consumo por Servicio Farmacéutico",
         [("e", "Base analítica final"), ("p", "Agrupar por SERVICIO_FARMACEUTICO"), ("p", "Sumar la cantidad consumida"),
          ("p", "Ordenar de mayor a menor"), ("s", "Consumo total por servicio")]),
    28: ("Diagrama de flujo de la validación de la distribución por Servicio Farmacéutico",
         [("e", "Base analítica final y consumo agrupado por servicio"), ("p", "Comparar el número de registros"),
          ("p", "Comparar la cantidad total consumida"), ("d", "¿Diferencia igual a cero?"), ("s", "Agrupación validada")]),
    30: ("Diagrama de flujo de la conciliación de registros y cantidades por Servicio Farmacéutico",
         [("e", "Resumen por servicio y base analítica final"), ("p", "Sumar los registros y las cantidades agrupadas"),
          ("p", "Restar los totales de la base analítica"), ("s", "Diferencia de cero en registros y en consumo")]),
    31: ("Diagrama de flujo de la identificación del universo de productos y su consumo acumulado",
         [("e", "Base analítica final"), ("p", "Agrupar por código y descripción del producto"),
          ("p", "Calcular el número de registros y la cantidad total"), ("p", "Ordenar por consumo total"),
          ("s", "Universo de productos analizados")]),
    32: ("Diagrama de flujo de la frecuencia mensual de consumo por producto",
         [("e", "Base analítica final"), ("p", "Contar los períodos mensuales distintos con consumo de cada producto"),
          ("p", "Ordenar por meses con consumo"), ("s", "Meses con consumo por producto (1 a 29)")]),
    33: ("Diagrama de flujo de la distribución de productos según la frecuencia histórica de consumo",
         [("e", "Meses con consumo por producto"), ("p", "Contar los productos por número de meses con consumo"),
          ("p", "Calcular el porcentaje de productos"), ("s", "Distribución de la frecuencia de consumo")]),
    34: ("Diagrama de flujo del cálculo de la intensidad mensual del consumo por producto",
         [("e", "Consumo total y meses con consumo por producto"), ("p", "Dividir el consumo total entre los meses con consumo"),
          ("s", "Consumo promedio por mes activo")]),
    36: ("Diagrama de flujo de la identificación de los productos de consumo sostenido con mayor consumo",
         [("e", "Productos con patrón de frecuencia"), ("p", "Filtrar el patrón de consumo sostenido"),
          ("p", "Ordenar por consumo total acumulado"), ("s", "Veinte productos con mayor consumo")]),
    37: ("Diagrama de flujo de la consolidación del consumo mensual por producto",
         [("e", "Base analítica final"), ("p", "Agrupar por código, descripción y período mensual"), ("p", "Sumar la cantidad consumida"),
          ("s", "Tabla producto-mes de 68.888 registros")]),
    38: ("Diagrama de flujo de la construcción de la matriz de consumo mensual por producto",
         [("e", "Tabla producto-mes"), ("p", "Colocar los productos en filas y los períodos en columnas"),
          ("p", "Completar con cero los meses sin consumo"), ("s", "Matriz de 11.887 productos por 29 períodos")]),
    39: ("Diagrama de flujo del cálculo de los indicadores de variabilidad del consumo",
         [("e", "Matriz de consumo mensual por producto"), ("p", "Calcular el promedio mensual"),
          ("p", "Calcular la desviación estándar"), ("p", "Calcular el coeficiente de variación = desviación estándar / promedio"),
          ("s", "Indicadores de variabilidad por producto")]),
    40: ("Diagrama de flujo del cálculo de los estadísticos descriptivos del coeficiente de variación",
         [("e", "Indicadores de variabilidad por producto (11.887 productos)"),
          ("p", "Calcular mínimo, percentiles 25, 50, 75, 90 y 95 y máximo del coeficiente de variación"),
          ("s", "Tabla de estadísticos descriptivos del coeficiente de variación")]),
}

# %%
if __name__ == "__main__":
    for n, (_, nodos) in PASOS.items():
        flujo(f"fig_{n:02d}", nodos)
    print(len(PASOS), "diagramas en", FIG)
