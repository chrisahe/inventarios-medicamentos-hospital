# %% [markdown]
# # 00 · Diagramas de flujo
# Reemplazan las capturas de código del ACA 1 (retroalimentación: "Utiliza diagramas de flujo").
# Genera una figura por proceso en `outputs/figuras/`.

# %%
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

FIG = os.environ.get("OUT_FIG", "outputs/figuras")
os.makedirs(FIG, exist_ok=True)
COLORES = {"entrada": "#dbe8f5", "proceso": "#eef3e2", "decision": "#fbe9c8", "salida": "#e3d9f0"}


def flujo(nombre, titulo, nodos, ancho=7.5, sep=0.38, chars=68):
    """nodos: lista de (texto, tipo). Se dibujan en columna, con el texto ajustado
    al ancho de la caja, y se unen en orden con flechas."""
    import textwrap
    lineas = [textwrap.wrap(t, chars) for t, _ in nodos]
    altos = [0.32 + 0.2 * len(l) for l in lineas]
    alto = sum(altos) + sep * (len(nodos) - 1) + 1.1
    fig, ax = plt.subplots(figsize=(ancho, alto))
    ax.set_xlim(0, 10); ax.set_ylim(0, alto); ax.axis("off")
    ax.text(5, alto - 0.35, titulo, ha="center", va="center", fontsize=11, fontweight="bold")
    y_top, ys = alto - 0.8, []
    for (texto, tipo), l, h in zip(nodos, lineas, altos):
        y = y_top - h / 2
        ax.add_patch(FancyBboxPatch((0.9, y - h / 2), 8.2, h, boxstyle="round,pad=0.02,rounding_size=0.12",
                                    fc=COLORES[tipo], ec="#555555", lw=1))
        ax.text(5, y, chr(10).join(l), ha="center", va="center", fontsize=8.6, linespacing=1.25)
        ys.append((y, h)); y_top -= h + sep
    for (y1, h1), (y2, h2) in zip(ys[:-1], ys[1:]):
        ax.add_patch(FancyArrowPatch((5, y1 - h1 / 2), (5, y2 + h2 / 2), arrowstyle="-|>",
                                     mutation_scale=11, color="#555555"))
    fig.tight_layout(); fig.savefig(f"{FIG}/{nombre}.png", dpi=150); plt.close(fig)


# %% [markdown]
# ## Objetivo 1: extracción, depuración y base analítica (reemplaza Figuras 1-23)

# %%
flujo("flujo_01_etl_objetivo1", "Flujo ETL de la base analítica de consumo (Objetivo 1)", [
    ("Archivos mensuales .xlsx 2024, 2025 y ene-may 2026 (29 archivos, 7 variables)", "entrada"),
    ("Verificar estructura y contar registros por archivo (zipfile / pandas)", "proceso"),
    ("Validar calidad: nulos, cantidades ≤ 0, consistencia de año y mes", "proceso"),
    ("Consolidar en una base histórica: 11.625.461 registros + variable PERIODO", "proceso"),
    ("Excluir bodegas 'Se utiliza = No', productos NPT y regla de la bodega 829: −240.553", "proceso"),
    ("Excluir 69 cantidades no válidas (68 negativas y 1 cero)", "proceso"),
    ("Excluir oxígeno medicinal (código 100020287, gestión de Biomédica): −43.153", "proceso"),
    ("Base analítica final: 11.341.686 registros", "salida"),
    ("Asignar SERVICIO_FARMACEUTICO con la base maestra de Bodegas", "proceso"),
])

# %% [markdown]
# ## Objetivo 2: costos, ABC, VED y matriz

# %%
flujo("flujo_02_clasificacion_abc_ved", "Flujo de clasificación ABC-VED (Objetivo 2)", [
    ("Base analítica agregada: producto × bodega × servicio × período", "entrada"),
    ("Asignar costo unitario (Biblia): bodega activa → mediana activa → mediana inactiva → sin costo", "proceso"),
    ("Valor de consumo = cantidad × costo unitario, últimos 12 meses", "proceso"),
    ("ABC: orden descendente por valor; A hasta 80 %, B hasta 95 %, C el resto", "proceso"),
    ("VED: VEN de la Biblia (V, E, N→D); sin dato → E; 'Siempre deben haber existencias' → V", "proceso"),
    ("Matriz ABC-VED (3 × 3): productos y valor por celda", "proceso"),
    ("Validación: estabilidad entre ventanas, sensibilidad al costo y a los cortes, cruce con patrones", "decision"),
    ("Clasificación por producto → insumo del Objetivo 3", "salida"),
])

# %% [markdown]
# ## Objetivo 3: niveles de inventario y evaluación

# %%
flujo("flujo_03_niveles_inventario", "Flujo de estimación de niveles de inventario (Objetivo 3)", [
    ("Matriz de demanda mensual por producto × servicio (ceros en meses sin consumo)", "entrada"),
    ("Ventana de estimación: últimos 12 meses; media, desviación y fracción de meses activos", "proceso"),
    ("Patrón de frecuencia: sostenido, frecuente, intermitente, esporádico", "decision"),
    ("Nivel de servicio objetivo según la celda ABC-VED", "proceso"),
    ("Sostenido/frecuente: modelo normal · Intermitente: cuantil empírico · Esporádico: V con existencia, resto bajo demanda", "proceso"),
    ("Mínimo (stock de seguridad), punto de reposición y máximo", "proceso"),
    ("Reglas de Novedades vigentes (descontinuado, a necesidad, existencia permanente, Invima vencido)", "proceso"),
    ("Niveles por producto y servicio farmacéutico", "salida"),
])

flujo("flujo_04_evaluacion", "Protocolo de evaluación retrospectiva (origen móvil)", [
    ("Para cada mes de origen t (desde el mes 12): usar solo datos hasta t", "entrada"),
    ("Recalcular ABC, patrón y niveles con la ventana que termina en t (sin fuga de información)", "proceso"),
    ("Comparar con el consumo real del mes t+1 para la propuesta y las líneas base B1, B2, A1", "proceso"),
    ("Métricas: fill rate, disponibilidad ante demanda, valor de inventario promedio teórico", "proceso"),
    ("Capital equivalente: reescalar B2 para igualar el capital de la propuesta", "proceso"),
    ("IC 95 % por bootstrap sobre meses y prueba de Wilcoxon pareada", "proceso"),
    ("Verificar criterios C1-C3 y reportar sensibilidad (ventana, L, nivel de servicio)", "decision"),
])

flujo("flujo_05_arquitectura", "Arquitectura de la solución", [
    ("Fuentes institucionales: consumo mensual, Biblia, Bodegas, Disponibilidad, Novedades", "entrada"),
    ("Anonimización y depuración manual previa (sin datos de pacientes)", "proceso"),
    ("Python (Google Colab): ETL, clasificación, niveles y evaluación — repositorio en GitHub", "proceso"),
    ("Tablas de resultados (CSV) · SQL Server para almacenamiento", "proceso"),
    ("Power BI: tablero de clasificación, niveles y escenarios presupuestales (Objetivo 4)", "salida"),
])
print("Diagramas generados en", FIG)
