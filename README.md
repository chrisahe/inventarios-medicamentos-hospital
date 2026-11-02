# Modelo analítico para la optimización de inventarios de medicamentos y dispositivos médicos

Proyecto de grado — Especialización en Analítica de Datos (CUN).
Autores: Angy Daiana Carmona Vasco y Christian Camilo Saldarriaga Hernández · Director: Andrés Felipe Solis Pino.

Código reproducible de los objetivos específicos 1 a 3: caracterización del consumo, clasificación ABC-VED y
niveles de inventario (mínimo, punto de reposición y máximo) para un hospital de alta complejidad,
con datos de enero de 2024 a mayo de 2026.

## Estructura

```
src/inventario.py                       funciones (costos, ABC, VED, políticas, evaluación)
scripts/00_diagramas_de_flujo.py        diagramas de flujo de los procesos
scripts/00b_construir_base_analitica.py depuración y base analítica a partir de los 29 archivos mensuales
scripts/01_diccionario_y_ficha.py       diccionario de datos y ficha técnica
scripts/02_costos_y_linea_base.py       costo unitario y línea base de inventario
scripts/02b_estabilidad_temporal.py     ¿la caída del consumo fue generalizada o concentrada?
scripts/03_objetivo2_clasificacion_abc_ved.py   Objetivo 2
scripts/04_objetivo3_niveles_inventario.py      Objetivo 3 y evaluación
scripts/04b_diagnostico_vitales.py      diagnóstico del criterio C1 (vitales por patrón)
notebooks/                              mismos scripts como cuadernos (Colab)
docs/                                   protocolo de evaluación, bitácora de decisiones
outputs/figuras/                        figuras generadas (las tablas con datos por producto no se publican)
data/                                   datos de entrada (NO se versionan)
tests/generar_base_sintetica.py         datos sintéticos para probar que el código corre
```

El cuaderno del Objetivo 1 (extracción, depuración y EDA) debe ubicarse en `notebooks/01_ETL_EDA_objetivo1.ipynb`
y terminar exportando la base analítica:
`base.to_parquet("data/base_analitica.parquet")` con las columnas
`CODIGO_PROPIO, DES_MEDICAMENTO, CANTIDAD, BODEGA, DES_UF, ANIO, MES, SERVICIO_FARMACEUTICO`.

## Cómo ejecutar

En Google Colab: abrir el cuaderno desde GitHub (reemplazar `github.com` por `colab.research.google.com/github`),
descomentar la primera celda para clonar el repositorio e instalar los requisitos, y subir los datos a `data/`.

En local:

```bash
pip install -r requirements.txt
python scripts/00_diagramas_de_flujo.py
RAW_DIR=data/raw_bases BODEGAS=data/bodegas.xlsx python scripts/00b_construir_base_analitica.py
DATA_DIR=data python scripts/01_diccionario_y_ficha.py
DATA_DIR=data python scripts/02_costos_y_linea_base.py
DATA_DIR=data python scripts/02b_estabilidad_temporal.py
DATA_DIR=data python scripts/03_objetivo2_clasificacion_abc_ved.py
DATA_DIR=data python scripts/04_objetivo3_niveles_inventario.py
DATA_DIR=data python scripts/04b_diagnostico_vitales.py
```

Los scripts 03 y 04 deben ejecutarse en ese orden (04 lee la clasificación que escribe 03).
Parámetros de diseño (cortes ABC, niveles de servicio, tiempo de reposición, ventana) en `src/inventario.py`
y al inicio de cada script.

## Disponibilidad de datos

Los datos son institucionales y no se publican en este repositorio (`data/` y `outputs/tablas/` están en `.gitignore`). Se retiraron
los campos de pacientes y ventas antes de cualquier procesamiento. El diccionario de datos y la ficha técnica
están en `outputs/tablas/diccionario_y_ficha.xlsx`. Para probar el código sin datos reales:

```bash
python tests/generar_base_sintetica.py data data_sintetica
```

Los resultados con datos sintéticos **no representan al hospital** y no deben citarse.

## Uso de inteligencia artificial

Se usó un asistente de IA (Claude) como apoyo para estructurar el código y los borradores de texto. Los autores
verifican las cifras contra la ejecución del código, las referencias contra sus fuentes originales y asumen la
responsabilidad por cada decisión (ver `docs/bitacora_decisiones.md`).
