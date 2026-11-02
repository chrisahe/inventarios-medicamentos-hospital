"""Genera una base de consumo SINTÉTICA para probar que los scripts corren de extremo a extremo.

NO representa al hospital: los resultados obtenidos con ella no deben citarse en el documento.
Usa los códigos reales de la Biblia solo para que las uniones de costo funcionen.

Uso:  python tests/generar_base_sintetica.py <carpeta_datos> <carpeta_salida> [n_productos]
"""
import sys, os
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))
import inventario as inv

datos, salida = sys.argv[1], sys.argv[2]
n = int(sys.argv[3]) if len(sys.argv) > 3 else 2500
os.makedirs(salida, exist_ok=True)
rng = np.random.default_rng(7)

b = inv.leer_biblia(os.path.join(datos, "Biblia (1).xlsx"))
b = b[b["ES_ACTIVO"] & b["COSTO_FINAL"].notna()]
prod = b.drop_duplicates("CODIGO_PROPIO").sample(n, random_state=1)
bod_por_cod = b.groupby("CODIGO_PROPIO")["BODEGA"].first()
servicios = ["Hospitalización", "Almacén", "UCI-UCE", "Cirugía", "Urgencias"]
periodos = [(a, m) for a in (2024, 2025, 2026) for m in range(1, 13) if (a, m) <= (2026, 5)]

filas = []
for cod, des in zip(prod["CODIGO_PROPIO"], prod["DES_MEDICAMENTO"]):
    tipo = rng.choice(["sost", "frec", "inter", "espor"], p=[0.10, 0.12, 0.23, 0.55])
    nivel = rng.lognormal(3.0, 1.5)
    for i, (a, m) in enumerate(periodos):
        tendencia = 1.0 - 0.012 * i
        p_act = {"sost": 0.97, "frec": 0.65, "inter": 0.25, "espor": 0.06}[tipo]
        if rng.random() < p_act:
            q = max(1, int(rng.gamma(4, nivel * tendencia / 4)))
            for s in rng.choice(servicios, size=rng.integers(1, 3), replace=False):
                filas.append((cod, des, q, bod_por_cod[cod], "UF", a, m, s))
df = pd.DataFrame(filas, columns=inv.COLUMNAS_BASE)
df.to_parquet(os.path.join(salida, "base_analitica.parquet"), index=False)
print(len(df), "registros sintéticos;", df["CODIGO_PROPIO"].nunique(), "productos")
