"""Reconstruye la base analítica del Objetivo 1 a partir de los archivos mensuales de consumo.

Replica la depuración del cuaderno "Proyecto final parte 2" (decisiones D-01 a D-08 de la bitácora):
  1. Conserva 7 variables y excluye cualquier dato de pacientes o ventas.
  2. Excluye registros de bodegas con `Se utiliza = No`.
  3. Excluye productos cuya descripción inicia con "NPT".
  4. Bodega 829: solo se conservan las bolsas EVA de nutrición parenteral de 500, 1000 y 3000 ml.
  5. Excluye cantidades <= 0.
  6. Excluye oxígeno medicinal (código 100020287; lo gestiona Biomédica).
  7. Asigna SERVICIO_FARMACEUTICO desde bodegas.xlsx.

Uso:  RAW_DIR=data/raw_bases/Bases BODEGAS=../bodegas.xlsx python scripts/00b_construir_base_analitica.py
Entrada: RAW_DIR con 2024.zip, 2025.zip, 2026.zip (o carpetas 2024/, 2025/, 2026/ con los .xlsx).
Salida:  data/base_analitica.parquet y outputs/tablas/depuracion_resumen.csv
"""
import io
import os
import sys
import zipfile

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

RAW_DIR = os.environ.get("RAW_DIR", "data/raw_bases/Bases")
BODEGAS = os.environ.get("BODEGAS", "../bodegas.xlsx")
DATA_DIR = os.environ.get("DATA_DIR", "data")
OUT_TAB = os.environ.get("OUT_TAB", "outputs/tablas")

COLUMNAS = ["CODIGO_PROPIO", "DES_MEDICAMENTO", "CANTIDAD", "BODEGA", "DES_UF", "MES", "AÑO"]
MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6, "JULIO": 7,
         "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}
PERMITIDOS_829 = ["BOLSA EVA NUTRICION PARENTERAL 500 ML",
                  "BOLSA EVA NUTRICION PARENTERAL 1000 ML",
                  "BOLSA EVA NUTRICION PARENTERAL 3000 ML"]
CODIGO_OXIGENO = "100020287"


def leer_excel(origen):
    df = pd.read_excel(origen, engine="calamine")
    df.columns = df.columns.astype(str).str.strip().str.upper()
    return df[COLUMNAS].copy()


def cargar_anio(anio):
    zip_path = os.path.join(RAW_DIR, f"{anio}.zip")
    partes = []
    if os.path.exists(zip_path):
        with zipfile.ZipFile(zip_path) as z:
            for n in sorted(x for x in z.namelist() if x.lower().endswith(".xlsx")):
                partes.append(leer_excel(io.BytesIO(z.read(n))))
                print(f"  {n}: {len(partes[-1]):,}")
    else:
        carpeta = os.path.join(RAW_DIR, str(anio))
        for n in sorted(x for x in os.listdir(carpeta) if x.lower().endswith(".xlsx")):
            partes.append(leer_excel(os.path.join(carpeta, n)))
            print(f"  {n}: {len(partes[-1]):,}")
    return pd.concat(partes, ignore_index=True)


def main():
    resumen = []

    def paso(nombre, antes, despues):
        resumen.append({"paso": nombre, "registros_excluidos": antes - despues, "registros_restantes": despues})
        print(f"{nombre}: -{antes - despues:,} -> {despues:,}")

    base = pd.concat([cargar_anio(a) for a in (2024, 2025, 2026)], ignore_index=True)
    print("Registros iniciales:", f"{len(base):,}")
    resumen.append({"paso": "Registros iniciales", "registros_excluidos": 0, "registros_restantes": len(base)})

    bod = pd.read_excel(BODEGAS)
    bod.columns = bod.columns.astype(str).str.strip()
    bod["BODEGA_KEY"] = pd.to_numeric(bod["Bodega"], errors="coerce")
    uso = bod.drop_duplicates("BODEGA_KEY").set_index("BODEGA_KEY")["Se utiliza"]
    clave = pd.to_numeric(base["BODEGA"], errors="coerce")
    desc = base["DES_MEDICAMENTO"].astype(str).str.strip().str.upper()

    es_no = clave.map(uso).astype(str).str.strip().str.upper().eq("NO")
    es_npt = desc.str.startswith("NPT")
    excluir_829 = (clave == 829) & ~desc.isin(PERMITIDOS_829)
    excluir = es_no | es_npt | excluir_829
    print(f"  bodega No: {es_no.sum():,} | NPT: {es_npt.sum():,} | bodega 829 excluida: {excluir_829.sum():,}")
    n0 = len(base)
    base = base[~excluir].reset_index(drop=True)
    paso("Bodega No, NPT y bodega 829 no permitida", n0, len(base))

    n0 = len(base)
    base = base[base["CANTIDAD"] > 0].reset_index(drop=True)
    paso("Cantidad <= 0", n0, len(base))

    n0 = len(base)
    base = base[base["CODIGO_PROPIO"].astype(str).str.strip() != CODIGO_OXIGENO].reset_index(drop=True)
    paso("Oxígeno medicinal", n0, len(base))

    base["MES"] = base["MES"].astype(str).str.strip().str.upper().map(MESES).astype(int)
    base = base.rename(columns={"AÑO": "ANIO"})
    base["ANIO"] = base["ANIO"].astype(int)
    mapa = bod.drop_duplicates("Bodega").set_index("Bodega")["Servicio Farmaceutico"]
    base["SERVICIO_FARMACEUTICO"] = pd.to_numeric(base["BODEGA"], errors="coerce").map(mapa)
    print("Sin servicio farmacéutico:", f"{base['SERVICIO_FARMACEUTICO'].isna().sum():,}")
    print("Periodos:", base.groupby(["ANIO", "MES"]).ngroups)

    base["CODIGO_PROPIO"] = base["CODIGO_PROPIO"].astype(str).str.strip()
    base["BODEGA"] = base["BODEGA"].astype(str).str.strip()
    base["DES_MEDICAMENTO"] = base["DES_MEDICAMENTO"].astype(str)
    base["DES_UF"] = base["DES_UF"].astype(str)
    cols = ["CODIGO_PROPIO", "DES_MEDICAMENTO", "CANTIDAD", "BODEGA", "DES_UF", "ANIO", "MES", "SERVICIO_FARMACEUTICO"]
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(OUT_TAB, exist_ok=True)
    base[cols].to_parquet(os.path.join(DATA_DIR, "base_analitica.parquet"), index=False)
    pd.DataFrame(resumen).to_csv(os.path.join(OUT_TAB, "depuracion_resumen.csv"), index=False)
    print("Base analítica final:", f"{len(base):,}", "registros")


if __name__ == "__main__":
    main()
