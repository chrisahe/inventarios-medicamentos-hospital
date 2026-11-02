"""Convierte las citas numéricas del documento del ACA 1, p. ej. (1,3), a citas autor-año de APA 7.

Uso:  python scripts/convertir_citas_apa7.py entrada.docx salida.docx

* Reemplaza solo citas entre paréntesis formadas por números de 1 a 2 cifras (no toca años como (2021)).
* Varias fuentes se ordenan alfabéticamente y se separan con punto y coma, como pide APA 7.
* Si el número no existe en la lista de referencias (p. ej. 23) NO se modifica y se informa para revisión manual.
* Trabaja sobre cada fragmento de texto (run) para conservar el formato; las citas partidas entre
  fragmentos se listan al final para corregirlas a mano.

La numeración corresponde a la lista de referencias del ACA 1 con las dos correcciones documentadas en
docs/bitacora_decisiones.md (referencias 4 y 5).
"""
import re
import sys as _s
try:
    _s.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import sys
from docx import Document

AUTOR_ANIO = {
    1: "Abu Zwaida et al., 2021",
    2: "Saha & Ray, 2019",
    3: "Silver et al., 1998",
    4: "Rehman et al., 2023",          # antes "Dube et al., 2023" (error de autoría)
    5: "Kabera & Mukanyangezi, 2024",  # antes "Muhoza et al., 2023" (error de autoría y año)
    6: "Devnani et al., 2010",
    7: "Ramanathan, 2006",
    8: "Ballou, 2004",
    9: "World Health Organization, 2011",
    10: "World Health Organization, 2017",
    11: "Wang et al., 2018",
    12: "Raghupathi & Raghupathi, 2014",
    13: "Choi et al., 2018",
    14: "Kelle et al., 2012",
    15: "Mustafee et al., 2017",
    16: "Naciones Unidas, s. f.-a",
    17: "Naciones Unidas, s. f.-b",
    18: "Chapman et al., 2000",
}
PATRON = re.compile(r"\((\d{1,2}(?:\s*[,–-]\s*\d{1,2})*)\s*\.?\s*\)")


def expandir(txt):
    nums = []
    partes = re.split(r"\s*,\s*", txt)
    for p in partes:
        m = re.fullmatch(r"(\d{1,2})\s*[–-]\s*(\d{1,2})", p)
        nums += list(range(int(m[1]), int(m[2]) + 1)) if m else [int(p)]
    return nums


def convertir(texto, sin_resolver):
    def sub(m):
        nums = expandir(m[1])
        malos = [n for n in nums if n not in AUTOR_ANIO]
        if malos:
            sin_resolver.append((m[0], malos))
            return m[0]
        citas = sorted({AUTOR_ANIO[n] for n in nums}, key=str.lower)
        return "(" + "; ".join(citas) + ")"
    return PATRON.sub(sub, texto)


def procesar(entrada, salida):
    doc = Document(entrada)
    sin_resolver, cambios = [], 0

    def parrafos(d):
        yield from d.paragraphs
        for t in d.tables:
            for fila in t.rows:
                for celda in fila.cells:
                    yield from celda.paragraphs

    for p in parrafos(doc):
        for r in p.runs:
            nuevo = convertir(r.text, sin_resolver)
            if nuevo != r.text:
                r.text, cambios = nuevo, cambios + 1
        # citas partidas entre fragmentos: se detectan en el texto completo del párrafo
        if PATRON.search(p.text):
            sin_resolver.append((p.text[:90] + "…", "cita partida entre fragmentos de formato"))
    doc.save(salida)
    print(f"{cambios} fragmentos modificados → {salida}")
    if sin_resolver:
        print("\nRevisar manualmente:")
        for item in sin_resolver:
            print("  ", item)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    procesar(sys.argv[1], sys.argv[2])
