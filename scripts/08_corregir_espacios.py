"""Corrige espacios sobrantes en el documento: dobles espacios dentro de un fragmento, dobles espacios entre fragmentos
consecutivos de un párrafo y espacios al inicio o al final de cada párrafo. No toca ecuaciones ni campos.

Uso:  python scripts/08_corregir_espacios.py entrada.docx salida.docx
"""
import re
import sys

from docx import Document

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def corregir(entrada, salida):
    doc = Document(entrada)
    cambios = {"dobles": 0, "cruces": 0, "inicio": 0, "final": 0}
    for p in doc.element.body.iter(W + "p"):
        ts = [t for t in p.iter(W + "t") if t.text]
        for t in ts:
            nuevo = re.sub(r"[  ]{2,}", " ", t.text)
            if nuevo != t.text:
                cambios["dobles"] += 1; t.text = nuevo
        for a, b in zip(ts, ts[1:]):                       # fragmentos consecutivos: "... " + " ..."
            if a.text.endswith(" ") and b.text.startswith(" "):
                b.text = b.text.lstrip(" "); cambios["cruces"] += 1
        ts = [t for t in ts if t.text]
        if ts and ts[0].text.startswith(" "):
            ts[0].text = ts[0].text.lstrip(" "); cambios["inicio"] += 1
        if ts and ts[-1].text.endswith(" "):
            ts[-1].text = ts[-1].text.rstrip(" "); cambios["final"] += 1
        for t in p.iter(W + "t"):                          # conservar los espacios que quedan en los bordes de un fragmento
            if t.text and (t.text.startswith(" ") or t.text.endswith(" ")):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    doc.save(salida)
    print(cambios)


if __name__ == "__main__":
    corregir(sys.argv[1], sys.argv[2])
