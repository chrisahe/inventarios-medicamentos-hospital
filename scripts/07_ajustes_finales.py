"""Ajustes finales sobre el documento ya editado: elimina el párrafo repetido de 8.2 y convierte las fórmulas de 6.1
(texto con caracteres matemáticos) en ecuaciones nativas de Word (OMML).

Uso:  python scripts/07_ajustes_finales.py entrada.docx salida.docx
"""
import re
import sys
import unicodedata

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' \
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def r(txt, recto=False):
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if recto else ""
    return f'<m:r>{sty}<m:t xml:space="preserve">{txt}</m:t></m:r>'


def frac(num, den):
    return f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"


def raiz(contenido):
    return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{contenido}</m:e></m:rad>'


def ecuacion(cuerpo):
    return parse_xml(f'<m:oMathPara {M}><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>{cuerpo}</m:oMath></m:oMathPara>')


# Ecuaciones de la Sección 6.1 (se identifican por su inicio en el texto convertido del PDF)
ECUACIONES = {
    "EOQ": r("EOQ", True) + r(" = ") + raiz(frac(r("2DS"), r("H"))),
    "StockMínimo": r("Stock mínimo", True) + r(" = ") + r("D") + r(" × ") + r("TR", True),
    "StockMáximo": r("Stock máximo", True) + r(" = ") + r("Stock mínimo", True) + r(" + ") + r("Lote de reposición", True),
    "SS": r("SS", True) + r(" = ") + r("Z") + r(" × ") + r("σ") + r(" × ") + raiz(r("TR", True)),
    "PR": r("PR", True) + r(" = ") + r("D") + r(" × ") + r("TR", True) + r(" + ") + r("SS", True),
}


def clave_ecuacion(texto):
    t = unicodedata.normalize("NFKC", texto).replace("\n", "").replace("\t", "").replace(" ", "")
    if t.startswith("EOQ="):
        return "EOQ"
    if t.startswith("StockM") and "Lote" not in t and t.endswith("TR") and "+" not in t:
        return "StockMínimo"
    if t.startswith("StockM") and "Lote" in t:
        return "StockMáximo"
    if t.startswith("SS="):
        return "SS"
    if t.startswith("PR="):
        return "PR"
    return None


def main(entrada, salida):
    doc = Document(entrada)
    P = doc.paragraphs

    # 1. párrafo repetido en 8.2: dos párrafos seguidos que repiten a los dos anteriores
    def comp(t):
        return re.sub(r"\s+", " ", t).strip().lower()
    for i in range(2, len(P) - 1):
        a_, b_ = comp(P[i - 2].text), comp(P[i].text)
        if len(a_) > 150 and a_ == b_:
            print("repetido eliminado:", b_[:70])
            quitar = [P[i]]
            if comp(P[i + 1].text)[:60] == comp(P[i - 1].text)[:60]:
                quitar.append(P[i + 1])
            for q in quitar:
                q._p.getparent().remove(q._p)
            break

    # 2. ecuaciones de la Sección 6.1 y párrafos "donde ..."
    P = doc.paragraphs
    inicio = next(k for k, p in enumerate(P) if p.style.name.startswith("Heading") and p.text.strip().startswith("6.1."))
    fin = next(k for k, p in enumerate(P) if k > inicio and p.style.name.startswith("Heading") and p.text.strip().startswith("6.2."))
    n_ec = 0
    for p in P[inicio:fin]:
        clave = clave_ecuacion(p.text) if re.search("[\U0001D400-\U0001D7FF]", p.text) else None
        if clave:
            for run in list(p._p.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")):
                p._p.remove(run)
            p._p.append(ecuacion(ECUACIONES[clave]))
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.left_indent = Cm(0)
            pf.line_spacing = 1.5; pf.space_before = Pt(6); pf.space_after = Pt(6)
            n_ec += 1
        elif re.search("[\U0001D400-\U0001D7FF]", p.text):          # "donde D representa..." con letras matemáticas
            for run in p.runs:
                run.text = re.sub(r"\s+", " ", unicodedata.normalize("NFKC", run.text).replace("\t", " "))
            if p.runs:
                p.runs[0].text = p.runs[0].text.lstrip()
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.first_line_indent = Cm(1.27); pf.left_indent = Cm(0)
    print("ecuaciones convertidas:", n_ec)
    doc.save(salida)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
