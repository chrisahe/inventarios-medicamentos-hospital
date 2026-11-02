"""Arma el documento del ACA 2: documento del ACA 1 (convertido desde el PDF) + contenido nuevo, en APA 7.

Pasos: limpia la conversión (números de página, saltos de sección, párrafos partidos), reconoce los títulos con los
marcadores del PDF, convierte las citas numéricas a autor-año, reconstruye la lista de referencias (APA 7, orden
alfabético, sangría francesa), inserta el contenido de `05_generar_documento.py` en la Sección 8 y aplica el formato.

Uso:  python scripts/06_integrar_documento.py ACA1.pdf ACA1_convertido.docx salida.docx
Variables: OUT_TAB, OUT_FIG, REPO_URL, COLAB_URL (las usa 05_generar_documento.py).
"""
import copy, os, re, runpy, sys
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph
import pymupdf

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from convertir_citas_apa7 import convertir

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import importlib.util, io
from PIL import Image
_spec = importlib.util.spec_from_file_location("pasos", os.path.join(os.path.dirname(os.path.abspath(__file__)), "00c_diagramas_pasos_objetivo1.py"))
_mod = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(_mod)
PASOS = _mod.PASOS
DIR_PASOS = os.path.join(os.environ.get("OUT_FIG", "outputs/figuras"), "pasos")

PDF, BASE, SALIDA = sys.argv[1], sys.argv[2], sys.argv[3]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FUENTE, NEGRO = "Times New Roman", RGBColor(0, 0, 0)

# ---------------------------------------------------------------- 1. contenido nuevo sobre el documento base
os.environ["BASE_DOCX"] = BASE
sys.argv = ["05", os.path.join(os.path.dirname(os.path.abspath(BASE)), "_nuevo_tmp.docx")]
ns = runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "05_generar_documento.py"))
doc, N_BASE = ns["doc"], ns["N_BASE"]
body = doc.element.body
elems = list(body)
assert elems[-1].tag == W + "sectPr"
base_e, new_e = elems[:N_BASE - 1], elems[N_BASE - 1:-1]
sect_final = elems[-1]


def txt(e):
    return Paragraph(e, doc).text if e.tag == W + "p" else ""


def norm(s):
    return re.sub(r"\s+", " ", s).strip().lower()


def quitar_numero(s):
    return re.sub(r"^[\d\.\s]+(?=[A-Za-zÁÉÍÓÚÑ¿])", "", s.strip())


def tiene_dibujo(e):
    return e.find(".//" + W + "drawing") is not None


def poner_texto(p, texto, negrita=None, cursiva=None):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    r = p.add_run(texto)
    if negrita is not None:
        r.bold = negrita
    if cursiva is not None:
        r.italic = cursiva
    return r


# ---------------------------------------------------------------- 2. limpieza del documento base
# 2.0 tablas de maquetación (fragmentos de títulos) y meses partidos del cronograma
for e in list(base_e):
    if e.tag == W + "tbl":
        celdas = [" ".join("".join(x.text or "" for x in c.iter(W + "t")).split()) for c in e.iter(W + "tc")]
        if len(celdas) == 4 and re.fullmatch(r"\d+(\.\d+)*\.?", celdas[0]) and re.fullmatch(r"\d+(\.\d+)*\.?", celdas[2]):
            print("tabla de maquetación eliminada:", celdas)
            body.remove(e); base_e.remove(e)
        elif any(re.fullmatch(r"Di\s*c( 2025)?", c) for c in celdas):
            for t_ in e.iter(W + "t"):
                if t_.text and t_.text.strip() in ("Di c", "En e", "Fe b", "Se p", "Oc t", "No v"):
                    t_.text = t_.text.replace(" ", "")
            from docx.table import Table
            tb_ = Table(e, doc); tb_.autofit = False
            for cel in tb_.rows[0].cells:
                m_ = re.match(r"^([A-Za-z]{1,2})\s+([a-z]{1,2})\s+(\d{4})$", " ".join(cel.text.split()))
                if m_:
                    for extra in cel.paragraphs[1:]:
                        extra._p.getparent().remove(extra._p)
                    poner_texto(cel.paragraphs[0], f"{m_[1]}{m_[2]} {m_[3]}", negrita=True)
            for fila in tb_.rows:
                for j, cel in enumerate(fila.cells):
                    cel.width = Cm(4.6) if j == 0 else Cm((15.9 - 4.6) / (len(fila.cells) - 1))
            continue
        elif False:
            for t_ in e.iter(W + "t"):
                t_.text = re.sub(r"^([A-Za-z]{1,2}) ([a-z]{1,2}) (\d{4})$", r" ", t_.text or "")
        # texto con espacios partidos en las celdas del cronograma ("Di c 2025")
    elif e.tag == W + "p":
        for t_ in e.iter(W + "t"):
            if t_.text and "enero2024" in t_.text:
                t_.text = t_.text.replace("enero2024", "enero 2024")

# 2.1 saltos de sección (uno por página en la conversión): se eliminan para que el texto fluya
for e in base_e:
    if e.tag == W + "p":
        ppr = e.find(W + "pPr")
        if ppr is not None and ppr.find(W + "sectPr") is not None:
            ppr.remove(ppr.find(W + "sectPr"))

# 2.2 marcadores de página (párrafos con solo un número): sirven para ubicar los títulos perdidos
paginas = {}
for e in base_e:
    t = txt(e).strip()
    if e.tag == W + "p" and re.fullmatch(r"\d{1,3}", t) and not tiene_dibujo(e):
        paginas.setdefault(int(t), e)

# 2.3 títulos del PDF (marcadores): niveles por numeración
marc = [(t.strip(), pg) for _, t, pg in pymupdf.open(PDF).get_toc() if not re.match(r"(Tabla|Figura) \d+\.", t.strip())]


def nivel(titulo):
    m = re.match(r"^(\d+(?:\.\d+)*)\.?\s", titulo + " ")
    if m:
        return min(m.group(1).count(".") + 1, 3)
    return 2 if titulo.startswith(("Preliminares")) else (3 if titulo.startswith("Declaración") else 1)


def clave(titulo):
    return norm(quitar_numero(titulo))[:40]


def compacto(s):
    return re.sub(r"[\s\.	]+", "", s.lower())


def coincide(t, titulo):
    a, b = compacto(t), compacto(titulo)
    return a == b or (len(a) >= 12 and a.startswith(b[:40]) and len(a) <= len(b) + 3)


inicio = next(i for i, e in enumerate(base_e) if e is paginas[10]) + 1
ultimo_insertado = {}
for titulo, pg in marc:                      # reinserta los títulos que la conversión perdió
    k = clave(titulo)
    esta = any(coincide(txt(e).strip(), titulo) for e in base_e[inicio:] if e.tag == W + "p")
    if not esta and pg in paginas:
        ref = ultimo_insertado.get(pg, paginas[pg])
        nuevo = OxmlElement("w:p"); ref.addnext(nuevo); ultimo_insertado[pg] = nuevo
        r = OxmlElement("w:r"); t_ = OxmlElement("w:t"); t_.text = titulo; r.append(t_); nuevo.append(r)
        print("título reinsertado:", titulo, "(p.", pg, ")")
NUEVOS = set(new_e)
base_e = [e for e in list(body)[:-1] if e not in NUEVOS]

# 2.4 quitar marcadores de página y párrafos vacíos (sin dibujos ni saltos)
def vacio(e):
    return e.tag == W + "p" and not txt(e).strip() and not tiene_dibujo(e) and e.find(".//" + W + "br") is None


for e in list(base_e):
    t = txt(e).strip()
    if e.tag == W + "p" and ((re.fullmatch(r"\d{1,3}", t) and not tiene_dibujo(e)) or vacio(e)):
        body.remove(e)
base_e = [e for e in list(body)[:-1] if e not in NUEVOS]

# 2.5 une los párrafos partidos entre páginas (el segundo empieza en minúscula y el primero no cierra la oración)
def es_texto_corrido(e):
    return e.tag == W + "p" and not tiene_dibujo(e)


i = 1
while i < len(base_e):
    a, b = base_e[i - 1], base_e[i]
    ta, tb = txt(a).rstrip(), txt(b).lstrip("\t ")
    if es_texto_corrido(a) and es_texto_corrido(b) and len(ta) > 40 and tb[:1].islower() \
            and not re.search(r"[\.:;\?!\)»”\"]\s*$", ta) and "\t•" not in ta:
        pa, pb = Paragraph(a, doc), Paragraph(b, doc)
        pa.add_run(" ")
        for r in list(pb._p.findall(W + "r")):
            a.append(r)
        body.remove(b); base_e.pop(i)
    else:
        i += 1

# 2.6 duplicados señalados en la retroalimentación (primer párrafo de Alcance repetido en Limitaciones)
vistos = {}
for e in list(base_e):
    t = txt(e).strip()
    if e.tag == W + "p" and len(t) > 200:
        if t in vistos and t.startswith("El presente proyecto estará orientado"):
            print("párrafo duplicado eliminado:", t[:70])
            body.remove(e); base_e.remove(e)
        elif t in vistos:
            print("AVISO párrafo repetido (se conserva, revisar):", t[:90])
        else:
            vistos[t] = e

# ---------------------------------------------------------------- 3. títulos con estilos de encabezado
pos = 0
paras = [e for e in base_e if e.tag == W + "p"]
inicio = next(i for i, e in enumerate(paras) if txt(e).strip() == "Resumen" and i > 20) if True else 0
heads = {}
for titulo, pg in marc:
    k = clave(titulo)
    for j in range(pos, len(paras)):
        t = txt(paras[j]).strip()
        if j >= inicio and coincide(t, titulo):
            heads[paras[j]] = (titulo, nivel(titulo)); pos = j + 1; break
    else:
        print("sin coincidencia para:", titulo)
for e, (titulo, niv) in heads.items():
    p = Paragraph(e, doc)
    if e.find(W + "pPr") is not None:
        e.remove(e.find(W + "pPr"))
    p.style = doc.styles[f"Heading {niv}"]
    poner_texto(p, re.sub(r"\s+", " ", titulo))

# ---------------------------------------------------------------- 4. párrafo 4 y 5 (autores corregidos) y citas a autor-año
CORRECCIONES = {
    "Dube et al. (2023) evidenciaron": (
        "Durante escenarios de alta presión asistencial, como la pandemia por COVID-19, Rehman et al. (2023) modelaron y compararon "
        "mediante simulación políticas de inventario hospitalario (punto de reorden, pedido periódico y just-in-time) para evaluar su desempeño."),
    "Muhoza et al. (2023) demostraron": (
        "Kabera y Mukanyangezi (2024) analizaron la influencia de las prácticas de gestión de inventarios sobre la disponibilidad de medicamentos "
        "obstétricos de emergencia en hospitales públicos de la Provincia Sur de Ruanda, un contexto donde los recursos son limitados."),
}
for e in base_e:
    if e.tag == W + "p":
        t = txt(e)
        for clave_, nuevo in CORRECCIONES.items():
            if clave_ in t:
                poner_texto(Paragraph(e, doc), nuevo); print("párrafo reescrito:", clave_)
sin_resolver = []
fin_cuerpo = next(i for i, e in enumerate(base_e) if e in heads and heads[e][0].startswith("9."))
for e in base_e[:fin_cuerpo]:
    if e.tag == W + "p":
        for r in Paragraph(e, doc).runs:
            n = convertir(r.text.replace("(23, 12)", "(12)"), sin_resolver)     # la fuente 23 no existe en la lista de referencias
            if n != r.text:
                r.text = n
    elif e.tag == W + "tbl":
        for p in [Paragraph(x, doc) for x in e.iter(W + "p")]:
            for r in p.runs:
                r.text = convertir(r.text, sin_resolver)

# ---------------------------------------------------------------- 5. referencias en APA 7
paras = [e for e in base_e if e.tag == W + "p"]
ih = next(i for i, e in enumerate(paras) if e in heads and heads[e][0].startswith("9."))
cola, entradas, esperado = paras[ih + 1:], [], 1
for e in cola:
    t = txt(e).strip()
    if re.match(rf"^{esperado}\.\s", t + " ") and (not entradas or True):
        entradas.append([e]); esperado += 1
    elif entradas:
        entradas[-1].append(e)
print("referencias del ACA 1:", len(entradas))
NUEVAS = {4: [("Rehman, A. U., Mian, S. H., Usmani, Y. S., Abidi, M. H., & Mohammed, M. K. (2023). Modelling and analysis of hospital inventory policies during COVID-19 pandemic. ", 0),
              ("Processes, 11", 1), ("(4), 1062. https://doi.org/10.3390/pr11041062", 0)],
          5: [("Kabera, J. C., & Mukanyangezi, M. F. (2024). Influence of inventory management practices on the availability of emergency obstetric drugs in Rwandan public hospitals: A case of Rwanda Southern Province. ", 0),
              ("BMC Health Services Research, 24", 1), (", Artículo 14. https://doi.org/10.1186/s12913-023-10459-x", 0)]}
EXTRA = [[("Hyndman, R. J., & Athanasopoulos, G. (2021). ", 0), ("Forecasting: Principles and practice", 1), (" (3.ª ed.). OTexts. https://otexts.com/fpp3/", 0)],
         [("Syntetos, A. A., & Boylan, J. E. (2005). The accuracy of intermittent demand estimates. ", 0),
          ("International Journal of Forecasting, 21", 1), ("(2), 303–314. https://doi.org/10.1016/j.ijforecast.2004.10.001", 0)]]
lista = []
for n, grupo in enumerate(entradas, 1):
    p0 = Paragraph(grupo[0], doc)
    for extra in grupo[1:]:                      # líneas de continuación al primer párrafo
        p0.add_run(" ")
        for r in list(extra.findall(W + "r")):
            grupo[0].append(r)
        body.remove(extra)
    if n in NUEVAS:
        poner_texto(p0, "")
        for t, cur in NUEVAS[n]:
            p0.add_run(t).italic = bool(cur)
    else:
        runs = p0.runs
        if runs:                                 # quita la numeración "n.\t"
            runs[0].text = re.sub(rf"^\s*{n}\.\s*", "", runs[0].text)
        for r in p0.runs:
            r.text = r.text.replace("\t", " ")
        if "(s. f.)" in p0.text:
            suf = "-a" if "Objetivo de Desarrollo Sostenible 3" in p0.text else "-b"
            for r in p0.runs:
                r.text = r.text.replace("(s. f.)", f"(s. f.{suf})")
    lista.append(grupo[0])
for partes in EXTRA:
    par = OxmlElement("w:p"); lista[-1].addnext(par); lista.append(par)
    p = Paragraph(par, doc)
    for t, cur in partes:
        p.add_run(t).italic = bool(cur)
ITALICAS = ["Applied Sciences, 11", "Business logistics/supply chain management: Planning, organizing, and controlling the supply chain",
            "Production and Operations Management, 27", "Journal of Young Pharmacists, 2", "Operations Research for Health Care, 1",
            "Journal of the Operational Research Society, 68", "Health Information Science and Systems, 2",
            "Computers & Operations Research, 33", "Computers & Industrial Engineering, 137", "Inventory management and production planning and scheduling",
            "Technological Forecasting and Social Change, 126", "The world medicines situation 2011: Access to essential medicines as part of the right to health",
            "Guidelines on the management of health supply chains", "Objetivo de Desarrollo Sostenible 3: Salud y bienestar",
            "Objetivo de Desarrollo Sostenible 12: Producción y consumo responsables"]
for e in lista:
    p = Paragraph(e, doc)
    if any(r.italic for r in p.runs):
        continue
    texto = " ".join(p.text.split()).replace("– ", "–")
    for sub in ITALICAS:
        if sub in texto:
            a, b = texto.split(sub, 1)
            poner_texto(p, a); p.add_run(sub).italic = True; p.add_run(b)
            break
    else:
        poner_texto(p, texto)
for e in lista:
    _p = Paragraph(e, doc)
    if _p.runs:
        _p.runs[0].text = _p.runs[0].text.lstrip()
lista.sort(key=lambda e: norm(txt(e)))
ancla = paras[ih]
for e in lista:
    body.remove(e) if e.getparent() is not None else None
for e in reversed(lista):
    ancla.addnext(e)
    pf = Paragraph(e, doc).paragraph_format
    pf.left_indent, pf.first_line_indent = Cm(1.27), Cm(-1.27)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.line_spacing = 2.0; pf.space_after = Pt(0); pf.space_before = Pt(0)

# ---------------------------------------------------------------- 6. rótulos de tablas y figuras (APA 7: rótulo en negrita y título en cursiva)
rot = re.compile(r"^(Tabla|Figura) (\d+)\.\s*(.*)$", re.S)
leyendas = {"Tabla": [], "Figura": []}
for e in list(base_e):
    if e.tag != W + "p" or not e.getparent() is body:
        continue
    t = txt(e).strip()
    m = rot.match(t)
    if m and e not in heads and t != "" and e in set(list(body)):
        # solo los rótulos del cuerpo (los de las listas iniciales están antes del Resumen)
        if list(body).index(e) < list(body).index(next(x for x in body if x in heads and heads[x][0] == "Resumen")):
            continue
        p = Paragraph(e, doc)
        poner_texto(p, f"{m.group(1)} {m.group(2)}", negrita=True)
        titulo = " ".join(m.group(3).split())
        num = int(m.group(2))
        reemplazar = m.group(1) == "Figura" and (num in PASOS or num == 35)
        if reemplazar and num in PASOS:
            titulo = PASOS[num][0]
        p2 = OxmlElement("w:p"); e.addnext(p2)
        Paragraph(p2, doc).add_run(titulo).italic = True
        if reemplazar:                       # sustituye la captura de código por el diagrama de flujo (o recorta el gráfico)
            img = p2.getnext()
            while img is not None and not tiene_dibujo(img):
                img = img.getnext()
            if img is not None:
                if num == 35:                # la figura mezcla código y gráfico: se conserva solo el gráfico
                    blip = next(img.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
                    blob = doc.part.related_parts[blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")].blob
                    im = Image.open(io.BytesIO(blob)).convert("RGB"); im = im.crop((0, 262, 700, im.size[1] - 10))
                    ruta_img = os.path.join(DIR_PASOS, "fig_35_recorte.png"); os.makedirs(DIR_PASOS, exist_ok=True); im.save(ruta_img)
                    ancho = 12.5
                else:
                    ruta_img, ancho = os.path.join(DIR_PASOS, f"fig_{num:02d}.png"), 11.5
                for r_ in list(img.findall(W + "r")):
                    if r_.find(".//" + W + "drawing") is not None:
                        img.remove(r_)
                Paragraph(img, doc).add_run().add_picture(ruta_img, width=Cm(ancho))
                resto = img.getnext()            # otras partes de la misma captura (imágenes en párrafos siguientes)
                while resto is not None and resto.tag == W + "p" and not txt(resto).strip():
                    sig_ = resto.getnext()
                    if tiene_dibujo(resto):
                        body.remove(resto)
                    resto = sig_
                if num != 35:                # la descripción ya no habla de "código" sino del procedimiento
                    sig = img.getnext()
                    while sig is not None and not txt(sig).strip():
                        sig = sig.getnext()
                    if sig is not None and sig.tag == W + "p":
                        for r_ in Paragraph(sig, doc).runs:
                            r_.text = re.sub(r"el código desarrollado en Python para", "el procedimiento desarrollado en Python para", r_.text)
                            r_.text = re.sub(r"^(\s*)El código ", lambda m_: m_.group(1) + "El procedimiento ", r_.text)
        for q in (p, Paragraph(p2, doc)):
            q.paragraph_format.keep_with_next = True
        leyendas[m.group(1)].append((m.group(2), titulo, p2))
print({k: len(v) for k, v in leyendas.items()})

# ---------------------------------------------------------------- 7. inserción del contenido nuevo
bloques, actual = [], None
for e in new_e:
    sty = Paragraph(e, doc).style.name if e.tag == W + "p" else ""
    if sty in ("Heading 1", "Heading 2"):
        actual = [txt(e).strip(), []]; bloques.append(actual)
    if actual is not None:
        actual[1].append(e)
ORDEN = [("Disponibilidad de datos y código", "8.13.1. Disponibilidad de datos y código", 3),
         ("8.1.13", "8.13.2. Análisis descriptivo del costo unitario", 3),
         ("8.1.14", "8.13.3. Estabilidad temporal del consumo", 3),
         ("8.1.15", "8.13.4. Del análisis exploratorio a las decisiones de inventario", 3),
         ("Diagramas de flujo", "8.13.5. Diagramas de flujo generales del Objetivo 1", 3),
         ("7.9 Protocolo", "8.14. Protocolo de evaluación", 2),
         ("8.2 Desarrollo", "8.15. Desarrollo del Objetivo Específico 2", 2),
         ("8.3 Desarrollo", "8.16. Desarrollo del Objetivo Específico 3", 2),
         ("Ficha técnica", "Anexo A. Ficha técnica de las fuentes y diccionario de datos", 1),
         ("Anexo B", None, 1)]
sel = {}
for clave_, nuevo, niv in ORDEN:
    b = next(b for b in bloques if b[0].startswith(clave_))
    sel[clave_] = (b, nuevo, niv)
    if nuevo:
        h = Paragraph(b[1][0], doc); h.style = doc.styles[f"Heading {niv}"]; poner_texto(h, nuevo)
    # quitar saltos de página sueltos dentro del bloque
    b[1][:] = [x for x in b[1] if not (x.tag == W + "p" and not txt(x).strip() and x.find(".//" + W + "br") is not None)
               ] if clave_ != "Anexo B" else b[1]
for e in new_e:
    body.remove(e)
ref_head = next(e for e, (t, n) in heads.items() if t.startswith("9."))
cursor = ref_head.getprevious()


def insertar_antes(ancla, elementos):
    for e in elementos:
        ancla.addprevious(e)


h813 = OxmlElement("w:p"); insertar_antes(ref_head, [h813])
ph = Paragraph(h813, doc); ph.style = doc.styles["Heading 2"]; poner_texto(ph, "8.13. Complementos al Objetivo Específico 1")
intro = OxmlElement("w:p"); insertar_antes(ref_head, [intro])
Paragraph(intro, doc).add_run("Esta sección reúne los análisis complementarios del Objetivo 1 que responden a la retroalimentación del ACA 1 y que "
                              "sustentan el diseño de los Objetivos 2 y 3.")
for clave_ in ["Disponibilidad de datos y código", "8.1.13", "8.1.14", "8.1.15", "Diagramas de flujo"]:
    insertar_antes(ref_head, sel[clave_][0][1])
    if clave_ == "Diagramas de flujo":
        pass
for clave_ in ["7.9 Protocolo", "8.2 Desarrollo", "8.3 Desarrollo"]:
    insertar_antes(ref_head, sel[clave_][0][1])
# anexos al final del cuerpo (después de las referencias)
fin = sect_final
for clave_ in ["Ficha técnica", "Anexo B"]:
    pb = OxmlElement("w:p"); r = OxmlElement("w:r"); br = OxmlElement("w:br"); br.set(qn("w:type"), "page"); r.append(br); pb.append(r)
    fin.addprevious(pb)
    for e in sel[clave_][0][1]:
        fin.addprevious(e)

# ---------------------------------------------------------------- 8. renumeración y referencias cruzadas del contenido nuevo
cruz = [(r"\b8\.2\.(\d)", r"8.15.\1"), (r"\b8\.3\.(\d)", r"8.16.\1"), (r"Sección 8\.2\b", "Sección 8.15"), (r"Sección 8\.3\b", "Sección 8.16"),
        (r"Sección 7\.9\b", "Sección 8.14"), (r"Secciones 8\.2\.(\d)", r"Secciones 8.15.\1")]
nuevos_orden = [e for e in body if e in set(sum([sel[k][0][1] for k in sel], [])) or e in {h813, intro}]
mapa_num = {"Tabla": {}, "Figura": {}}
cont = {"Tabla": max(int(n) for n, _, _ in leyendas["Tabla"]), "Figura": max(int(n) for n, _, _ in leyendas["Figura"])}
for e in nuevos_orden:
    t = txt(e).strip()
    m = re.fullmatch(r"(Tabla|Figura) (\d+)", t)
    if m and e.tag == W + "p":
        cont[m.group(1)] += 1
        mapa_num[m.group(1)][int(m.group(2))] = cont[m.group(1)]
print("tablas nuevas:", mapa_num["Tabla"]); print("figuras nuevas:", mapa_num["Figura"])


def renumerar(s):
    for pat, rep in cruz:
        s = re.sub(pat, rep, s)
    return re.sub(r"\b(Tabla|Figura) (\d+)\b", lambda m: f"{m[1]} {mapa_num[m[1]].get(int(m[2]), int(m[2]))}", s)


def todos_los_parrafos(e):
    return [Paragraph(x, doc) for x in e.iter(W + "p")]


rotulos_nuevos = {"Tabla": [], "Figura": []}
for e in nuevos_orden:
    for p in todos_los_parrafos(e):
        for r in p.runs:
            if r.text:
                r.text = renumerar(r.text)
    t = txt(e).strip()
    m = re.fullmatch(r"(Tabla|Figura) (\d+)", t) if e.tag == W + "p" else None
    if m:
        sig = e.getnext()
        rotulos_nuevos[m.group(1)].append((m.group(2), txt(sig).strip()))

# ---------------------------------------------------------------- 8.5 saltos de línea y tabulaciones heredados del PDF
# La conversión conserva el final de cada renglón del PDF como salto de línea dentro del párrafo; se reemplazan por espacios.
resumen_el = next(e for e, (t, n) in heads.items() if t == "Resumen")
refs_el = next(e for e, (t, n) in heads.items() if t.startswith("9."))
zona = False
for e in list(body):
    if e is resumen_el:
        zona = True
    if e is refs_el:
        break
    if not zona or e.tag != W + "p" or e in heads or tiene_dibujo(e):
        continue
    t = txt(e)
    if "•" in t or re.search("[\U0001D400-\U0001D7FF]", t):
        continue
    lineas = t.split("\n")
    if len(lineas) > 1 and all(len(x.strip()) < 60 for x in lineas):
        continue                                          # bloques de líneas cortas (firmas, fechas)
    for br in list(e.iter(W + "br")):
        if br.get(W + "type") in (None, "textWrapping"):
            nuevo_t = OxmlElement("w:t"); nuevo_t.text = " "; nuevo_t.set(qn("xml:space"), "preserve")
            br.getparent().replace(br, nuevo_t)
    for tab in list(e.iter(W + "tab")):
        if tab.getparent().tag == W + "r":
            nuevo_t = OxmlElement("w:t"); nuevo_t.text = " "; nuevo_t.set(qn("xml:space"), "preserve")
            tab.getparent().replace(tab, nuevo_t)
    for t_ in e.iter(W + "t"):
        if t_.text:
            t_.text = re.sub(r"[ \t]{2,}", " ", t_.text)

# ---------------------------------------------------------------- 9. formato APA 7 sobre todo el documento (texto negro)
caption_rx = re.compile(r"^(Tabla|Figura) \d+$")
portada_fin = next((e for e in body if e.tag == W + "p" and txt(e).strip() == "Tabla de contenido"), None)
portada_ids = set()
for e in body:
    if e is portada_fin:
        break
    if e.tag == W + "p":
        portada_ids.add(e)
in_tabla = {x for tb in body.iter(W + "tbl") for x in tb.iter(W + "p")}
for p_el in body.iter(W + "p"):
    p = Paragraph(p_el, doc)
    for r in p.runs:
        r.font.name = FUENTE; r._r.get_or_add_rPr()
        rf = r._r.rPr.find(W + "rFonts")
        if rf is not None:
            rf.set(qn("w:eastAsia"), FUENTE)
        r.font.color.rgb = NEGRO; r.font.highlight_color = None
        if p_el not in in_tabla and r.font.size is not None and not caption_rx.match(p.text):
            r.font.size = Pt(12) if r.font.size >= Pt(11) else r.font.size
    pf = p.paragraph_format
    estilo = p.style.name
    es_titulo = estilo.startswith("Heading")
    _prev = p_el.getprevious()
    _titulo_rotulo = _prev is not None and caption_rx.match(txt(_prev).strip())
    if not es_titulo and p_el not in in_tabla and not _titulo_rotulo and len(p.text) > 40 and p.runs             and not caption_rx.match(p.text.strip()) and not p.text.startswith("Nota."):
        _n_cur = sum(len(r.text) for r in p.runs if r.italic)
        if _n_cur > 0.9 * sum(len(r.text) for r in p.runs) and p_el not in portada_ids:
            for r in p.runs:
                r.italic = False
    if es_titulo:
        continue
    if p_el in in_tabla:
        pf.first_line_indent = Cm(0); pf.line_spacing = 1.0; pf.space_after = Pt(0)
        continue
    t = p.text.strip()
    if portada_fin is not None and p_el in portada_ids:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.line_spacing = 1.15; pf.space_after = Pt(8)
        for r in p.runs:
            r.bold = r.bold or t.startswith("Modelo analítico")
        continue
    if t in ("Tabla de contenido", "Lista de tablas", "Lista de figuras"):
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.line_spacing = 2.0; pf.page_break_before = True
        for r in p.runs:
            r.bold = True
        continue
    if tiene_dibujo(p_el):
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.line_spacing = 1.0; pf.keep_with_next = True
    elif caption_rx.match(t) or (p_el.getprevious() is not None and caption_rx.match(txt(p_el.getprevious()).strip()) and p.runs and p.runs[0].italic):
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.first_line_indent = Cm(0); pf.left_indent = Cm(0); pf.line_spacing = 2.0; pf.keep_with_next = True
    elif t.startswith(("Nota.", "Fuente:")):
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.first_line_indent = Cm(0); pf.line_spacing = 1.0; pf.space_after = Pt(12)
    elif estilo == "List Bullet" or "•" in t[:4]:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.line_spacing = 2.0
        pf.left_indent, pf.first_line_indent = Cm(1.9), Cm(-0.63)
        if p.runs:
            p.runs[0].text = p.runs[0].text.lstrip("	")
    elif re.search("[\U0001D400-\U0001D7FF]", t):
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.line_spacing = 2.0
    elif pf.first_line_indent is not None and pf.first_line_indent < 0:
        pass                                             # referencias (sangría francesa)
    else:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.line_spacing = 2.0; pf.space_after = Pt(0); pf.space_before = Pt(0)
        pf.first_line_indent = Cm(1.27); pf.left_indent = Cm(0)
        if p.runs and p.runs[0].text.startswith("\t"):
            p.runs[0].text = p.runs[0].text.lstrip("\t")
for s in doc.sections:
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2.54)

# ---------------------------------------------------------------- 10. rótulos con numeración automática (SEQ) y listas
def run_campo(tipo, instr=None, texto=None, negrita=True):
    r = OxmlElement("w:r")
    if negrita:
        rpr = OxmlElement("w:rPr"); b = OxmlElement("w:b"); rpr.append(b); r.append(rpr)
    if tipo in ("begin", "separate", "end"):
        c = OxmlElement("w:fldChar"); c.set(qn("w:fldCharType"), tipo)
    elif tipo == "instr":
        c = OxmlElement("w:instrText"); c.set(qn("xml:space"), "preserve"); c.text = instr
    else:
        c = OxmlElement("w:t"); c.text = texto; c.set(qn("xml:space"), "preserve")
    r.append(c)
    return r


# Cada rótulo queda en un solo párrafo: "Tabla N" (campo SEQ, negrita) + salto de línea + título en cursiva (APA 7).
for e in list(body):
    if e.tag != W + "p":
        continue
    m = re.fullmatch(r"(Tabla|Figura) (\d+)", txt(e).strip())
    nxt = e.getnext()
    if not m or nxt is None or nxt.tag != W + "p" or not txt(nxt).strip():
        continue
    for r in list(e.findall(W + "r")):
        e.remove(r)
    e.append(run_campo("texto", texto=m.group(1) + " "))
    for tipo, instr, texto in (("begin", None, None), ("instr", f" SEQ {m.group(1)} \\* ARABIC ", None), ("separate", None, None),
                               ("texto", None, m.group(2)), ("end", None, None)):
        e.append(run_campo(tipo, instr, texto))
    salto = OxmlElement("w:r"); salto.append(OxmlElement("w:br")); e.append(salto)
    for r in list(nxt.findall(W + "r")):
        e.append(r)
    body.remove(nxt)


def parrafo_campo(instr, aviso):
    par = OxmlElement("w:p")
    for tipo, texto in (("begin", None), ("instr", instr), ("separate", None), ("texto", aviso), ("end", None)):
        par.append(run_campo(tipo, instr if tipo == "instr" else None, texto, negrita=False))
    return par


def titulo_aparte(texto):
    par = OxmlElement("w:p")
    p_ = Paragraph(par, doc)
    r = p_.add_run(texto); r.bold = True
    pf = p_.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER; pf.first_line_indent = Cm(0); pf.line_spacing = 2.0; pf.page_break_before = True
    return par


# portada | tabla de contenido y listas generadas por Word | Resumen
resumen_h = next(e for e, (t, n) in heads.items() if t == "Resumen")
if portada_fin is not None:
    actual = portada_fin
    while actual is not None and actual is not resumen_h:      # elimina la tabla de contenido y las listas convertidas del PDF
        sig = actual.getnext()
        body.remove(actual)
        actual = sig
for elemento in (titulo_aparte("Tabla de contenido"),
                 parrafo_campo(' TOC \\o "1-3" \\h \\z \\u ', "Actualizar campo para generar la tabla de contenido."),
                 titulo_aparte("Lista de tablas"),
                 parrafo_campo(' TOC \\h \\z \\c "Tabla" ', "Actualizar campo para generar la lista de tablas."),
                 titulo_aparte("Lista de figuras"),
                 parrafo_campo(' TOC \\h \\z \\c "Figura" ', "Actualizar campo para generar la lista de figuras.")):
    resumen_h.addprevious(elemento)
Paragraph(resumen_h, doc).paragraph_format.page_break_before = True

doc.save(SALIDA)
print("Documento integrado:", SALIDA)
if sin_resolver:
    print("\nCitas por revisar a mano:")
    for item in sin_resolver:
        print("  ", item)
