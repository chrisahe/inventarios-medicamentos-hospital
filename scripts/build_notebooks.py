"""Convierte los scripts en formato de celdas (# %%) en cuadernos .ipynb para Google Colab."""
import glob, os, re
import nbformat as nbf

ARRANQUE = """# Colab: descomentar, reemplazar la URL por la del repositorio del proyecto y ejecutar una vez.
# !git clone https://github.com/USUARIO/REPOSITORIO.git
# %cd REPOSITORIO
# !pip install -q -r requirements.txt"""

for ruta in sorted(glob.glob(os.path.join(os.path.dirname(__file__), "0*.py"))):
    celdas, tipo, buf = [], None, []
    def cerrar():
        global buf
        txt = "\n".join(buf).strip("\n")
        if txt:
            if tipo == "md":
                txt = "\n".join(re.sub(r"^# ?", "", l) for l in txt.splitlines())
                celdas.append(nbf.v4.new_markdown_cell(txt))
            else:
                celdas.append(nbf.v4.new_code_cell(txt))
        buf = []
    for linea in open(ruta, encoding="utf-8").read().splitlines():
        if linea.startswith("# %%"):
            cerrar(); tipo = "md" if "[markdown]" in linea else "code"
        else:
            buf.append(linea)
    cerrar()
    celdas.insert(1, nbf.v4.new_code_cell(ARRANQUE))
    nb = nbf.v4.new_notebook(cells=celdas)
    destino = os.path.join(os.path.dirname(__file__), "..", "notebooks", os.path.basename(ruta).replace(".py", ".ipynb"))
    nbf.write(nb, destino)
    print("->", os.path.normpath(destino))
