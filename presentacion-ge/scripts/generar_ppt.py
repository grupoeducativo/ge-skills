#!/usr/bin/env python3
"""
generar_ppt.py — Genera una presentación .pptx de Grupo Educativo a partir de
un guion JSON, usando los layouts de la plantilla oficial (Plantilla GE v0.6.potx).

Uso:
    python3 generar_ppt.py guion.json salida.pptx [--plantilla RUTA.potx]

El script:
  1. Abre la plantilla .potx (parchea el content-type en memoria).
  2. Elimina las láminas de instrucciones/changelog que trae la plantilla.
  3. Crea cada lámina con el layout oficial que corresponde a su "tipo".
  4. Dibuja los componentes (tarjetas, cifras, proceso, tabla…) con la paleta
     del Brandbook 2026 y Red Hat Display.
  5. Imprime advertencias de densidad (texto que probablemente no cabe).

Ver SKILL.md para el esquema completo del guion.
"""
import copy
import io
import json
import math
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- marca GE
C = {
    "teal": "005062", "magenta": "C0006E", "verde": "BBC100", "verde_texto": "8B9000",
    "slate": "414E55", "azul": "30486B", "violeta": "60306C",
    "purpura": "90186D", "frambuesa": "AB2F52", "cafe": "965F37", "oliva": "818F1B",
    "verde_claro": "90B86D", "verde_vivo": "51A71A", "verde_medio": "368F35", "verde_bosque": "18774F",
    "fondo": "F5F6F7", "blanco": "FFFFFF", "linea": "E1E6E9", "fila": "EEF1F3",
    "slate_suave": "6B7A82",
}
# Design System GE: el verde es detalle, nunca texto ni fondo de números -> fuera de la rotación
ACENTOS = ["teal", "magenta", "azul", "violeta"]
# Líneas de acción (Design System GE · Iconografía)
LINEAS = {"infancias": "magenta", "tp": "verde_vivo", "tecnico profesional": "verde_vivo",
          "formacion continua": "cafe", "mejoramiento": "violeta", "mejoramiento educativo": "violeta"}
FUENTE = "Red Hat Display"

LAYOUT = {
    "portada": "Título imágenes incrustadas",
    "portada_fotos": "Título imágenes personalizadas",
    "portada_teal": "Título (teal)",
    "seccion_blanco": "Sección (blanco)",
    "seccion_teal": "Sección (teal)",
    "contenido_blanco": "Un elemento (blanco)",
    "contenido_teal": "Un elemento (teal)",
    "dos_columnas": "Dos elementos (blanco)",
    "solo_titulo": "Sólo título (blanco)",
    "vacia": "Vacía sin logo GE (blanco)",
    "cierre_blanco": "Cierre (blanco)",
    "cierre_teal": "Cierre (teal)",
}

# Área útil bajo el título en layouts de contenido (pulgadas)
X0, Y0, W0, H0 = 0.92, 2.0, 11.5, 4.7
AVISOS = []


def rgb(nombre):
    return RGBColor.from_string(C.get(nombre, nombre))


def color_texto(nombre):
    """El verde institucional no tiene contraste para texto: usar su sombra."""
    return "verde_texto" if nombre == "verde" else nombre


def acento(d, i):
    """Color de un elemento: el explícito, el de su línea de acción, o la rotación."""
    if d.get("linea"):
        return LINEAS.get(d["linea"].lower(), "teal")
    return d.get("color", ACENTOS[i % len(ACENTOS)])


# ---------------------------------------------------------------- plantilla
def abrir_plantilla(ruta):
    data = Path(ruta).read_bytes()
    zin = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            b = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                b = b.replace(b"presentationml.template.main+xml",
                              b"presentationml.presentation.main+xml")
            zout.writestr(item, b)
    buf.seek(0)
    prs = Presentation(buf)
    # borrar láminas de instrucciones de la plantilla
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    return prs


def layout(prs, clave):
    nombre = LAYOUT[clave]
    for l in prs.slide_layouts:
        if l.name == nombre:
            return l
    raise KeyError(f"La plantilla no trae el layout '{nombre}'")


def ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    return None


def quitar_ph_vacios(slide):
    for p in list(slide.placeholders):
        if p.has_text_frame and not p.text_frame.text.strip():
            p._element.getparent().remove(p._element)


def numero_lamina(slide):
    """Copia el marcador de número de diapositiva del layout (idx 12)."""
    for sp in slide.slide_layout.placeholders:
        if sp.placeholder_format.idx == 12:
            el = copy.deepcopy(sp._element)
            txBody = el.find(qn("p:txBody"))
            if txBody is not None:
                for p in txBody.findall(qn("a:p")):
                    txBody.remove(p)
                p = etree.SubElement(txBody, qn("a:p"))
                fld = etree.SubElement(p, qn("a:fld"), id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}", type="slidenum")
                etree.SubElement(fld, qn("a:rPr"), lang="es-CL")
                t = etree.SubElement(fld, qn("a:t"))
                t.text = "‹#›"
            slide.shapes._spTree.append(el)
            return


# ---------------------------------------------------------------- texto
def escribir(tf, texto, size=18, color="slate", bold=False, italic=False,
             align=None, fuente=FUENTE, nuevo=False, espacio=None):
    """Escribe texto en un text_frame. Admite **negrita** y ==destacado== inline."""
    p = tf.add_paragraph() if nuevo else tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    if espacio:
        p.space_after = Pt(espacio)
    for trozo, estilo in partir_inline(texto):
        r = p.add_run()
        r.text = trozo
        f = r.font
        f.name = fuente
        f.size = Pt(size)
        f.bold = bold or estilo == "b"
        f.italic = italic
        f.color.rgb = rgb("magenta" if estilo == "h" else color)
        if estilo == "h":
            f.bold = True
    return p


def partir_inline(texto):
    """'Hola **mundo** y ==esto==' -> [(Hola ,''),(mundo,'b'),(' y ',''),(esto,'h')]"""
    out, i, buf, modo = [], 0, "", ""
    while i < len(texto):
        if texto.startswith("**", i) and modo in ("", "b"):
            if buf: out.append((buf, modo))
            buf, modo = "", ("" if modo == "b" else "b"); i += 2; continue
        if texto.startswith("==", i) and modo in ("", "h"):
            if buf: out.append((buf, modo))
            buf, modo = "", ("" if modo == "h" else "h"); i += 2; continue
        buf += texto[i]; i += 1
    if buf: out.append((buf, modo))
    return out or [("", "")]


def texto_plano(t):
    return t.replace("**", "").replace("==", "")


def caja(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, margen=0.0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(margen))
    return tb, tf


def lineas_estimadas(texto, size_pt, ancho_in):
    chars_por_linea = max(1, (ancho_in * 72) / (size_pt * 0.52))
    return sum(max(1, math.ceil(len(texto_plano(par)) / chars_por_linea)) for par in texto.split("\n"))


def avisar_si_no_cabe(n, texto_total_lineas, size, alto_in, donde):
    alto = texto_total_lineas * size * 1.25 / 72
    if alto > alto_in:
        AVISOS.append(f"Lámina {n} ({donde}): el texto ocupa ~{alto:.1f}\" y hay {alto_in:.1f}\". Recorta o divide la lámina.")


def rect(slide, x, y, w, h, fill, redondeado=True, borde=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if redondeado else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if redondeado:
        shp.adjustments[0] = min(0.12, 0.12 / max(w, h) * 6)
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    if borde:
        shp.line.color.rgb = rgb(borde); shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    shp.text_frame.text = ""
    return shp


def cuadro(slide, x, y, d, fill, texto="", size=16, color="blanco"):
    """Marcador numérico cuadrado (el DS GE usa esquinas rectas)."""
    return circulo(slide, x, y, d, fill, texto, size, color, forma=MSO_SHAPE.RECTANGLE)


def circulo(slide, x, y, d, fill, texto="", size=16, color="blanco", forma=MSO_SHAPE.OVAL):
    shp = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background(); shp.shadow.inherit = False
    tf = shp.text_frame
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if texto:
        escribir(tf, texto, size=size, color=color, bold=True, align=PP_ALIGN.CENTER)
    return shp


# ---------------------------------------------------------------- láminas
def set_titulo(slide, texto):
    """Título en el placeholder del layout, heredando tamaño y color de la plantilla.
    ==x== se pinta magenta."""
    t = ph(slide, 0)
    if t is None:
        return
    tf = t.text_frame
    tf.text = ""
    p = tf.paragraphs[0]
    for trozo, estilo in partir_inline(texto):
        r = p.add_run(); r.text = trozo
        if estilo == "h":
            r.font.color.rgb = rgb("magenta")


def set_sub(slide, texto):
    s = ph(slide, 1)
    if s is not None and texto:
        s.text_frame.text = texto_plano(texto)


def lista(tf, puntos, size, color="slate", nivel=0, primero=[True]):
    """Viñetas usando los niveles del placeholder (hereda viñeta de la plantilla)."""
    for item in puntos:
        if isinstance(item, list):
            lista(tf, item, size - 2, color, nivel + 1, primero)
            continue
        p = tf.paragraphs[0] if primero[0] else tf.add_paragraph()
        primero[0] = False
        p.level = nivel
        for trozo, estilo in partir_inline(item):
            r = p.add_run(); r.text = trozo
            r.font.size = Pt(size)
            if estilo == "b": r.font.bold = True
            if estilo == "h": r.font.bold = True; r.font.color.rgb = rgb("magenta")


def contar(puntos):
    return [x for it in puntos for x in (contar(it) if isinstance(it, list) else [it])]


def l_portada(prs, d, n):
    clave = {"fotos": "portada_fotos", "teal": "portada_teal"}.get(d.get("variante"), "portada")
    s = prs.slides.add_slide(layout(prs, clave))
    set_titulo(s, d["titulo"]); set_sub(s, d.get("subtitulo", ""))
    # anclar el título arriba para que un título de 2 líneas no suba sobre las fotos
    largo = len(texto_plano(d["titulo"]))
    if largo > 24:  # a 60pt no cabe en 1 línea: bajar a 40pt para que no suba sobre las fotos
        t = ph(s, 0)
        t.left, t.top, t.width, t.height = Inches(0.92), Inches(4.15), Inches(11.5), Inches(1.5)
        t.text_frame._txBody.find(qn("a:bodyPr")).set("anchor", "b")
        for r in t.text_frame.paragraphs[0].runs:
            r.font.size = Pt(40)
    if largo > 60:
        AVISOS.append(f"Lámina {n}: título de portada largo ({largo} car.). Máximo ~60 (2 líneas); pasa el resto al subtítulo.")
    if clave != "portada":
        for idx, img in zip((10, 11, 12), d.get("imagenes", [])):
            p = ph(s, idx)
            if p is not None and Path(img).exists():
                p.insert_picture(img)
    quitar_ph_vacios(s)
    return s


def l_seccion(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "seccion_blanco" if d.get("fondo") == "blanco" else "seccion_teal"))
    set_titulo(s, d["titulo"]); set_sub(s, d.get("subtitulo", ""))
    quitar_ph_vacios(s)
    return s


def l_contenido(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "contenido_teal" if d.get("fondo") == "teal" else "contenido_blanco"))
    set_titulo(s, d["titulo"])
    body = ph(s, 1)
    tf = body.text_frame; tf.text = ""
    size = d.get("tamano", 20)
    if d.get("intro"):
        p = tf.paragraphs[0]
        for trozo, estilo in partir_inline(d["intro"]):
            r = p.add_run(); r.text = trozo; r.font.size = Pt(size)
            if estilo in ("b", "h"): r.font.bold = True
        pPr = p._p.get_or_add_pPr(); pPr.set("marL", "0"); pPr.set("indent", "0")
        etree.SubElement(pPr, qn("a:buNone"))
        p.space_after = Pt(12)
        lista(tf, d.get("puntos", []), size, primero=[False])
    else:
        lista(tf, d.get("puntos", []), size, primero=[True])
    items = contar(d.get("puntos", [])) + ([d["intro"]] if d.get("intro") else [])
    if len(items) > 7:
        AVISOS.append(f"Lámina {n}: {len(items)} viñetas. Máximo recomendado 6; divide o usa tarjetas.")
    avisar_si_no_cabe(n, sum(lineas_estimadas(t, size, 11.2) for t in items) + len(items) * 0.3, size, H0, "contenido")
    numero_lamina(s)
    return s


def l_dos_columnas(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "dos_columnas"))
    set_titulo(s, d["titulo"])
    size = d.get("tamano", 18)
    for idx, lado in ((1, "izquierda"), (2, "derecha")):
        col = d.get(lado, {})
        tf = ph(s, idx).text_frame; tf.text = ""
        prim = [True]
        if col.get("subtitulo"):
            p = tf.paragraphs[0]; prim = [False]
            r = p.add_run(); r.text = texto_plano(col["subtitulo"])
            r.font.size = Pt(size + 2); r.font.bold = True
            r.font.color.rgb = rgb(color_texto(col.get("color", "teal" if idx == 1 else "magenta")))
            pPr = p._p.get_or_add_pPr(); pPr.set("marL", "0"); pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone")); p.space_after = Pt(8)
        lista(tf, col.get("puntos", []), size, primero=prim)
        items = contar(col.get("puntos", []))
        avisar_si_no_cabe(n, sum(lineas_estimadas(t, size, 5.4) for t in items) + 1.5, size, H0, lado)
    numero_lamina(s)
    return s


def l_tarjetas(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d["titulo"])
    tarjetas = d["tarjetas"]
    k = len(tarjetas)
    if k > 4:
        cols = 3 if k in (5, 6) else 4
    else:
        cols = k
    filas = math.ceil(k / cols)
    y = Y0 + (0.5 if d.get("intro") else 0)
    if d.get("intro"):
        _, tf = caja(s, X0, Y0 - 0.1, W0, 0.5)
        escribir(tf, d["intro"], size=18, color="slate_suave")
    gap = 0.3
    w = (W0 - gap * (cols - 1)) / cols
    h_total = Y0 + H0 - y
    h = (h_total - gap * (filas - 1)) / filas
    for i, t in enumerate(tarjetas):
        c = acento(t, i)
        x = X0 + (i % cols) * (w + gap)
        yy = y + (i // cols) * (h + gap)
        rect(s, x, yy, w, h, "blanco", redondeado=False, borde="linea")
        rect(s, x, yy, 0.08, h, c, redondeado=False)  # barra izquierda de acento (DS GE)
        pad = 0.32
        ty = yy + 0.3
        if t.get("numero") or d.get("numerar"):
            cuadro(s, x + pad, ty, 0.46, c, str(t.get("numero", f"{i+1:02d}")), size=13)
            ty += 0.6
        _, tf = caja(s, x + pad, ty, w - 2 * pad, h - (ty - yy) - 0.2)
        escribir(tf, t["titulo"], size=18 if cols >= 4 else 20, color=color_texto(c), bold=True, espacio=6)
        if t.get("texto"):
            escribir(tf, t["texto"], size=15 if cols >= 4 else 17, color="slate", nuevo=True)
        tl = lineas_estimadas(t["titulo"], 20, w - 2 * pad) + lineas_estimadas(t.get("texto", ""), 17, w - 2 * pad)
        avisar_si_no_cabe(n, tl, 18, h - (ty - yy) - 0.2, f"tarjeta {i+1}")
    numero_lamina(s)
    return s


def l_cifras(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d["titulo"])
    cifras = d["cifras"]; k = len(cifras)
    gap = 0.4; w = (W0 - gap * (k - 1)) / k
    y = Y0 + 0.5
    for i, c in enumerate(cifras):
        col = acento(c, i)
        x = X0 + i * (w + gap)
        rect(s, x, y, 0.08, 2.6, col, redondeado=False)
        _, tf = caja(s, x + 0.3, y - 0.1, w - 0.3, 1.4, anchor=MSO_ANCHOR.BOTTOM)
        escribir(tf, c["valor"], size=66 if k <= 3 else 54, color=color_texto(col), bold=True)
        _, tf = caja(s, x + 0.3, y + 1.35, w - 0.3, 1.3)
        escribir(tf, c["etiqueta"], size=18, color="slate", bold=True, espacio=4)
        if c.get("detalle"):
            escribir(tf, c["detalle"], size=14, color="slate_suave", nuevo=True)
    if d.get("nota"):
        _, tf = caja(s, X0, Y0 + H0 - 0.9, W0, 0.8, anchor=MSO_ANCHOR.BOTTOM)
        escribir(tf, d["nota"], size=16, color="slate_suave", italic=True)
    numero_lamina(s)
    return s


def l_proceso(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d["titulo"])
    pasos = d["pasos"]; k = len(pasos)
    gap = 0.25; w = (W0 - gap * (k - 1)) / k
    y = Y0 + 0.7
    # línea conectora
    ln = s.shapes.add_connector(1, Inches(X0 + w / 2), Inches(y + 0.45), Inches(X0 + W0 - w / 2), Inches(y + 0.45))
    ln.line.color.rgb = rgb("linea"); ln.line.width = Pt(3)
    for i, p in enumerate(pasos):
        col = acento(p, i)
        x = X0 + i * (w + gap)
        cuadro(s, x + w / 2 - 0.45, y, 0.9, col, str(i + 1), size=24)
        _, tf = caja(s, x, y + 1.2, w, 3.2)
        escribir(tf, p["titulo"], size=22, color=color_texto(col), bold=True, align=PP_ALIGN.CENTER, espacio=6)
        if p.get("texto"):
            escribir(tf, p["texto"], size=18 if k <= 4 else 16, color="slate", align=PP_ALIGN.CENTER, nuevo=True)
        if p.get("cuando"):
            escribir(tf, p["cuando"], size=14, color="slate_suave", bold=True, align=PP_ALIGN.CENTER, nuevo=True)
        tl = lineas_estimadas(p["titulo"], 22, w) + lineas_estimadas(p.get("texto", ""), 18, w)
        avisar_si_no_cabe(n, tl, 19, 3.2, f"paso {i+1}")
    numero_lamina(s)
    return s


def l_tabla(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d["titulo"])
    enc, filas = d["encabezados"], d["filas"]
    nf, nc = len(filas) + 1, len(enc)
    size = d.get("tamano", 14 if nf > 10 else (16 if nf > 8 else 18))
    alto_fila = size * 2.1 / 72
    shp = s.shapes.add_table(nf, nc, Inches(X0), Inches(Y0), Inches(W0), Inches(alto_fila * nf))
    tbl = shp.table
    # quitar estilo de tabla por defecto
    tblPr = tbl._tbl.tblPr
    for a in ("bandRow", "firstRow"):
        tblPr.set(a, "0")
    anchos = d.get("anchos")
    if anchos:
        tot = sum(anchos)
        for j, a in enumerate(anchos):
            tbl.columns[j].width = Inches(W0 * a / tot)
    for i in range(nf):
        for j in range(nc):
            cell = tbl.cell(i, j)
            val = enc[j] if i == 0 else str(filas[i - 1][j])
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb("teal" if i == 0 else ("blanco" if i % 2 else "fila"))
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.text = ""
            escribir(tf, val, size=size, color="blanco" if i == 0 else "slate", bold=(i == 0 or (j == 0 and d.get("primera_col_negrita", True))),
                     align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    if alto_fila * nf > H0:
        AVISOS.append(f"Lámina {n}: la tabla ({nf} filas) probablemente no cabe; máximo ~{int(H0/alto_fila)-1} filas a {size}pt.")
    if d.get("fuente"):
        _, tf = caja(s, X0, 6.45, W0, 0.35)
        escribir(tf, "Fuente: " + d["fuente"], size=11, color="slate_suave", italic=True)
    numero_lamina(s)
    return s


def l_pregunta(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "vacia"))
    _, tf = caja(s, 1.6, 1.2, 10.1, 5.1, anchor=MSO_ANCHOR.MIDDLE)
    escribir(tf, d["texto"], size=d.get("tamano", 40), color="slate", bold=True, align=PP_ALIGN.CENTER, espacio=24)
    if d.get("bajada"):
        escribir(tf, d["bajada"], size=20, color="slate_suave", align=PP_ALIGN.CENTER, nuevo=True)
    avisar_si_no_cabe(n, lineas_estimadas(d["texto"], d.get("tamano", 40), 10.1), d.get("tamano", 40), 5.1, "pregunta")
    quitar_ph_vacios(s)
    return s


def l_cita(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d.get("titulo", ""))
    rect(s, X0, Y0 + 0.1, 0.1, 3.6, "magenta", redondeado=False)
    _, tf = caja(s, X0 + 0.45, Y0, W0 - 0.6, 4.0, anchor=MSO_ANCHOR.MIDDLE)
    escribir(tf, "“" + d["texto"] + "”", size=d.get("tamano", 26), color="slate", italic=True, espacio=14)
    if d.get("fuente"):
        escribir(tf, "— " + d["fuente"], size=16, color="teal", bold=True, nuevo=True)
    avisar_si_no_cabe(n, lineas_estimadas(d["texto"], d.get("tamano", 26), W0 - 0.6) + 1, d.get("tamano", 26), 4.0, "cita")
    if not d.get("titulo"):
        quitar_ph_vacios(s)
    numero_lamina(s)
    return s


def l_agenda(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d.get("titulo", "Agenda"))
    items = d["items"]; k = len(items)
    alto = min(0.85, H0 / k)
    for i, it in enumerate(items):
        y = Y0 + i * alto
        col = ACENTOS[i % len(ACENTOS)]
        cuadro(s, X0, y + (alto - 0.55) / 2, 0.55, col, str(i + 1), size=16)
        _, tf = caja(s, X0 + 0.85, y, 8.8, alto, anchor=MSO_ANCHOR.MIDDLE)
        escribir(tf, it["texto"], size=22, color="slate")
        if it.get("min"):
            _, tf = caja(s, X0 + 9.8, y, 1.7, alto, anchor=MSO_ANCHOR.MIDDLE)
            escribir(tf, f"{it['min']}’", size=20, color="slate_suave", bold=True, align=PP_ALIGN.RIGHT)
    numero_lamina(s)
    return s


def l_imagen(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "solo_titulo"))
    set_titulo(s, d["titulo"])
    img = d.get("imagen")
    alto = H0 - (0.45 if d.get("pie") else 0)
    if img and Path(img).exists():
        from PIL import Image
        iw, ih = Image.open(img).size
        esc = min(W0 / iw, alto / ih)
        w, h = iw * esc, ih * esc
        s.shapes.add_picture(img, Inches(X0 + (W0 - w) / 2), Inches(Y0), Inches(w), Inches(h))
    else:
        r = rect(s, X0, Y0, W0, alto, "fila", redondeado=False, borde="linea")
        escribir(r.text_frame, f"[Imagen: {d.get('imagen', 'por definir')}]", size=18, color="slate_suave", align=PP_ALIGN.CENTER)
        AVISOS.append(f"Lámina {n}: imagen no encontrada, quedó un marcador.")
    if d.get("pie"):
        _, tf = caja(s, X0, Y0 + alto + 0.1, W0, 0.35)
        escribir(tf, d["pie"], size=12, color="slate_suave", italic=True, align=PP_ALIGN.CENTER)
    numero_lamina(s)
    return s


def l_cierre(prs, d, n):
    s = prs.slides.add_slide(layout(prs, "cierre_blanco" if d.get("fondo") == "blanco" else "cierre_teal"))
    if d.get("texto"):
        oscuro = d.get("fondo") != "blanco"
        _, tf = caja(s, 0.92, 5.9, 11.5, 0.6, anchor=MSO_ANCHOR.MIDDLE)
        escribir(tf, d["texto"], size=18, color="blanco" if oscuro else "slate", align=PP_ALIGN.CENTER)
    quitar_ph_vacios(s)
    return s


TIPOS = {
    "portada": l_portada, "seccion": l_seccion, "contenido": l_contenido,
    "dos_columnas": l_dos_columnas, "tarjetas": l_tarjetas, "cifras": l_cifras,
    "proceso": l_proceso, "tabla": l_tabla, "pregunta": l_pregunta, "cita": l_cita,
    "agenda": l_agenda, "imagen": l_imagen, "cierre": l_cierre,
}


def construir(guion, salida, plantilla):
    prs = abrir_plantilla(plantilla)
    prs.core_properties.title = guion.get("titulo", "")
    prs.core_properties.author = guion.get("autor", "Grupo Educativo")
    for n, d in enumerate(guion["laminas"], 1):
        tipo = d.get("tipo")
        if tipo not in TIPOS:
            raise ValueError(f"Lámina {n}: tipo '{tipo}' no existe. Opciones: {', '.join(TIPOS)}")
        s = TIPOS[tipo](prs, d, n)
        if d.get("notas"):
            s.notes_slide.notes_text_frame.text = d["notas"]
    prs.save(salida)
    return len(guion["laminas"])


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__); sys.exit(1)
    plantilla = None
    if "--plantilla" in args:
        i = args.index("--plantilla"); plantilla = args[i + 1]; del args[i:i + 2]
    if plantilla is None:
        base = Path(__file__).resolve().parent.parent
        cand = sorted(base.glob("recursos/*.potx")) or sorted(base.glob("*.potx"))
        if not cand:
            sys.exit("No encuentro la plantilla .potx junto a la skill; usa --plantilla.")
        plantilla = cand[0]
    guion = json.loads(Path(args[0]).read_text(encoding="utf-8"))
    n = construir(guion, args[1], plantilla)
    print(f"OK: {n} láminas → {args[1]}")
    if AVISOS:
        print("\nAVISOS DE DENSIDAD (revisar antes de entregar):")
        for a in AVISOS:
            print("  -", a)


if __name__ == "__main__":
    main()
