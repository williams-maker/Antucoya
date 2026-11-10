"""Lámina «Resumen ejecutivo Assurance»: banda de título, 4 indicadores y tabla nativa por foco.

agregar_resumen(pptx_bytes, datos) -> (bytes, log)
datos = {titulo, subtitulo, kpis:[{titulo,valor,color}], filas:[{foco,ambito,frente,ito,herramienta,evidencia,estado,estado_color}], pie}
"""
import io
import math

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

MARINO = RGBColor(0x14, 0x23, 0x4B)
TEXTO = RGBColor(0x22, 0x22, 0x22)
BORDE = RGBColor(0xD5, 0xD9, 0xDE)
KPI = {"verde": RGBColor(0x2E, 0x6B, 0x2E), "azul": RGBColor(0x1F, 0x4E, 0x8C),
       "ambar": RGBColor(0xB8, 0x6A, 0x12), "negro": RGBColor(0x22, 0x22, 0x22)}
EST = {"verde": RGBColor(0x2E, 0x6B, 0x2E), "ambar": RGBColor(0xD9, 0x9A, 0x00),
       "rojo": RGBColor(0xC0, 0x1F, 0x1F), "gris": RGBColor(0x55, 0x55, 0x55)}
ANCHOS = (0.19, 0.20, 0.17, 0.33, 0.11)
TITULOS = ("Foco Assurance / Ámbito", "Frente / Empresa · ITO", "Herramienta SGE",
           "Evidencia de cumplimiento / Acción en terreno", "Estado")
NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _s(v):
    return "" if v is None else str(v).strip()


def _lineas(texto, ancho_in, pt):
    """Líneas estimadas que ocupa un texto (varios párrafos) en una columna de ancho_in pulgadas."""
    por_linea = max(6, int(ancho_in * 72 / (pt * 0.5)))
    return sum(max(1, math.ceil(len(p) / por_linea)) for p in (texto.split("\n") if texto else [""]))


def _bordes(celda):
    tcPr = celda._tc.get_or_add_tcPr()
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        for viejo in tcPr.findall(f"{{{NS}}}{tag}"):
            tcPr.remove(viejo)
    # Los bordes deben ir antes del relleno dentro de tcPr
    for i, tag in enumerate(("lnL", "lnR", "lnT", "lnB")):
        ln = etree.Element(f"{{{NS}}}{tag}", w="9525", cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, f"{{{NS}}}solidFill")
        etree.SubElement(sf, f"{{{NS}}}srgbClr", val="000000")
        etree.SubElement(ln, f"{{{NS}}}prstDash", val="solid")
        tcPr.insert(i, ln)


def _parrafos(celda, partes, pt, alin=PP_ALIGN.LEFT):
    """partes = [(texto, negrita, color)]; cada elemento es un párrafo."""
    tf = celda.text_frame
    tf.word_wrap = True
    primero = True
    for texto, negrita, color in partes:
        for linea in (texto.split("\n") if texto else [""]):
            p = tf.paragraphs[0] if primero else tf.add_paragraph()
            primero = False
            p.alignment = alin
            r = p.add_run(); r.text = linea
            r.font.name = "Calibri"; r.font.size = Pt(pt); r.font.bold = negrita; r.font.color.rgb = color


def _celda(celda, relleno, partes, pt, alin=PP_ALIGN.LEFT):
    celda.fill.solid(); celda.fill.fore_color.rgb = relleno
    celda.margin_left = celda.margin_right = Inches(0.06)
    celda.margin_top = celda.margin_bottom = Inches(0.04)
    celda.vertical_anchor = MSO_ANCHOR.MIDDLE
    _parrafos(celda, partes, pt, alin)
    _bordes(celda)


def agregar_resumen(pptx_bytes, datos):
    datos = datos if isinstance(datos, dict) else {}
    filas = [f for f in (datos.get("filas") or []) if isinstance(f, dict)]
    if not filas:
        return pptx_bytes, ["Resumen ejecutivo: sin filas, no se agregó la lámina."]
    log = []
    prs = Presentation(io.BytesIO(pptx_bytes))
    W = prs.slide_width / 914400
    H = prs.slide_height / 914400
    layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    sl = prs.slides.add_slide(layout)
    for sh in list(sl.shapes):
        if sh.is_placeholder:
            sh._element.getparent().remove(sh._element)

    # Banda de título
    mx = 0.3
    ancho = W - 2 * mx
    banda = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(mx), Inches(0.25), Inches(ancho), Inches(1.15))
    banda.fill.solid(); banda.fill.fore_color.rgb = MARINO; banda.line.fill.background()
    try: banda.shadow.inherit = False
    except Exception: pass
    tf = banda.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = _s(datos.get("titulo")) or "RESUMEN EJECUTIVO ASSURANCE"
    r.font.name = "Calibri"; r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = _s(datos.get("subtitulo"))
    r2.font.name = "Calibri"; r2.font.size = Pt(13); r2.font.color.rgb = RGBColor(0xDD, 0xE3, 0xEE)

    # Indicadores
    kpis = [k for k in (datos.get("kpis") or []) if isinstance(k, dict)][:4]
    if kpis:
        sep = 0.25
        kw = (ancho - sep * (len(kpis) - 1)) / len(kpis)
        for i, k in enumerate(kpis):
            x = mx + i * (kw + sep)
            c = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.65), Inches(kw), Inches(1.15))
            try: c.adjustments[0] = 0.12
            except Exception: pass
            c.fill.solid(); c.fill.fore_color.rgb = RGBColor(255, 255, 255)
            c.line.color.rgb = BORDE; c.line.width = Pt(0.75)
            tfk = c.text_frame; tfk.word_wrap = True; tfk.vertical_anchor = MSO_ANCHOR.MIDDLE
            pk = tfk.paragraphs[0]; pk.alignment = PP_ALIGN.CENTER
            rk = pk.add_run(); rk.text = _s(k.get("titulo")).upper()
            rk.font.name = "Calibri"; rk.font.size = Pt(12); rk.font.bold = True; rk.font.color.rgb = RGBColor(0x4A, 0x4A, 0x4A)
            pv = tfk.add_paragraph(); pv.alignment = PP_ALIGN.CENTER
            rv = pv.add_run(); rv.text = _s(k.get("valor"))
            rv.font.name = "Calibri"; rv.font.size = Pt(15); rv.font.bold = True
            rv.font.color.rgb = KPI.get(_s(k.get("color")).lower(), KPI["negro"])

    # Tabla: tamaño de letra adaptable al alto disponible
    top = 3.05
    alto_disp = H - top - 0.75
    cols_in = [ancho * a for a in ANCHOS]
    textos = []
    for f in filas:
        textos.append((
            _s(f.get("foco")) + "\n" + _s(f.get("ambito")),
            _s(f.get("frente")) + "\n" + _s(f.get("ito")),
            _s(f.get("herramienta")), _s(f.get("evidencia")), _s(f.get("estado"))))

    def alturas(pt):
        return [max(0.45, max(_lineas(t, cols_in[j] - 0.12, pt) for j, t in enumerate(fila)) * pt * 1.2 / 72 + 0.12) for fila in textos]

    pt = 11
    for pt in (11, 10, 9, 8):
        hs = alturas(pt)
        if sum(hs) + 0.42 <= alto_disp:
            break
    hs = alturas(pt)
    if pt < 10:
        log.append(f"⚠ Resumen ejecutivo: muchas filas, letra reducida a {pt} pt")
    if sum(hs) + 0.42 > alto_disp:
        log.append("⚠ Resumen ejecutivo: la tabla excede el alto de la lámina; reduce el texto o quita filas")

    gf = sl.shapes.add_table(len(filas) + 1, 5, Inches(mx), Inches(top), Inches(ancho), Inches(0.42 + sum(hs)))
    tbl = gf.table
    tbl.first_row = False; tbl.horz_banding = False
    for j, a in enumerate(ANCHOS):
        tbl.columns[j].width = Inches(ancho * a)
    tbl.rows[0].height = Inches(0.42)
    for j, t in enumerate(TITULOS):
        _celda(tbl.cell(0, j), MARINO, [(t, True, RGBColor(255, 255, 255))], 11, PP_ALIGN.CENTER if j == 4 else PP_ALIGN.LEFT)
    for i, f in enumerate(filas, 1):
        tbl.rows[i].height = Inches(hs[i - 1])
        fondo = RGBColor(255, 255, 255) if i % 2 else RGBColor(0xF2, 0xF5, 0xF9)
        _celda(tbl.cell(i, 0), fondo, [(_s(f.get("foco")), True, TEXTO), (_s(f.get("ambito")), True, MARINO)], pt)
        _celda(tbl.cell(i, 1), fondo, [(_s(f.get("frente")), False, TEXTO), (_s(f.get("ito")), False, TEXTO)], pt)
        _celda(tbl.cell(i, 2), fondo, [(_s(f.get("herramienta")), False, TEXTO)], pt)
        _celda(tbl.cell(i, 3), fondo, [(_s(f.get("evidencia")), False, TEXTO)], pt)
        _celda(tbl.cell(i, 4), fondo, [(_s(f.get("estado")), True, EST.get(_s(f.get("estado_color")).lower(), EST["gris"]))], pt, PP_ALIGN.CENTER)

    pie = _s(datos.get("pie"))
    if pie:
        tb = sl.shapes.add_textbox(Inches(mx), Inches(H - 0.5), Inches(ancho), Inches(0.3))
        pr = tb.text_frame.paragraphs[0]
        rp = pr.add_run(); rp.text = pie
        rp.font.name = "Calibri"; rp.font.size = Pt(9); rp.font.color.rgb = RGBColor(0x6B, 0x75, 0x80)

    out = io.BytesIO(); prs.save(out)
    log.insert(0, f"✔ Agregada lámina «Resumen ejecutivo Assurance» ({len(filas)} filas, {len(kpis)} indicadores)")
    return out.getvalue(), log
