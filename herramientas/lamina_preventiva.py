"""Láminas «Gestión preventiva ejecutada en terreno» (se arman desde el informe JEJ en la página).

agregar_preventiva(pptx_bytes, datos) -> (bytes, log). Hasta 4 bloques por lámina (cuadrícula 2x2).
datos = {lema, titulo, subtitulo, bloques:[{titulo, intro, items:[...], fotos:[base64,...]}]}
"""
import io
import re
import base64

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from PIL import Image

TINTA = RGBColor(0x1F, 0x2D, 0x3D)
GRIS = RGBColor(0x5F, 0x6B, 0x73)
TEAL = RGBColor(0x0E, 0x7C, 0x86)
BORDE = RGBColor(0xDA, 0xDF, 0xE3)
NEGRO = RGBColor(0x22, 0x22, 0x22)
BLANCO = RGBColor(255, 255, 255)


def _s(v):
    return "" if v is None else str(v).strip()


def _sin_sombra(sh):
    try:
        sh.shadow.inherit = False
    except Exception:
        pass


def _rect(slide, x, y, w, h):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        sh.adjustments[0] = 0.05
    except Exception:
        pass
    sh.fill.solid(); sh.fill.fore_color.rgb = BLANCO
    sh.line.color.rgb = BORDE; sh.line.width = Pt(0.75)
    _sin_sombra(sh)
    return sh


def _caja(slide, x, y, w, h, anclaje=MSO_ANCHOR.TOP):
    sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anclaje
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    return sh, tf


def _run(p, texto, tam, color, negrita=False):
    r = p.add_run()
    r.text = texto
    r.font.name = "Calibri"; r.font.size = Pt(tam); r.font.bold = negrita; r.font.color.rgb = color
    return r


def _texto(slide, texto, x, y, w, h, tam, color, negrita=False, alin=PP_ALIGN.LEFT):
    _, tf = _caja(slide, x, y, w, h)
    p = tf.paragraphs[0]
    p.alignment = alin
    _run(p, texto, tam, color, negrita)


def _limpiar(items):
    out = []
    for it in items or []:
        t = re.sub(r"^\s*(?:[•●▪■\-–*]|\d+[.)])\s*", "", _s(it)).strip()
        if t:
            out.append(t)
    return out


def _tam_letra(total):
    return 12 if total <= 350 else 11 if total <= 550 else 10 if total <= 800 else 9


def _foto(slide, b64, x, y, w, h):
    """Recorta al centro ('cover') a la proporción exacta de la celda y la inserta."""
    im = Image.open(io.BytesIO(base64.b64decode(b64.split(",")[-1])))
    im = im.convert("RGB")
    ratio = w / h
    iw, ih = im.size
    if iw / ih > ratio:
        nw = int(ih * ratio); izq = (iw - nw) // 2
        im = im.crop((izq, 0, izq + nw, ih))
    else:
        nh = int(iw / ratio); sup = (ih - nh) // 2
        im = im.crop((0, sup, iw, sup + nh))
    if im.size[0] > 1000:
        im = im.resize((1000, max(1, int(1000 / ratio))))
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=85); buf.seek(0)
    pic = slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))
    pic.line.color.rgb = BORDE; pic.line.width = Pt(0.75)


def _celdas(n, x, y, w, h, sep=0.08):
    if n == 1:
        return [(x, y, w, h)]
    if n == 2:
        hh = (h - sep) / 2
        return [(x, y, w, hh), (x, y + hh + sep, w, hh)]
    if n == 3:
        h1 = (h - sep) * 0.55; h2 = h - sep - h1; ww = (w - sep) / 2
        return [(x, y, w, h1), (x, y + h1 + sep, ww, h2), (x + ww + sep, y + h1 + sep, ww, h2)]
    ww = (w - sep) / 2; hh = (h - sep) / 2
    return [(x, y, ww, hh), (x + ww + sep, y, ww, hh), (x, y + hh + sep, ww, hh), (x + ww + sep, y + hh + sep, ww, hh)]


def _bloque(slide, b, x, y, w, h, log):
    _rect(slide, x, y, w, h)
    fotos = [f for f in (b.get("fotos") or []) if _s(f)]
    if len(fotos) > 4:
        log.append(f"⚠ «{_s(b.get('titulo'))}»: {len(fotos)} fotos, se usaron 4")
        fotos = fotos[:4]
    zona_foto = w * 0.42 if fotos else 0
    ancho_txt = w - 0.5 - (zona_foto + 0.1 if fotos else 0)

    _texto(slide, _s(b.get("titulo")).upper(), x + 0.25, y + 0.15, ancho_txt, 0.35, 12, TEAL, True)
    items = _limpiar(b.get("items"))
    intro = _s(b.get("intro"))
    tam = _tam_letra(len(intro) + sum(len(i) + 2 for i in items))
    _, tf = _caja(slide, x + 0.25, y + 0.55, ancho_txt, h - 0.7)
    primero = True
    if intro:
        p = tf.paragraphs[0]; primero = False
        p.space_after = Pt(3)
        _run(p, intro, tam, NEGRO)
    for it in items:
        p = tf.paragraphs[0] if primero else tf.add_paragraph()
        primero = False
        p.space_after = Pt(3)
        _run(p, "•  ", tam, TEAL, True)
        m = re.match(r"^([^:]{2,60}):\s*(.+)$", it)
        if m:
            _run(p, m.group(1) + ":", tam, NEGRO, True)
            _run(p, " " + m.group(2), tam, NEGRO)
        else:
            _run(p, it, tam, NEGRO)

    if fotos:
        fx = x + w - zona_foto - 0.2
        for b64, (cx, cy, cw, ch) in zip(fotos, _celdas(len(fotos), fx, y + 0.2, zona_foto, h - 0.4)):
            try:
                _foto(slide, b64, cx, cy, cw, ch)
            except Exception as e:
                log.append(f"⚠ Foto no válida en «{_s(b.get('titulo'))}»: {e}")


def agregar_preventiva(pptx_bytes, datos):
    datos = datos if isinstance(datos, dict) else {}
    bloques = [b for b in (datos.get("bloques") or []) if isinstance(b, dict)]
    if not bloques:
        return pptx_bytes, ["Gestión preventiva: sin bloques, no se agregaron láminas."]

    prs = Presentation(io.BytesIO(pptx_bytes))
    W = prs.slide_width / 914400; H = prs.slide_height / 914400
    logo = None
    for sh in prs.slides[0].shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                logo = sh.image.blob
            except Exception:
                pass
            break
    layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    grupos = [bloques[i:i + 4] for i in range(0, len(bloques), 4)]
    log = []
    nfotos = 0
    for gi, grupo in enumerate(grupos, 1):
        slide = prs.slides.add_slide(layout)
        for sh in list(slide.shapes):
            if sh.is_placeholder:
                sh._element.getparent().remove(sh._element)
        for cx, col in zip((0.28, 0.48, 0.68, 0.88), ("F5A300", "E3001B", "5BC8D0", "0E7C86")):
            c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(0.45), Inches(0.15), Inches(0.15))
            c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(col)
            c.line.fill.background(); _sin_sombra(c)
        if logo:
            try:
                slide.shapes.add_picture(io.BytesIO(logo), Inches(W - 1.95), Inches(0.25), Inches(1.8), Inches(0.7))
            except Exception:
                pass
        sub = _s(datos.get("subtitulo"))
        if len(grupos) > 1:
            sub = f"{sub} ({gi}/{len(grupos)})".strip()
        _texto(slide, _s(datos.get("lema")).upper(), 1.2, 0.38, 9.0, 0.3, 12, TEAL, True)
        _texto(slide, _s(datos.get("titulo")), 0.25, 0.7, W - 2.6, 0.6, 30, TINTA, True)
        _texto(slide, sub, 0.3, 1.32, W - 2.6, 0.35, 14, GRIS)

        top, bot, mx, gap = 1.85, H - 0.45, 0.35, 0.25
        n = len(grupo)
        filas = 1 if n <= 2 else 2
        alto = (bot - top - gap * (filas - 1)) / filas
        ancho = (W - 2 * mx - gap) / 2
        for i, b in enumerate(grupo):
            fx = mx + (i % 2) * (ancho + gap)
            fy = top + (i // 2) * (alto + gap)
            _bloque(slide, b, fx, fy, ancho, alto, log)
            nfotos += min(4, len([f for f in (b.get("fotos") or []) if _s(f)]))
        _texto(slide, f"{len(prs.slides):02d}", W - 1.0, H - 0.38, 0.7, 0.3, 8.5, GRIS, alin=PP_ALIGN.RIGHT)

    out = io.BytesIO(); prs.save(out)
    log.insert(0, f"✔ Agregadas {len(grupos)} lámina(s) «{_s(datos.get('titulo'))}» ({len(bloques)} bloques, {nfotos} fotos)")
    return out.getvalue(), log


if __name__ == "__main__":
    import sys
    base = open(sys.argv[1], "rb").read()
    def foto(c):
        b = io.BytesIO(); Image.new("RGB", (800, 500), c).save(b, "PNG"); return base64.b64encode(b.getvalue()).decode()
    bl = [
        {"titulo": "Coordinación", "intro": "Trabajos TFDS con EECC GPRO.", "items": ["Focos de cada empresa.", "Análisis y reflexión accidente fatal."], "fotos": [foto((90, 120, 160))]},
        {"titulo": "Procesos y rol", "items": ["Confirmación de procesos a tareas críticas."], "fotos": [foto((200, 120, 60)), foto((60, 160, 90))]},
        {"titulo": "Ejecución", "intro": "Chequeo de focos entregados.", "items": ["Izajes: Plan de izaje específico, control de línea de fuego.", "Aislamiento: Integridad y correcta fijación."], "fotos": [foto((150, 50, 50)), foto((50, 50, 150)), foto((50, 150, 150))]},
        {"titulo": "Visita gerencial", "intro": "Gerente Residente realiza visita al área de trabajo.", "fotos": [foto((120, 120, 120))]},
        {"titulo": "Quinto bloque", "items": ["Sin fotos."]},
    ]
    r, lg = agregar_preventiva(base, {"lema": "Gestión en seguridad · fin de semana 06–08 nov", "titulo": "Gestión preventiva ejecutada en terreno", "subtitulo": "Acciones desarrolladas durante el período", "bloques": bl})
    open("/tmp/prev.pptx", "wb").write(r); print(lg)
