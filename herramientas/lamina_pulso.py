"""Lámina manual al final del Assurance: «Seguridad | ¿Qué hicimos para sostener el resultado?».

agregar_pulso(pptx_bytes, pulso) -> (bytes, log). Todo el texto lo escribe el usuario en la página.
"""
import io
import re
import copy

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

DEF = {
    "lema": "GESTIONAMOS · DETECTAMOS · TRANSFERIMOS",
    "titulo": "Seguridad | ¿Qué hicimos para sostener el resultado?",
    "subtitulo": (
        "Lectura objetiva de las señales del turno: qué ocurrió, dónde se "
        "concentró el riesgo y cuán oportunamente cerramos."
    ),
    "kpis": [
        {
            "etiqueta": "EDC REVISADAS – Gestión del NO",
            "valor": "",
            "nota": "Durante el turno",
        },
        {
            "etiqueta": "ARTP REVISADAS",
            "valor": "",
            "nota": "Durante el turno",
        },
        {
            "etiqueta": "OTRA HERRAMIENTA DE GESTIÓN (CdP/CdR, etc.)",
            "valor": "",
            "nota": "Durante el turno",
        },
    ],
    "paneles": [
        {
            "rotulo": "RECIBIMOS",
            "titulo": "Focos del turno",
            "items": [],
            "estilo": "numeros",
        },
        {
            "rotulo": "GESTIONAMOS",
            "titulo": "Qué hicimos frente al foco",
            "items": [],
            "estilo": "numeros",
        },
        {
            "rotulo": "TRANSFERIMOS",
            "titulo": "Qué debe continuar",
            "items": [],
            "estilo": "vinetas",
        },
    ],
    "cierre_rotulo": "PREGUNTA DE CIERRE",
    "cierre": (
        "¿Qué hicimos que debemos repetir para seguir evitando accidentes y "
        "sostener los buenos resultados? (Buenas Prácticas)"
    ),
    "pie": "Antucoya · Seguridad y Salud Ocupacional · Pulso de Seguridad",
}

TINTA = RGBColor(0x1F, 0x2D, 0x3D)
GRIS = RGBColor(0x5F, 0x6B, 0x73)
TEAL = RGBColor(0x0E, 0x7C, 0x86)
NUM = RGBColor(0x13, 0x8A, 0x9C)
ORO = RGBColor(0xF5, 0xB8, 0x00)
BORDE = RGBColor(0xDA, 0xDF, 0xE3)
NEGRO = RGBColor(0x22, 0x22, 0x22)
CIERRE = RGBColor(0x3A, 0x3A, 0x3A)
ROJO = RGBColor(0xD9, 0x3A, 0x3A)
NARANJA = RGBColor(0xF0, 0x7F, 0x3C)

def _mezclar(pulso):
    """Aplica valores por defecto, respetando cadenas vacías y listas por posición."""
    pulso = pulso if isinstance(pulso, dict) else {}
    resultado = {}

    for clave, defecto in DEF.items():
        if clave in ("kpis", "paneles"):
            continue
        valor = pulso.get(clave, None)
        resultado[clave] = copy.deepcopy(defecto) if valor is None else valor

    for clave in ("kpis", "paneles"):
        base = DEF[clave]
        recibido = pulso.get(clave, None)

        if recibido is None or not isinstance(recibido, list):
            resultado[clave] = copy.deepcopy(base)
            continue

        combinados = []
        for indice, defecto_item in enumerate(base):
            item_recibido = recibido[indice] if indice < len(recibido) else None

            if item_recibido is None or not isinstance(item_recibido, dict):
                combinados.append(copy.deepcopy(defecto_item))
                continue

            item = copy.deepcopy(defecto_item)
            for campo, defecto_campo in defecto_item.items():
                valor = item_recibido.get(campo, None)
                item[campo] = copy.deepcopy(defecto_campo) if valor is None else valor
            combinados.append(item)

        resultado[clave] = combinados

    return resultado

def _limpiar_items(lista):
    """Limpia líneas, quitando numeración o viñetas iniciales."""
    if not isinstance(lista, list):
        return []

    resultado = []
    for item in lista:
        if item is None:
            continue

        texto = str(item).strip()
        texto = re.sub(r"^\d+[.)-]?\s*", "", texto)
        texto = re.sub(r"^[•●\-*]\s*", "", texto).strip()

        if texto:
            resultado.append(texto)

    return resultado

def _sin_sombra(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass

def _rect(slide, x, y, w, h, fill_color, line_color, line_width, redondeado=True):
    tipo = MSO_SHAPE.ROUNDED_RECTANGLE if redondeado else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        tipo, Inches(x), Inches(y), Inches(w), Inches(h)
    )

    if redondeado:
        try:
            shape.adjustments[0] = 0.08
        except Exception:
            pass

    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color

    shape.line.color.rgb = line_color
    shape.line.width = Pt(line_width)

    _sin_sombra(shape)
    return shape

def _texto(
    slide,
    texto,
    x,
    y,
    w,
    h,
    tamano,
    color,
    negrita=False,
    alineacion=PP_ALIGN.LEFT,
    anclaje=MSO_ANCHOR.TOP,
):
    shape = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    _sin_sombra(shape)

    try:
        shape.fill.background()
        shape.line.fill.background()
    except Exception:
        pass

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anclaje
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)

    p = tf.paragraphs[0]
    p.text = "" if texto is None else str(texto)
    p.alignment = alineacion
    p.font.name = "Calibri"
    p.font.size = Pt(tamano)
    p.font.bold = negrita
    p.font.color.rgb = color
    for r in p.runs:
        r.font.name, r.font.size, r.font.bold, r.font.color.rgb = "Calibri", Pt(tamano), negrita, color

    return shape

def _configurar_parrafo(parrafo, tamano, color, space_after=0):
    parrafo.font.name = "Calibri"
    parrafo.font.size = Pt(tamano)
    parrafo.font.color.rgb = color
    parrafo.space_after = Pt(space_after)
    for r in parrafo.runs:
        if r.font.size is None:
            r.font.name, r.font.size, r.font.color.rgb = "Calibri", Pt(tamano), color

def _agregar_contenido_panel(slide, panel, x, y, w, h):
    items = _limpiar_items(panel.get("items", []))
    cantidad = len(items)

    if cantidad <= 6:
        tamano = 11
    elif cantidad <= 9:
        tamano = 10
    elif cantidad <= 12:
        tamano = 9
    else:
        tamano = 8

    contenido = _texto(
        slide,
        "",
        x + 0.22,
        y + 0.85,
        w - 0.4,
        h - 1.0,
        tamano,
        NEGRO,
        anclaje=MSO_ANCHOR.TOP,
    )

    tf = contenido.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)

    estilo = panel.get("estilo", "numeros")

    if estilo == "vinetas":
        if not items:
            items = ["", "", ""]

        for indice, item in enumerate(items):
            p = tf.paragraphs[0] if indice == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            _configurar_parrafo(p, tamano, NEGRO, 3)

            bullet = p.add_run()
            bullet.text = "●  "
            bullet.font.name = "Calibri"
            bullet.font.size = Pt(tamano + 2)
            bullet.font.color.rgb = ROJO if indice == 0 else NARANJA

            texto = p.add_run()
            texto.text = item
            texto.font.name = "Calibri"
            texto.font.size = Pt(tamano)
            texto.font.color.rgb = NEGRO
    else:
        if not items:
            items = [""]

        for indice, item in enumerate(items, start=1):
            p = tf.paragraphs[0] if indice == 1 else tf.add_paragraph()
            p.text = f"{indice}. {item}" if item else f"{indice}."
            p.alignment = PP_ALIGN.LEFT
            _configurar_parrafo(p, tamano, NEGRO, 3)

def agregar_pulso(pptx_bytes, pulso):
    datos = _mezclar(pulso)
    prs = Presentation(io.BytesIO(pptx_bytes))

    logo_blob = None
    if len(prs.slides) > 0:
        for shape in prs.slides[0].shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    logo_blob = shape.image.blob
                except Exception:
                    logo_blob = None
                break

    layout = min(prs.slide_layouts, key=lambda item: len(item.placeholders))
    slide = prs.slides.add_slide(layout)

    for shape in list(slide.shapes):
        if shape.is_placeholder:
            element = shape._element
            element.getparent().remove(element)

    colores_circulos = [
        RGBColor(0xF5, 0xA3, 0x00),
        RGBColor(0xE3, 0x00, 0x1B),
        RGBColor(0x5B, 0xC8, 0xD0),
        TEAL,
    ]
    for x, color in zip((0.28, 0.48, 0.68, 0.88), colores_circulos):
        circulo = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(x), Inches(0.45), Inches(0.15), Inches(0.15)
        )
        circulo.fill.solid()
        circulo.fill.fore_color.rgb = color
        circulo.line.fill.background()
        _sin_sombra(circulo)

    if logo_blob:
        try:
            logo = slide.shapes.add_picture(
                io.BytesIO(logo_blob),
                Inches(11.35),
                Inches(0.25),
                Inches(1.8),
                Inches(0.7),
            )
            _sin_sombra(logo)
        except Exception:
            pass

    _texto(slide, datos["lema"], 0.25, 0.65, 8.0, 0.3, 12, TEAL, True)
    _texto(slide, datos["titulo"], 0.25, 0.95, 10.8, 0.55, 26, TINTA, True)
    _texto(slide, datos["subtitulo"], 0.35, 1.5, 10.8, 0.35, 12, GRIS)

    for kpi, x in zip(datos["kpis"], (0.54, 4.40, 8.24)):
        _rect(slide, x, 2.0, 3.33, 1.3, RGBColor(255, 255, 255), BORDE, 0.75)

        _texto(
            slide, kpi["etiqueta"], x + 0.28, 2.12, 2.93, 0.38, 9, GRIS, True
        )
        _texto(
            slide, kpi["valor"], x + 0.28, 2.5, 2.93, 0.5, 24, NUM, True
        )
        _texto(
            slide, kpi["nota"], x + 0.28, 2.98, 2.93, 0.26, 9, GRIS
        )

    posiciones_paneles = (
        (0.57, 3.85),
        (4.56, 4.32),
        (8.98, 3.85),
    )

    for panel, (x, w) in zip(datos["paneles"], posiciones_paneles):
        _rect(slide, x, 3.65, w, 4.1, RGBColor(255, 255, 255), BORDE, 1)

        _texto(slide, panel["rotulo"], x + 0.25, 3.73, w - 0.4, 0.28, 10, ORO, True)
        _texto(slide, panel["titulo"], x + 0.25, 4.01, w - 0.4, 0.42, 15, TINTA, True)
        _agregar_contenido_panel(slide, panel, x, 3.65, w, 4.1)

    _texto(slide, datos["cierre_rotulo"], 0.6, 8.0, 6.0, 0.3, 10, GRIS, True)
    _texto(slide, datos["cierre"], 0.6, 8.3, 12.3, 0.55, 17, CIERRE, True)

    _texto(slide, datos["pie"], 0.4, 9.12, 9.0, 0.3, 8.5, GRIS)
    _texto(
        slide,
        f"{len(prs.slides):02d}",
        12.3,
        9.12,
        0.7,
        0.3,
        8.5,
        GRIS,
        alineacion=PP_ALIGN.RIGHT,
    )

    salida = io.BytesIO()
    prs.save(salida)

    return (
        salida.getvalue(),
        ["✔ Agregada la lámina «Seguridad | ¿Qué hicimos para sostener el resultado?» (manual)"],
    )
