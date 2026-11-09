"""Láminas extra al final del Assurance: "GESTIÓN SSO GPRO" y "Reportabilidad, Autorreportes y Focos de turno".

- calcular(an, params) -> datos: cifras de los 7 días del turno, solo eventos de PUMA y COX.
- agregar(pptx_bytes, datos) -> (bytes, log): dibuja las 2 láminas al final del PPT.
"""
import io
import re
import unicodedata
from datetime import date, timedelta

import catalogo_temas as T

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]
EMPRESAS = ["PUMA", "COX S.A"]
CORTO = {"PUMA": "PUMA", "COX S.A": "COX"}
SEVERIDAD = ["AAP", "CAP", "HAP", "ABP", "CBP", "HBP"]
CONECTORES = {"de", "del", "la", "las", "los", "el", "y", "e"}

# ----------------------------------------------------------------------
# PARTE 1 · datos
# ----------------------------------------------------------------------

def _n(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def _cortar(s, n):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    if len(s) <= n:
        return s
    corte = s[:n - 1].rsplit(" ", 1)[0].rstrip(" ,;:-–(")
    return corte + "…"


def _mayus(s):
    s = str(s or "").strip()
    return s[:1].upper() + s[1:]


def _minus(s):
    s = str(s or "").strip()
    if len(s) > 1 and s[:2].isupper():       # sigla: se deja igual
        return s
    return s[:1].lower() + s[1:]


def area_bonita(s):
    if not s or _n(s) == "SIN AREA":
        return ""
    out = []
    for i, w in enumerate(str(s).split()):
        wl = w.lower()
        if i and wl in CONECTORES:
            out.append(wl)
        elif re.search(r"[\d\-]", w) or (len(re.sub(r"\W", "", w)) <= 3 and wl not in CONECTORES):
            out.append(w.upper())
        else:
            out.append(wl[:1].upper() + wl[1:])
    return " ".join(out)


def _empresa(ev):
    e = _n(ev.get("empresa"))
    return "PUMA" if "PUMA" in e else "COX S.A" if "COX" in e else None


def tipo6(ev):
    c, p = _n(ev.get("clasificacion")), _n(ev.get("potencialidad"))
    le = {"HALLAZGO": "H", "CUASI ACCIDENTE": "C", "ACCIDENTE": "A"}.get(c)
    po = "AP" if "ALTO" in p else "BP" if "BAJO" in p else None
    if le and po:
        return le + po
    return {"CAAP": "CAP", "CBP": "CBP", "HAP": "HAP", "HBP": "HBP", "ACC": "ABP"}.get(ev.get("tipo"))


def _vinculado(ev):
    r = _n(ev.get("riesgo"))
    return bool(r) and not r.startswith("OTRO") and r not in ("SIN INFORMACION", "NO APLICA")


def _control_ok(ev):
    return _n(ev.get("control_critico")) not in ("", "OTRO", "SIN INFORMACION", "NO APLICA")


def grupo_reportador(s):
    t = _n(s)
    if not t:
        return None
    if "JEJ" in t or "J.E.J" in t or "J E J" in t:
        return "JEJ"
    if "ANTUCOYA" in t or "AMSA" in t or re.search(r"\bANT\b", t):
        return "ANT"
    if "COX" in t:
        return "COX"
    if "PUMA" in t:
        return "PUMA"
    return "Otros"


def _f(s):
    try:
        return date.fromisoformat(str(s)[:10])
    except Exception:
        return None


def rango_turno(d1, d2):
    if not d1 or not d2:
        return ""
    m1, m2 = MESES[d1.month - 1].capitalize(), MESES[d2.month - 1].capitalize()
    if d1.year != d2.year:
        return f"{d1.day:02d} de {m1} {d1.year} al {d2.day:02d} de {m2} {d2.year}"
    if d1.month != d2.month:
        return f"{d1.day:02d} de {m1} al {d2.day:02d} de {m2} {d2.year}"
    return f"{d1.day:02d} al {d2.day:02d} de {m2} {d2.year}"


def _entre(a, b):
    if a.month == b.month:
        return f"entre el {a.day} y {b.day} de {MESES[b.month - 1]}"
    return f"entre el {a.day} de {MESES[a.month - 1]} y el {b.day} de {MESES[b.month - 1]}"


def _insight(evs, d1, d2):
    N = len(evs)
    if not N:
        return {"negrita": "Sin reportes de PUMA y COX en el turno.", "normal": ""}
    por_dia = {}
    for e in evs:
        por_dia[e["_f"]] = por_dia.get(e["_f"], 0) + 1
    neg = ""
    if N >= 3 and d1 and d2 and (d2 - d1).days >= 2:
        mejor, k = None, -1
        d = d1
        while d + timedelta(days=2) <= d2:
            c = sum(por_dia.get(d + timedelta(days=i), 0) for i in range(3))
            if c > k:
                mejor, k = d, c
            d += timedelta(days=1)
        if k / N >= 0.5:
            neg = (f"{round(100 * k / N)}% de los reportes ({k}/{N}) se concentran "
                   f"{_entre(mejor, mejor + timedelta(days=2))}.")
    if not neg:
        dia, k = max(por_dia.items(), key=lambda x: (x[1], x[0]))
        neg = f"El día con más reportes fue el {dia.day} de {MESES[dia.month - 1]} ({k} de {N})."
    areas = {}
    for e in sorted(evs, key=lambda x: x["_f"]):
        a = area_bonita(e.get("area_geo"))
        if a:
            areas[a] = (areas.get(a, (0, None))[0] + 1, e["_f"])
    normal = ""
    if areas:
        a, (n, _) = max(areas.items(), key=lambda x: (x[1][0], x[1][1]))
        normal = f" Área de mayor incidencia: {a} ({n} de {N})."
    return {"negrita": neg, "normal": normal}


def _accion_corta(texto):
    """Primera parte de una instrucción (hasta ';' o punto), con corte en palabra si es larga."""
    t = re.sub(r"\s+", " ", str(texto or "")).strip()
    if not t:
        return ""
    seg = re.split(r";\s+|\.\s+", t, maxsplit=1)[0].strip().rstrip(".")
    seg = _mayus(_cortar(seg, 112))
    return seg if seg.endswith("…") else seg + "."


def _item(ev, repetidos, acciones):
    t6 = ev["_t6"]
    area = area_bonita(ev.get("area_geo"))
    try:
        tema = (T.asignar_tema(ev) or {}).get("tema_id")
    except Exception:
        tema = None
    titulo = _mayus(T.titulo_evento(ev, 58).rstrip(".")) or _mayus(_cortar(ev.get("titulo"), 58))
    frase = _mayus(T.resumir_evento(ev, tema, 105).rstrip("."))
    descripcion = (frase + (f" ({area})." if area else ".")) if frase else (f"({area})." if area else "")
    riesgo = str(ev.get("riesgo") or "").strip()
    ctrl = str(ev.get("control_critico") or "").strip() if _control_ok(ev) else ""
    edc_largo = riesgo + (" — " + _minus(_cortar(ctrl, 85)) if ctrl else "")
    edc_largo += "" if edc_largo.endswith(("…", ".")) else "."
    edc_corto = riesgo + (" · " + _mayus(_cortar(ctrl, 50)) if ctrl else "")
    t = T._tema_por_id(tema) if tema else None
    accion = (_accion_corta(acciones.get(str(ev.get("id"))))          # instrucción del foco del Assurance
              or _accion_corta((t or {}).get("instruccion"))             # o la del tema del catálogo
              or "Revisar en terreno con la supervisión y verificar el cierre antes del siguiente turno.")
    rojo = t6 in ("AAP", "CAP", "HAP") or repetidos.get(_n(riesgo), 0) >= 2
    return {"id": str(ev.get("id")), "tipo": t6 or "—", "empresa": ev["_emp"], "area": area,
            "abierto": bool(ev.get("abierto")), "color": "rojo" if rojo else "dorado", "titulo": titulo,
            "descripcion": descripcion, "edc_largo": edc_largo, "edc_corto": edc_corto, "accion": accion}


def calcular(an, params=None, plan=None):
    """Cifras de las 2 láminas a partir del análisis (`an`) del Assurance y del plan (focos) que se genera."""
    log = []
    acciones = {}
    for f in (plan or {}).get("focos") or []:
        for e in f.get("eventos") or []:
            for x in (e.get("ids") or [e.get("id")]):
                if x and f.get("instruccion"):
                    acciones.setdefault(str(x), f["instruccion"])
    tr = an.get("turno") or {}
    d1, d2 = _f(tr.get("desde")), _f(tr.get("hasta"))
    evs = []
    for ev in an.get("eventos") or []:
        if not ev.get("en_turno"):
            continue
        emp = _empresa(ev)
        if not emp:
            continue
        e = dict(ev)
        e["_emp"], e["_t6"], e["_f"] = emp, tipo6(ev), _f(ev.get("fecha"))
        if e["_f"] is None:
            continue
        evs.append(e)
    N = len(evs)
    cats = ["HBP", "HAP", "CAP", "CBP", "ABP"] + (["AAP"] if any(e["_t6"] == "AAP" for e in evs) else [])
    matriz = {emp: [sum(1 for e in evs if e["_emp"] == emp and e["_t6"] == c) for c in cats] for emp in EMPRESAS}
    sin_tipo = [e["id"] for e in evs if e["_t6"] not in cats]
    if sin_tipo:
        log.append(f"⚠ Gestión SSO: eventos sin clasificación reconocida (no van al gráfico): {', '.join(sin_tipo)}")

    vinc = [e for e in evs if _vinculado(e)]
    repetidos = {}
    for e in vinc:
        repetidos[_n(e.get("riesgo"))] = repetidos.get(_n(e.get("riesgo")), 0) + 1
    sev = {t: i for i, t in enumerate(SEVERIDAD)}
    vinc.sort(key=lambda e: (sev.get(e["_t6"], 99), -e["_f"].toordinal(), -int(e["id"]) if str(e["id"]).isdigit() else 0))
    items = [_item(e, repetidos, acciones) for e in vinc[:4]]
    if len(vinc) > 4:
        log.append(f"Gestión SSO: {len(vinc)} eventos vinculados a EDC; se muestran los 4 más graves.")

    # Reportadores
    for e in evs:
        e["_g"] = grupo_reportador(e.get("reportador"))
    con_dato = any(e["_g"] for e in evs)
    orden = ["JEJ", "ANT", "PUMA", "COX", "Otros"]
    conteo = {g: sum(1 for e in evs if e["_g"] == g) for g in orden}
    grupos = [g for g in orden if conteo[g] > 0]
    tabla = {g: {**{emp: sum(1 for e in evs if e["_g"] == g and e["_emp"] == emp) for emp in EMPRESAS},
                 "Total": conteo[g]} for g in grupos}
    totales = {**{emp: sum(tabla[g][emp] for g in grupos) for emp in EMPRESAS}, "Total": sum(conteo[g] for g in grupos)}
    if N and not con_dato:
        log.append("⚠ Reportador vacío en el Excel: la tabla de autorreportes queda sin datos")
    elif N and any(e["_g"] is None for e in evs):
        log.append("⚠ Hay eventos sin Reportador: no se cuentan en la tabla de autorreportes")

    asp_evs = [e for e in evs if re.search(r"\bASP\b", str(e.get("titulo") or ""), re.I)]
    asp = {"total": len(asp_evs), "por_empresa": {emp: sum(1 for e in asp_evs if e["_emp"] == emp) for emp in EMPRESAS}}

    tot_emp = {emp: sum(1 for e in evs if e["_emp"] == emp) for emp in EMPRESAS}
    if not N:
        auto = {"titulo": "AUTORREPORTABILIDAD EECC", "valor": "—", "detalle": "Sin reportes de PUMA y COX en el turno."}
    else:
        props = {emp: sum(1 for e in evs if e["_emp"] == emp and e["_g"] == CORTO[emp]) for emp in EMPRESAS}
        pct = {emp: round(100 * props[emp] / tot_emp[emp]) for emp in EMPRESAS if tot_emp[emp]}
        elegida = max(pct, key=lambda emp: (pct[emp], tot_emp[emp]))
        orden_det = [elegida] + [x for x in pct if x != elegida]
        if con_dato:
            auto = {"titulo": f"AUTORREPORTABILIDAD EECC ({CORTO[elegida]})", "valor": f"{pct[elegida]}%",
                    "detalle": " · ".join(f"{CORTO[x]} {props[x]} de {tot_emp[x]}" for x in orden_det)}
        else:
            auto = {"titulo": f"AUTORREPORTABILIDAD EECC ({CORTO[elegida]})", "valor": "—",
                    "detalle": "El Excel no trae datos en la columna Reportador."}

    log.insert(0, f"✔ Láminas Gestión SSO GPRO y Reportabilidad: {N} reportes PUMA + COX del turno, "
                  f"{len(vinc)} vinculados a EDC")
    return {"rango_turno": rango_turno(d1, d2),
            "kpi": {"reportes": N, "vinculados_edc": len(vinc), "empresas": sum(1 for v in tot_emp.values() if v)},
            "empresas": EMPRESAS, "categorias": cats, "matriz": matriz, "insight": _insight(evs, d1, d2),
            "items": items,
            "reportadores": {"con_dato": con_dato, "grupos": grupos, "conteo": {g: conteo[g] for g in grupos},
                             "tabla": tabla, "totales": totales},
            "asp": asp, "autorreporte": auto, "log": log}

# ----------------------------------------------------------------------
# PARTE 2 · dibujo
# ----------------------------------------------------------------------

AZUL, ROJO, TEAL = "0F2D52", "D2343C", "0E7C86"
DORADO, DORADO_TXT = "F0A800", "A07800"
GRIS_TXT, TEXTO, GRIS_FONDO, BORDE, FONDO_FILA = "5F6B73", "222222", "EEF1F3", "DADFE3", "F4F7F8"
SERIF, SANS = "Georgia", "Calibri"


def _pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
    from pptx.chart.data import CategoryChartData
    return locals()


P = None   # módulos de python-pptx (se cargan al usar agregar)


def _rgb(h):
    return P["RGBColor"].from_string(h)


def _rect(sl, x, y, w, h, fill=None, line=None, lw=0.75, redondo=None):
    In = P["Inches"]
    forma = P["MSO_SHAPE"].ROUNDED_RECTANGLE if redondo is not None else P["MSO_SHAPE"].RECTANGLE
    sh = sl.shapes.add_shape(forma, In(x), In(y), In(w), In(h))
    if redondo is not None:
        sh.adjustments[0] = redondo
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = _rgb(line)
        sh.line.width = P["Pt"](lw)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.text = ""
    return sh


def _llenar(tf, parrafos, align="l", anchor="t", margen=0.05, margen_izq=None):
    """parrafos: lista de (runs, space_before_pt); run = (texto, pt, bold, color, font, italic)."""
    In, Pt = P["Inches"], P["Pt"]
    tf.word_wrap = True
    tf.auto_size = P["MSO_AUTO_SIZE"].NONE
    tf.margin_left = In(margen if margen_izq is None else margen_izq)
    tf.margin_right = tf.margin_top = tf.margin_bottom = In(margen)
    tf.vertical_anchor = {"t": P["MSO_ANCHOR"].TOP, "m": P["MSO_ANCHOR"].MIDDLE, "b": P["MSO_ANCHOR"].BOTTOM}[anchor]
    al = {"l": P["PP_ALIGN"].LEFT, "c": P["PP_ALIGN"].CENTER, "r": P["PP_ALIGN"].RIGHT}[align]
    for i, (runs, antes) in enumerate(parrafos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = al
        if antes:
            p.space_before = Pt(antes)
        for texto, pt, bold, color, font, italic in runs:
            r = p.add_run()
            r.text = str(texto)
            f = r.font
            f.size, f.bold, f.italic, f.name = Pt(pt), bold, italic, font
            f.color.rgb = _rgb(color)


def _run(texto, pt, bold=False, color=TEXTO, font=SANS, italic=False):
    return (texto, pt, bold, color, font, italic)


def _texto(sl, x, y, w, h, parrafos, align="l", anchor="t", margen=0.05, margen_izq=None):
    In = P["Inches"]
    tb = sl.shapes.add_textbox(In(x), In(y), In(w), In(h))
    _llenar(tb.text_frame, parrafos, align, anchor, margen, margen_izq)
    return tb


def _nueva_lamina(prs):
    lay = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    sl = prs.slides.add_slide(lay)
    for ph in list(sl.placeholders):
        ph._element.getparent().remove(ph._element)
    return sl


def _logo(prs):
    if not len(prs.slides):
        return None
    for sh in prs.slides[0].shapes:
        if sh.shape_type == P["MSO_SHAPE_TYPE"].PICTURE:
            return sh.image.blob
    return None


def _poner_logo(sl, logo):
    if logo:
        In = P["Inches"]
        sl.shapes.add_picture(io.BytesIO(logo), In(11.35), In(0.25), In(1.8), In(0.7))


def _col(item, fuerte=True):
    if item.get("color") == "rojo":
        return ROJO
    return DORADO if fuerte else DORADO_TXT


def _lamina_a(prs, d, logo):
    In, Pt = P["Inches"], P["Pt"]
    sl = _nueva_lamina(prs)
    _texto(sl, 0.19, 0.25, 9.0, 0.7, [([_run("GESTIÓN SSO GPRO", 30, True, AZUL, SERIF)], 0)], anchor="m")
    _poner_logo(sl, logo)
    k = d.get("kpi") or {}
    for x, val, etq, col in ((0.19, k.get("reportes", 0), "Reportes del período (turno)", AZUL),
                             (4.66, k.get("vinculados_edc", 0), "Reportes vinculados EDC", ROJO),
                             (9.15, k.get("empresas", 0), "Empresas EECC", TEAL)):
        y, w, h = 1.35, 3.98, 1.15
        _rect(sl, x, y, w, h, GRIS_FONDO, BORDE, redondo=0.08)
        _rect(sl, x + 0.12, y + 0.15, 0.11, h - 0.3, col)
        _texto(sl, x + 0.32, y + 0.08, 1.4, h - 0.16, [([_run(val, 44, True, col, SERIF)], 0)], anchor="m")
        _texto(sl, x + 1.7, y, w - 1.85, h, [([_run(etq, 14, False, GRIS_TXT)], 0)], align="c", anchor="m")

    # Panel izquierdo: gráfico
    _rect(sl, 0.19, 2.8, 6.3, 6.45, "FFFFFF", BORDE, redondo=0.03)
    _texto(sl, 0.45, 3.0, 5.8, 0.45, [([_run("CLASIFICACIÓN DE EVENTOS POR EMPRESA", 14, True, AZUL, SERIF)], 0)],
           anchor="m")
    cd = P["CategoryChartData"]()
    cd.categories = d.get("categorias") or []
    for emp in d.get("empresas") or []:
        cd.add_series(emp, (d.get("matriz") or {}).get(emp) or [0] * len(cd.categories))
    gf = sl.shapes.add_chart(P["XL_CHART_TYPE"].COLUMN_CLUSTERED, In(0.4), In(3.55), In(5.9), In(4.35), cd)
    ch = gf.chart
    ch.font.name, ch.font.size = SANS, Pt(11)
    ch.has_legend = True
    ch.legend.position = P["XL_LEGEND_POSITION"].BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(11), SANS
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.has_minor_gridlines = False
    va.tick_labels.font.size = Pt(1)
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.bold = Pt(12), True
    ca.tick_labels.font.color.rgb = _rgb("333333")
    ca.format.line.color.rgb = _rgb("BFC5CA")
    ca.has_major_gridlines = False
    plot = ch.plots[0]
    plot.gap_width, plot.overlap = 80, 0
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value = True
    dl.number_format, dl.number_format_is_linked = "0", False
    dl.position = P["XL_LABEL_POSITION"].OUTSIDE_END
    dl.font.size, dl.font.bold, dl.font.name = Pt(11), True, SANS
    dl.font.color.rgb = _rgb("333333")
    for serie, col in zip(plot.series, (AZUL, TEAL)):
        serie.format.fill.solid()
        serie.format.fill.fore_color.rgb = _rgb(col)
        serie.invert_if_negative = False
    ins = d.get("insight") or {}
    _texto(sl, 0.45, 8.0, 5.8, 1.0, [([_run(ins.get("negrita", ""), 12, True, TEXTO),
                                        _run(ins.get("normal", ""), 12, False, GRIS_TXT)], 0)])

    # Panel derecho: hallazgos
    _rect(sl, 6.8, 2.8, 6.33, 6.45, "FFFFFF", ROJO, lw=2, redondo=0.03)
    cab = _rect(sl, 6.8, 2.8, 6.33, 0.6, ROJO, redondo=0.08)
    _llenar(cab.text_frame, [([_run("HALLAZGOS RELEVANTES — ASOCIADOS A EDC", 14, True, "FFFFFF", SERIF)], 0)],
            anchor="m", margen_izq=0.2)
    items = d.get("items") or []
    if not items:
        _texto(sl, 6.8, 5.6, 6.33, 0.6, [([_run("Sin hallazgos asociados a EDC en el turno.", 12, False, GRIS_TXT)], 0)],
               align="c", anchor="m")
    for i, it in enumerate(items[:4]):
        x, y, w, h = (7.0, 10.06)[i % 2], (3.62, 6.42)[i // 2], 2.88, 2.65
        _rect(sl, x, y, w, h, "FFFFFF", BORDE, redondo=0.04)
        _rect(sl, x, y, 0.08, h, _col(it, False))
        _texto(sl, x + 0.18, y + 0.12, w - 0.3, h - 0.24, [
            ([_run(it["tipo"], 11.5, True, _col(it, False)),
              _run(f" · ID {it['id']} · {it['empresa']}", 11.5, True, AZUL)], 0),
            ([_run(it["descripcion"], 11, False, TEXTO)], 4),
            ([_run("EDC: ", 9.5, True, TEAL), _run(it["edc_largo"], 9.5, False, GRIS_TXT)], 8)])
    return sl


def _lamina_b(prs, d, logo):
    In, Pt = P["Inches"], P["Pt"]
    sl = _nueva_lamina(prs)
    _texto(sl, 0.19, 0.15, 11.0, 0.7, [([_run("Reportabilidad, Autorreportes y Focos de turno", 26, True, AZUL, SERIF)], 0)],
           anchor="m")
    _texto(sl, 0.19, 0.85, 10.8, 0.4, [([
        _run("Reportes por EECC, Programa ASP y entrega de turno de focos de Alto Potencial", 12, False, GRIS_TXT),
        _run("  |  ", 12, False, GRIS_TXT), _run(d.get("rango_turno", ""), 12, True, TEAL)], 0)], anchor="m")
    _poner_logo(sl, logo)

    # Reportabilidad por empresa
    _rect(sl, 0.19, 1.4, 7.8, 2.65, "FFFFFF", BORDE, redondo=0.04)
    _texto(sl, 0.4, 1.55, 7.4, 0.35, [([_run("REPORTABILIDAD / AUTORREPORTES POR EMPRESA", 13, True, AZUL)], 0)])
    _texto(sl, 0.4, 1.88, 4.5, 0.3, [([_run("Reportes proactivos del personal de cada EECC (según reportador).", 9.5,
                                            False, GRIS_TXT, italic=True)], 0)])
    rp = d.get("reportadores") or {}
    colores = {"JEJ": TEAL, "ANT": AZUL, "PUMA": ROJO, "COX": "9AA3AA", "Otros": "9AA3AA"}
    if rp.get("con_dato"):
        for i, g in enumerate((rp.get("grupos") or [])[:3]):
            x, y, w, h = (0.4, 1.9, 3.4)[i], 2.25, 1.37, 1.55
            _rect(sl, x, y, w, h, GRIS_FONDO, BORDE, redondo=0.06)
            _rect(sl, x, y, w, 0.08, colores.get(g, "9AA3AA"))
            num_col = "8A939A" if g in ("COX", "Otros") else colores.get(g, AZUL)
            _texto(sl, x, y + 0.2, w, 0.8, [([_run((rp.get("conteo") or {}).get(g, 0), 40, True, num_col, SERIF)], 0)],
                   align="c", anchor="m")
            _texto(sl, x, y + 1.05, w, 0.35, [([_run(g, 11, True, "333333")], 0)], align="c", anchor="m")
    else:
        _rect(sl, 0.4, 2.25, 4.37, 1.55, GRIS_FONDO, BORDE, redondo=0.06)
        _texto(sl, 0.4, 2.25, 4.37, 1.55, [([_run("Sin datos de reportador en el Excel", 12, False, GRIS_TXT)], 0)],
               align="c", anchor="m")
    _texto(sl, 4.95, 1.98, 2.95, 0.25, [([_run("ORIGEN DEL REPORTE  ·  REPORTADOR → EECC", 8.5, True, AZUL)], 0)])
    if rp.get("con_dato"):
        filas = [[g] + [rp["tabla"][g][k] for k in ("PUMA", "COX S.A", "Total")] for g in rp.get("grupos") or []]
        tot = rp.get("totales") or {}
        filas.append(["Total", tot.get("PUMA", 0), tot.get("COX S.A", 0), tot.get("Total", 0)])
    else:
        filas = [["Sin dato", "—", "—", "—"]]
    filas = [["Reportó", "PUMA", "COX", "Total"]] + filas
    anchos = (1.0, 0.65, 0.65, 0.65)
    gt = sl.shapes.add_table(len(filas), 4, In(4.95), In(2.28), In(sum(anchos)), In(0.27 * len(filas)))
    tbl = gt.table
    tbl.first_row = tbl.horz_banding = False
    for j, a in enumerate(anchos):
        tbl.columns[j].width = In(a)
    for i, fila in enumerate(filas):
        tbl.rows[i].height = In(0.27)
        es_cab, es_tot = i == 0, rp.get("con_dato") and i == len(filas) - 1
        for j, v in enumerate(fila):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = _rgb(AZUL if es_cab else "D9DEE2" if es_tot else "FFFFFF")
            c.margin_left = c.margin_right = In(0.05)
            c.margin_top = c.margin_bottom = In(0.01)
            _llenar(c.text_frame, [([_run(v, 9, es_cab or bool(es_tot), "FFFFFF" if es_cab else "333333")], 0)],
                    align="l" if j == 0 else "c", anchor="m", margen=0.01)
            c.margin_left = c.margin_right = In(0.05)

    # Programa ASP
    asp = d.get("asp") or {}
    n_asp = asp.get("total", 0)
    pe = asp.get("por_empresa") or {}
    _rect(sl, 8.25, 1.4, 4.9, 1.25, "FFF6DD", "E8B400", lw=1.5, redondo=0.06)
    b = _rect(sl, 8.4, 1.55, 1.25, 0.95, DORADO, redondo=0.08)
    _llenar(b.text_frame, [([_run(n_asp, 40, True, "FFFFFF", SERIF)], 0)], align="c", anchor="m")
    _texto(sl, 9.85, 1.55, 3.2, 0.35, [([_run("PROGRAMA ASP · SOCIO PREVENTIVO", 13, True, "A86E00")], 0)])
    _texto(sl, 9.85, 1.9, 3.2, 0.65, [([_run(
        f"{n_asp} {'reporte' if n_asp == 1 else 'reportes'} ASP en GPRO este turno (por título).  "
        f"COX {pe.get('COX S.A', 0)} · PUMA {pe.get('PUMA', 0)}.", 10, False, "333333")], 0)])

    # Autorreportabilidad
    au = d.get("autorreporte") or {}
    _rect(sl, 8.25, 2.8, 4.9, 1.25, "EAF2F4", TEAL, lw=1.5, redondo=0.06)
    b = _rect(sl, 8.4, 2.95, 1.25, 0.95, TEAL, redondo=0.08)
    _llenar(b.text_frame, [([_run(au.get("valor", "—"), 32, True, "FFFFFF", SERIF)], 0)], align="c", anchor="m")
    _texto(sl, 9.85, 2.95, 3.2, 0.55, [([_run(au.get("titulo", "AUTORREPORTABILIDAD EECC"), 13, True, AZUL)], 0)],
           anchor="m")
    _texto(sl, 9.85, 3.5, 3.2, 0.45, [([_run(au.get("detalle", ""), 9.5, False, GRIS_TXT)], 0)])

    # Tabla dinámica de entrega de turno
    banda = _rect(sl, 0.19, 4.25, 12.96, 0.55, AZUL, redondo=0.08)
    _llenar(banda.text_frame, [([_run("TABLA DINÁMICA DE ENTREGA DE TURNO — FOCOS PRIORITARIOS DE ALTO POTENCIAL",
                                       14, True, "FFFFFF", SERIF)], 0)], anchor="m", margen_izq=0.2)
    items = d.get("items") or []
    if not items:
        _texto(sl, 0.19, 5.2, 12.96, 0.6, [([_run("Sin focos de alto potencial asociados a EDC en el turno.", 12, False,
                                                  GRIS_TXT)], 0)], align="c", anchor="m")
    for i, it in enumerate(items[:4]):
        y, h = 4.95 + i * 0.95, 0.85
        col = _col(it)
        _rect(sl, 0.19, y, 12.96, h, FONDO_FILA, BORDE, redondo=0.1)
        _rect(sl, 0.19, y, 0.13, h, col)
        bd = _rect(sl, 0.42, y + 0.1, 1.25, h - 0.2, col, redondo=0.1)
        _llenar(bd.text_frame, [([_run(it["tipo"], 17, True, "FFFFFF", SERIF)], 0),
                                ([_run(f"ID {it['id']}", 8.5, False, "FFFFFF")], 0)], align="c", anchor="m", margen=0.02)
        _texto(sl, 1.85, y + 0.08, 3.9, 0.4, [([_run(it["titulo"], 12, True, AZUL)], 0)], anchor="m")
        _texto(sl, 1.85, y + 0.48, 3.9, 0.3, [([_run(it["empresa"], 9.5, True, _col(it, False)),
                                                _run(f"  ·  {it['area']}" if it.get("area") else "", 9.5, False,
                                                     GRIS_TXT)], 0)], anchor="m")
        _texto(sl, 5.85, y + 0.08, 3.4, h - 0.16, [([_run("EDC  ", 9.5, True, TEAL),
                                                     _run(it["edc_corto"], 9.5, False, "333333")], 0)], anchor="m")
        _texto(sl, 9.4, y + 0.06, 2.5, h - 0.12, [([_run("Acción  ", 9, True, GRIS_TXT),
                                                    _run(it["accion"], 9, False, "333333")], 0)], anchor="m")
        ab = it.get("abierto")
        est = _rect(sl, 12.05, y + 0.27, 0.95, 0.32, "FDECEC" if ab else "EEF6F7", ROJO if ab else TEAL, redondo=0.5)
        _llenar(est.text_frame, [([_run("ABIERTO" if ab else "CERRADO", 9, True, ROJO if ab else TEAL)], 0)],
                align="c", anchor="m", margen=0.01)
    return sl


def agregar(pptx_bytes, datos):
    """Agrega las 2 láminas al final del PPT y devuelve (bytes, log)."""
    global P
    if P is None:
        P = _pptx()
    prs = P["Presentation"](io.BytesIO(pptx_bytes))
    logo = _logo(prs)
    _lamina_a(prs, datos, logo)
    _lamina_b(prs, datos, logo)
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), [f"✔ Agregadas las láminas «Gestión SSO GPRO» y «Reportabilidad, Autorreportes y Focos de "
                            f"turno» ({len(prs.slides)} láminas en total)"]
