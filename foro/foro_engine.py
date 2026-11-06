"""
Motor de actualización del Foro SSO GPRO.

Recibe la presentación del foro (bytes .pptx) y el JSON de datos semanales
(extraído de Power BI por el agente) y devuelve la presentación actualizada.
Solo toca gráficos, tablas, KPIs y textos de datos; las láminas de reflexión
y cualquier otra lámina quedan intactas.

Corre igual en Python normal y en el navegador (Pyodide).
"""
import copy
import io
import re

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

VERDE, AMBAR, ROJO, GRIS = "3B6E22", "B07A00", "D22630", "6B7780"
FONDO_VERDE, FONDO_AMBAR, FONDO_ROJO, FONDO_GRIS = "E6F1E0", "FFF4DC", "FBEAEB", "EEF1F3"
AZUL_REALIZADA, GRIS_PENDIENTE = "01778B", "C9D0D9"


# ----------------------------------------------------------------- utilidades
def _norm(t):
    return re.sub(r"\s+", " ", (t or "")).strip().upper()


def _iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # grupo
            yield from _iter_shapes(sh.shapes)


def buscar_lamina(prs, *claves):
    """Primera lámina que contiene un texto con alguna de las claves."""
    claves = [_norm(c) for c in claves]
    for s in prs.slides:
        for sh in _iter_shapes(s.shapes):
            if sh.has_text_frame:
                t = _norm(sh.text_frame.text)
                if any(c in t for c in claves):
                    return s
    return None


def forma(slide, nombre):
    for sh in _iter_shapes(slide.shapes):
        if sh.name == nombre:
            return sh
    return None


def forma_texto(slide, inicio):
    inicio = _norm(inicio)
    for sh in _iter_shapes(slide.shapes):
        if sh.has_text_frame and _norm(sh.text_frame.text).startswith(inicio):
            return sh
    return None


def poner_texto(tf, texto, color=None):
    """Reemplaza el texto conservando el formato de cada párrafo.
    `texto` usa '\n' para separar párrafos."""
    if hasattr(tf, "text_frame"):
        tf = tf.text_frame
    lineas = str(texto).split("\n")
    parrafos = list(tf.paragraphs)
    # agregar párrafos copiando el último
    while len(parrafos) < len(lineas):
        nuevo = copy.deepcopy(parrafos[-1]._p)
        parrafos[-1]._p.addnext(nuevo)
        parrafos = list(tf.paragraphs)
    # quitar sobrantes
    for p in parrafos[len(lineas):]:
        p._p.getparent().remove(p._p)
    for p, linea in zip(tf.paragraphs, lineas):
        runs = list(p.runs)
        if not runs:
            r = p.add_run()
            runs = [r]
        runs[0].text = linea
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        if color:
            runs[0].font.color.rgb = RGBColor.from_string(color)


def poner_celda(celda, texto, color=None, fondo=None):
    poner_texto(celda.text_frame, texto, color)
    if fondo:
        celda.fill.solid()
        celda.fill.fore_color.rgb = RGBColor.from_string(fondo)


def ajustar_filas(tabla, n_cuerpo, enc=1, pie=0):
    """Deja la tabla con `n_cuerpo` filas entre encabezado y pie, copiando la
    última fila de cuerpo o eliminando las sobrantes."""
    tbl = tabla._tbl
    filas = tbl.tr_lst
    cuerpo = filas[enc:len(filas) - pie] if pie else filas[enc:]
    while len(cuerpo) < n_cuerpo:
        nueva = copy.deepcopy(cuerpo[-1])
        cuerpo[-1].addnext(nueva)
        cuerpo.append(nueva)
    while len(cuerpo) > n_cuerpo:
        tbl.remove(cuerpo.pop())


def reemplazar_grafico(chart, categorias, series, formato="General"):
    cd = CategoryChartData(number_format=formato)
    cd.categories = categorias
    for nombre, valores in series:
        cd.add_series(nombre, valores)
    chart.replace_data(cd)


def color_pct(p):
    if p is None:
        return GRIS, FONDO_GRIS
    if p >= 100:
        return VERDE, FONDO_VERDE
    if p >= 80:
        return AMBAR, FONDO_AMBAR
    return ROJO, FONDO_ROJO


def fpct(p, dec=0):
    if p is None:
        return "—"
    s = f"{p:.{dec}f}".replace(".", ",")
    return s + "%"


def fvar(v):
    return f"+{v}" if v > 0 else str(v)


def var_pct(ant, act):
    if not ant:
        return "-%" if act else "0%"
    return f"{round((act - ant) / ant * 100):d}%"


# ------------------------------------------------------------------ láminas
def aplicar_portada(prs, d, log):
    sh = None
    for s in list(prs.slides)[:2]:
        sh = forma_texto(s, "Revisión semanal")
        if sh:
            break
    if not sh:
        return log.append("Portada: no encontré 'Revisión semanal'")
    poner_texto(sh, d["texto"])
    log.append("Portada ✔")


def aplicar_radar(prs, d, log):
    s = buscar_lamina(prs, "RADAR SSO SEMANAL")
    if not s:
        return log.append("Radar SSO: lámina no encontrada")
    if forma(s, "TextBox 3"):
        poner_texto(forma(s, "TextBox 3"), d["periodo_txt"])
    k = d["kpis"]
    pares = [("TextBox 6", "TextBox 7", k["unicos"], "registro SAP", "registros SAP"),
             ("TextBox 8", "TextBox 9", k["en_temas"], None, None),
             ("TextBox 10", "TextBox 11", k["no_gcc"], None, None),
             ("TextBox 12", "TextBox 13", k["fuera"], None, None)]
    for v, et, n, sing, plur in pares:
        if forma(s, v):
            poner_texto(forma(s, v), str(n))
        if sing and forma(s, et):
            poner_texto(forma(s, et), sing if n == 1 else plur)
    # temas: encabezado -> (+1 conteo, +2 detalle)
    for titulo, t in d["temas"].items():
        enc = forma_texto(s, titulo)
        if not enc:
            log.append(f"Radar: tema '{titulo}' no encontrado")
            continue
        num = int(re.search(r"(\d+)$", enc.name).group(1))
        unidad = "respuesta" if "GCC" in titulo.upper() else "registro"
        cnt = forma(s, f"TextBox {num + 1}")
        det = forma(s, f"TextBox {num + 2}")
        if cnt:
            poner_texto(cnt, f"{t['n']} {unidad}{'' if t['n'] == 1 else 's'}")
        if det:
            poner_texto(det, t["texto"])
    if d.get("pie") and forma(s, "TextBox 45"):
        poner_texto(forma(s, "TextBox 45"), d["pie"])
    log.append("Radar SSO ✔")


def _aplicar_como_vamos(s, d, nombre, log):
    meses = [m.upper() for m in d["meses"]]
    g = forma(s, "Gráfico 21")
    if g and g.has_chart:
        reemplazar_grafico(g.chart, meses, [("CAAP", d["caap"]), ("CABP", d["cabp"])])
    t = forma(s, "Tabla 22")
    if t and t.has_table:
        tb = t.table
        a, b = d["caap"][-2:], d["cabp"][-2:]
        poner_celda(tb.cell(1, 0), d["meses"][-2].capitalize())
        poner_celda(tb.cell(2, 0), d["meses"][-1].capitalize() + ("*" if d.get("parcial") else ""))
        poner_celda(tb.cell(1, 1), str(a[0])); poner_celda(tb.cell(1, 2), str(b[0]))
        poner_celda(tb.cell(2, 1), str(a[1])); poner_celda(tb.cell(2, 2), str(b[1]))
        poner_celda(tb.cell(3, 1), var_pct(a[0], a[1]))
        poner_celda(tb.cell(3, 2), var_pct(b[0], b[1]))
    ind = d.get("indicadores")
    if ind:
        g = forma(s, "Gráfico 23")
        if g and g.has_chart:
            series = [(n, ind["series"].get(n, [None] * len(meses)))
                      for n in ("CTP", "STP TM", "STP PA", "AAP", "TRIFR YTD", "IF YTD", "IAAP")]
            # (los 4 primeros van en barras apiladas, los 3 últimos en líneas)
            reemplazar_grafico(g.chart, meses, series)
        t = forma(s, "Tabla 24")
        if t and t.has_table and ind.get("ytd"):
            y = ind["ytd"]
            num = lambda x: float(str(x).replace(",", ".") or 0)
            fmt = lambda v: (f"{v:.2f}".replace(".", ",") if isinstance(v, float) and v != int(v) else str(int(v)))
            for j, k in enumerate(("IF", "TRIFR", "IG", "AAP"), start=1):
                if "2025" in y:
                    poner_celda(t.table.cell(1, j), fmt(y["2025"][k]))
                ant = num(t.table.cell(1, j).text)
                act = y["2026"][k]
                poner_celda(t.table.cell(2, j), fmt(act))
                poner_celda(t.table.cell(3, j), f"{round((act - ant) / ant * 100)}%" if ant else "0%")
        t = forma(s, "Tabla 25")
        if t and t.has_table and ind.get("abp"):
            for j, k in enumerate(("DM", "STP", "CTP")):
                a0, a1 = ind["abp"]["2025"][k], ind["abp"]["2026"][k]
                poner_celda(t.table.cell(1, j), str(a0))
                poner_celda(t.table.cell(2, j), str(a1))
                poner_celda(t.table.cell(3, j), var_pct(a0, a1))
    if d.get("subtitulo") and forma(s, "Subtitulo"):
        poner_texto(forma(s, "Subtitulo"), d["subtitulo"])
    log.append(f"¿Cómo vamos? {nombre} ✔" + ("" if ind else " (CAAP/CABP; indicadores sin datos)"))


def aplicar_como_vamos(prs, d, log):
    for clave, titulo in (("antucoya", "¿CÓMO VAMOS? – ANTUCOYA"),
                          ("gpro", "¿CÓMO VAMOS? – GERENCIA DE PROYECTOS")):
        if clave not in d:
            continue
        s = buscar_lamina(prs, titulo)
        if not s:
            log.append(f"¿Cómo vamos? {clave}: lámina no encontrada")
            continue
        _aplicar_como_vamos(s, d[clave], clave, log)


def aplicar_como_nos_fue(prs, d, log):
    s = buscar_lamina(prs, "¿CÓMO NOS FUE EN")
    if not s:
        return log.append("¿Cómo nos fue?: lámina no encontrada")
    poner_texto(forma_texto(s, "¿CÓMO NOS FUE EN"), d["titulo"])
    if forma(s, "Subtitulo"):
        poner_texto(forma(s, "Subtitulo"), d["subtitulo"])
    for k, n in d["conteos"].items():
        sh = forma(s, f"Pirámide {k}")
        if sh:
            poner_texto(sh, f"{k}\n{n}")
    for k, txt in d.get("encabezados", {}).items():
        sh = forma(s, f"Encabezado {k}")
        if sh:
            poner_texto(sh, txt)
    for k, txt in d.get("detalles", {}).items():
        sh = forma(s, f"Detalle {k}")
        if sh:
            poner_texto(sh, txt)
    if d.get("puntos_titulo") and forma(s, "Encabezado puntos"):
        poner_texto(forma(s, "Encabezado puntos"), d["puntos_titulo"])
    for i, txt in enumerate(d.get("puntos", []), start=1):
        sh = forma(s, f"Punto {i}")
        if sh:
            poner_texto(sh, txt)
    if d.get("fuente") and forma(s, "Fuente"):
        poner_texto(forma(s, "Fuente"), d["fuente"])
    log.append("¿Cómo nos fue? ✔")


ACT_EECC = ["artp", "cts", "rol", "procesos", "total"]


def aplicar_passo_eecc(prs, d, log):
    s = buscar_lamina(prs, "PASSO — AVANCE SEMANAL EECC", "PASSO - AVANCE SEMANAL EECC")
    if not s:
        return log.append("PASSO EECC: lámina no encontrada")
    for nombre in ("Subtitulo", "Leyenda barras", "Lectura"):
        if d.get(nombre.lower().replace(" ", "_")) and forma(s, nombre):
            poner_texto(forma(s, nombre), d[nombre.lower().replace(" ", "_")])
    t = forma(s, "Mapa de calor PASSO")
    filas = d["empresas"]  # lista de {nombre, artp:[r,p], ...}
    if t and t.has_table:
        ajustar_filas(t.table, len(filas))
        for i, emp in enumerate(filas, start=1):
            poner_celda(t.table.cell(i, 0), emp["nombre"])
            for j, k in enumerate(ACT_EECC, start=1):
                r, p = emp[k]
                if not p:
                    poner_celda(t.table.cell(i, j), "Sin programa\n0 programadas", GRIS, FONDO_GRIS)
                    continue
                pct = round(r / p * 100)
                col, fon = color_pct(pct)
                poner_celda(t.table.cell(i, j), f"{pct}%\n{r}/{p}", col, fon)
    g = forma(s, "Chart")
    if g and g.has_chart:
        vals = [round(e["total"][0] / e["total"][1] * 100, 1) if e["total"][1] else 0 for e in filas]
        reemplazar_grafico(g.chart, [e["nombre"] for e in filas], [("", vals)], "0.0")
    log.append("PASSO EECC ✔")


def aplicar_passo_ant(prs, d, log):
    s = buscar_lamina(prs, "PASSO — CUMPLIMIENTO ANTUCOYA", "PASSO - CUMPLIMIENTO ANTUCOYA")
    if not s:
        return log.append("PASSO Antucoya: lámina no encontrada")
    for nombre, clave in (("Titulo", "titulo"), ("Subtitulo", "subtitulo"), ("Grupo barras", "grupo_barras"),
                          ("Leyenda barras", "leyenda_barras"), ("Lectura", "lectura"),
                          ("Nota personas por rol", "nota_roles")):
        if d.get(clave) and forma(s, nombre):
            poner_texto(forma(s, nombre), d[clave])
    acts = d["actividades"]  # [{nombre, kpi, aplic_ant, aplic}]
    g = forma(s, "Grafico cumplimiento PASSO")
    if g and g.has_chart:
        reemplazar_grafico(g.chart, [a["nombre"] for a in acts], [("Realizadas", [a["kpi"] for a in acts])])
    t = forma(s, "Mapa de calor PASSO")
    if t and t.has_table:
        cap = lambda v, k: min(v, k) if k else v
        tot_kpi = sum(a["kpi"] for a in acts)
        tot_ant = sum(cap(a["aplic_ant"], a["kpi"]) for a in acts)
        tot_act = sum(cap(a["aplic"], a["kpi"]) for a in acts)
        ajustar_filas(t.table, len(acts), pie=1)
        filas = acts + [{"nombre": "KPI mensual actividades", "kpi": tot_kpi, "aplic_ant": tot_ant, "aplic": tot_act}]
        for i, a in enumerate(filas, start=1):
            pct = min(100, round(a["aplic"] / a["kpi"] * 100)) if a["kpi"] else None
            var = a["aplic"] - a["aplic_ant"]
            peso = round(a["kpi"] / tot_kpi * 100) if tot_kpi else 0
            col, _ = color_pct(pct)
            vals = [a["nombre"], str(a["kpi"]), str(a["aplic_ant"]), str(a["aplic"]),
                    fpct(pct), fvar(var), f"{peso}%"]
            for j, v in enumerate(vals):
                poner_celda(t.table.cell(i, j), v, col if j == 4 else (VERDE if j == 5 and var > 0 else None))
    log.append("PASSO Antucoya ✔")


def aplicar_passo_persona(prs, d, log):
    s = buscar_lamina(prs, "PASSO — CUMPLIMIENTO POR PERSONA", "PASSO - CUMPLIMIENTO POR PERSONA")
    if not s:
        return log.append("PASSO por persona: lámina no encontrada")
    for nombre, clave in (("Subtitulo", "subtitulo"), ("Grupo barras", "grupo_barras"), ("Lectura", "lectura")):
        if d.get(clave) and forma(s, nombre):
            poner_texto(forma(s, nombre), d[clave])
    t = forma(s, "Mapa de calor PASSO")
    if t and t.has_table:
        tb = t.table
        meses = d["meses"]  # ej. ["JUL","AGO","SEP","OCT*"]
        for j, m in enumerate(meses, start=2):
            poner_celda(tb.cell(0, j), f"% {m}")
        personas = d["personas"]  # [{nombre, rol, pct:[..]}]
        # conservar la ortografía/rol con que la lámina ya nombra a cada persona
        import unicodedata
        sin = lambda x: unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode().lower().split()
        previos = {}
        for r in list(tb.rows)[1:-1]:
            nom = r.cells[0].text.strip()
            if nom:
                previos[tuple(sin(nom)[-2:])] = (nom, r.cells[1].text.strip())
        for p in personas:
            prev = previos.get(tuple(sin(p["nombre"])[-2:]))
            if prev:
                p["nombre"], p["rol"] = prev
        ajustar_filas(tb, len(personas), pie=1)
        filas = personas + [{"nombre": "Total Proyectos", "rol": f"{len(personas)} personas", "pct": d["total"]}]
        for i, p in enumerate(filas, start=1):
            poner_celda(tb.cell(i, 0), p["nombre"])
            poner_celda(tb.cell(i, 1), p["rol"])
            for j, v in enumerate(p["pct"], start=2):
                poner_celda(tb.cell(i, j), fpct(v), color_pct(v)[0])
    log.append("PASSO por persona ✔")


def aplicar_gcc(prs, d, log):
    s = buscar_lamina(prs, "GESTIÓN EDC MEDIANTE GCC", "GESTION EDC MEDIANTE GCC")
    if not s:
        return log.append("GCC: lámina no encontrada")
    # roles: [{"rotulo": "Rol 1", "titulo": "Gerencia de Proyectos", "edc": {"Veh. livianos": true,...},
    #          "grafico_id": 6056, "etiquetas": [ids...]}]
    for rol in d["roles"]:
        hechas = sum(1 for v in rol["edc"].values() if v)
        total = len(rol["edc"])
        if forma(s, rol["rotulo"]):
            poner_texto(forma(s, rol["rotulo"]), f"{rol['titulo']}  ·  {hechas} de {total}")
        # gráfico: puntos pares = EdC, impares = separador
        graf = next((sh for sh in s.shapes if sh.shape_id == rol["grafico_id"]), None)
        if graf is not None and graf.has_chart:
            ser = graf.chart.plots[0].series[0]._element
            dpts = {int(x.find(C + "idx").get("val")): x for x in ser.findall(C + "dPt")}
            for i, (edc, ok) in enumerate(rol["edc"].items()):
                dp = dpts.get(2 * i)
                if dp is None:
                    continue
                clr = dp.find(".//" + A + "srgbClr")
                if clr is not None:
                    clr.set("val", AZUL_REALIZADA if ok else GRIS_PENDIENTE)
        # etiquetas
        for sid in rol.get("etiquetas", []):
            sh = next((x for x in s.shapes if x.shape_id == sid), None)
            if sh is None or not sh.has_text_frame:
                continue
            nombre = re.sub(r"\s*\d+/\d+$", "", sh.text_frame.text.strip())
            if nombre in rol["edc"]:
                poner_texto(sh, f"{nombre} {1 if rol['edc'][nombre] else 0}/1")
    log.append("GCC ✔")


# ------------------------------------------------------------- roles del foro
def _roles_desde_programa(prs, fecha):
    """Lee la tabla 'PROGRAMA PRÓXIMOS LÍDERES' del mismo foro y devuelve los
    roles de la fecha (dd-mm-aaaa)."""
    s = buscar_lamina(prs, "PROXIMOS LIDERES", "PRÓXIMOS LÍDERES")
    if not s:
        return None
    t = next((sh for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table), None)
    if t is None:
        return None
    filas = [[c.text.strip() for c in r.cells] for r in t.table.rows]
    enc = [_norm(x) for x in filas[0]]
    col = lambda clave: next((j for j, h in enumerate(enc) if clave in h), None)
    out = {}
    for rol, clave in (("lider", "FECHA"), ("timekeeper", "TIME KEEPER"), ("compromisos", "COMPROMISOS"),
                       ("feedback", "CONFIRMACI")):
        j = col(clave)
        if j is None:
            continue
        for f in filas[1:]:
            if len(f) > j and f[j] == fecha:
                out[rol] = f[0]
                break
    return out


def aplicar_roles(prs, d, log):
    r = _roles_desde_programa(prs, d["fecha"])
    if not r:
        return log.append("Roles: no encontré el programa de líderes o la fecha " + d["fecha"])
    etiquetas = {"lider": ("1. LÍDER", "LÍDER"), "timekeeper": ("2. TIME KEEPER", "TIMEKEEPER"),
                 "compromisos": ("4. REGISTRA COMPROMISOS", "REGISTRA COMPROMISOS"),
                 "feedback": ("5. FEEDBACK", "FEEDBACK")}
    hechos = 0
    for s in prs.slides:
        textos = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
        for rol, nombre in r.items():
            for lab in etiquetas[rol]:
                et = next((sh for sh in textos if _norm(sh.text_frame.text) == lab), None)
                if et is None:
                    continue
                # nombre en la misma fila (a la derecha); si no hay, en la misma columna (debajo)
                excl = {_norm(x) for v in etiquetas.values() for x in v} | {"TODOS", "ROL", "PARTICIPANTES"}
                ok = lambda sh: (sh is not et and len(sh.text_frame.text) < 40 and
                                 _norm(sh.text_frame.text) not in excl and
                                 not re.match(r"^\d+\.", sh.text_frame.text.strip()))
                cand = [sh for sh in textos if ok(sh) and abs(sh.top - et.top) < 15 * 12700
                        and sh.left > et.left + et.width // 2]
                if not cand:
                    cand = [sh for sh in textos if ok(sh) and abs(sh.left - et.left) < 15 * 12700
                            and 0 < sh.top - et.top < 90 * 12700]
                if cand:
                    dest = min(cand, key=lambda sh: abs(sh.top - et.top) + abs(sh.left - et.left))
                    poner_texto(dest, nombre)
                    hechos += 1
                break
    log.append(f"Roles del foro ✔ (Líder {r.get('lider', '—')}, Time keeper {r.get('timekeeper', '—')}, "
               f"Compromisos {r.get('compromisos', '—')}, Feedback {r.get('feedback', '—')})" if hechos
               else "Roles: no encontré dónde escribirlos")


# ------------------------------------------------------- calendarios EDC
MESES_NOM = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
             "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def _sin_tildes(x):
    import unicodedata
    return unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode().lower()


def _eliminar_lamina(prs, slide):
    sldIdLst = prs.slides._sldIdLst
    for sid in list(sldIdLst):
        if prs.slides.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId)
            sldIdLst.remove(sid)
            return True
    return False


def aplicar_calendario(prs, d, log):
    """Quita del calendario las visitas de EDC ya completadas (según GCC) del
    dueño/verificador correspondiente, desde la fecha del foro en adelante."""
    import datetime as _dt
    desde = _dt.date.fromisoformat(d["fecha"])
    hechas = {_sin_tildes(k): {str(x) for x in v} for k, v in d["hechas"].items()}
    quitadas = []
    for s in list(prs.slides):
        tit = forma_texto(s, "EDC — CALENDARIO DE VERIFICACIONES") or forma_texto(s, "EDC - CALENDARIO DE VERIFICACIONES")
        if not tit:
            continue
        m = re.search(r"(" + "|".join(MESES_NOM) + r")\s+(\d{4})", _norm(tit.text_frame.text))
        if not m:
            continue
        mes, anio = MESES_NOM.index(m.group(1)) + 1, int(m.group(2))
        tabla = next((sh.table for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table), None)
        if tabla is None:
            continue
        filas = [list(r.cells) for r in list(tabla.rows)[1:]]
        celdas = [c for f in filas for c in f]
        planas = {}  # columna -> celda de día sin actividad (para copiar su formato)
        for f in filas:
            for j, c in enumerate(f):
                if re.fullmatch(r"\d{1,2}", c.text.strip()):
                    planas.setdefault(j, c)
        col_de = {id(c._tc): j for f in filas for j, c in enumerate(f)}
        for c in celdas:
            j = col_de[id(c._tc)]
            modelo = planas.get(j) or next((planas[k] for k in range(5) if k in planas), None)
            txt = c.text.strip()
            dia = re.match(r"^(\d{1,2})", txt)
            if not dia or "EDC" not in txt.upper() or modelo is None:
                continue
            try:
                fecha = _dt.date(anio, mes, int(dia.group(1)))
            except ValueError:
                continue
            if fecha < desde:
                continue
            edc = re.search(r"EDC\s*(\d+(?:\.\d+)?)", txt, re.I)
            if not edc:
                continue
            personas = _sin_tildes(txt.split("Con:")[0])
            dueños = [p for p in hechas if p in personas]
            if dueños and all(edc.group(1) in hechas[p] for p in dueños):
                # dejar la celda como un día sin actividad (formato de una celda vacía)
                nuevo_tc = copy.deepcopy(modelo._tc)
                c._tc.getparent().replace(c._tc, nuevo_tc)
                from pptx.table import _Cell
                poner_texto(_Cell(nuevo_tc, c._parent).text_frame, dia.group(1))
                quitadas.append(f"{fecha.strftime('%d-%m')} EDC {edc.group(1)}")
    # la lámina de verificaciones de la semana ya no se usa
    for s in list(prs.slides):
        if forma_texto(s, "EDC — VERIFICACIONES DE LA SEMANA") or forma_texto(s, "EDC - VERIFICACIONES DE LA SEMANA"):
            _eliminar_lamina(prs, s)
            log.append("Lámina 'EDC — Verificaciones de la semana' eliminada")
    log.append("Calendario EDC ✔ " + (f"(quitadas por GCC completada: {', '.join(quitadas)})" if quitadas
                                      else "(sin visitas que quitar)"))


def aplicar_passo_persona_mes(prs, d, log):
    """Actualiza solo la columna del mes en curso; los meses anteriores quedan como
    los dejó SSO (si cambia el mes, las columnas se corren a la izquierda)."""
    import unicodedata
    s = buscar_lamina(prs, "PASSO — CUMPLIMIENTO POR PERSONA", "PASSO - CUMPLIMIENTO POR PERSONA")
    if not s:
        return log.append("PASSO por persona: lámina no encontrada")
    t = forma(s, "Mapa de calor PASSO")
    if t is None or not t.has_table:
        return log.append("PASSO por persona: tabla no encontrada")
    tb = t.table
    sin = lambda x: unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode().lower().split()
    clave = lambda nom: tuple(sin(nom)[-2:])
    ncol = len(tb.columns)
    meses = [tb.cell(0, j).text.replace("%", "").replace("*", "").strip().upper() for j in range(2, ncol)]
    filas = list(tb.rows)
    previos = {}  # clave -> (nombre, rol, [valores])
    for r in filas[1:-1]:
        nom = r.cells[0].text.strip()
        if nom:
            previos[clave(nom)] = (nom, r.cells[1].text.strip(), [r.cells[j].text.strip() for j in range(2, ncol)])
    tot_prev = [filas[-1].cells[j].text.strip() for j in range(2, ncol)]
    mes = d["mes"].upper()
    correr = meses[-1] != mes
    if correr:
        meses = meses[1:] + [mes]
    personas = []
    for p in d["personas"]:
        prev = previos.get(clave(p["nombre"]))
        nombre, rol, vals = prev if prev else (p["nombre"], p["rol"], ["—"] * len(meses))
        vals = (vals[1:] + [None]) if correr else vals[:-1] + [None]
        vals[-1] = p["pct"]
        personas.append((nombre, rol, vals))
    personas.sort(key=lambda x: (-(x[2][-1] or 0), x[0]))
    tot = (tot_prev[1:] + [None]) if correr else tot_prev[:-1] + [None]
    tot[-1] = d["total"]
    for j, m in enumerate(meses, start=2):
        poner_celda(tb.cell(0, j), f"% {m}" + ("*" if j == ncol - 1 else ""))
    ajustar_filas(tb, len(personas), pie=1)
    num = lambda v: None if v in (None, "—", "") else (v if isinstance(v, (int, float)) else int(str(v).replace("%", "") or 0))
    fmt = lambda v: v if isinstance(v, str) else fpct(v)
    for i, (nombre, rol, vals) in enumerate(personas, start=1):
        poner_celda(tb.cell(i, 0), nombre)
        poner_celda(tb.cell(i, 1), rol)
        for j, v in enumerate(vals, start=2):
            poner_celda(tb.cell(i, j), fmt(v), color_pct(num(v))[0])
    n = len(personas)
    poner_celda(tb.cell(n + 1, 1), f"{n} personas")
    for j, v in enumerate(tot, start=2):
        poner_celda(tb.cell(n + 1, j), fmt(v))
    en100 = sum(1 for p in personas if (p[2][-1] or 0) >= 100)
    if forma(s, "Grupo barras"):
        poner_texto(forma(s, "Grupo barras"), f"CUMPLIMIENTO PASSO POR PERSONA  ·  Proyectos  ·  *{d['mes_nombre']} en curso")
    if forma(s, "Lectura"):
        poner_texto(forma(s, "Lectura"), f"Reconocimiento y compromiso\nEn {d['mes_nombre']} el cumplimiento va en "
                    f"{d['total']}%: {en100} de {n} personas en 100%. Cerremos las actividades a tiempo; "
                    f"el liderazgo visible en terreno depende de cada uno.")
    log.append(f"PASSO por persona ✔ (mes en curso {mes}" + (", columnas corridas)" if correr else ")"))


def aplicar_gcc_imagen(prs, b64png, log):
    """Reemplaza la imagen de la matriz EDC (lámina con la tabla 'N° EdC')."""
    import base64
    s = next((sl for sl in prs.slides if any(sh.name == "Tabla EdC" for sh in sl.shapes)), None)
    if s is None:
        return log.append("Matriz EDC: lámina no encontrada")
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    if not pics:
        return log.append("Matriz EDC: imagen no encontrada")
    old = max(pics, key=lambda p: p.width * p.height)
    data = base64.b64decode(b64png)
    from PIL import Image
    w, h = Image.open(io.BytesIO(data)).size
    ancho = old.width
    alto = int(ancho * h / w)
    top = old.top + (old.height - alto) // 2
    nueva = s.shapes.add_picture(io.BytesIO(data), old.left, top, ancho, alto)
    old._element.addprevious(nueva._element)  # misma posición en el orden de capas
    old._element.getparent().remove(old._element)
    log.append("Matriz EDC (imagen) ✔")


APLICADORES = [
    ("portada", aplicar_portada),
    ("roles", aplicar_roles),
    # ("radar", aplicar_radar),  # lámina enviada por SSO: no se modifica
    ("como_vamos", aplicar_como_vamos),
    ("como_nos_fue", aplicar_como_nos_fue),
    ("passo_eecc", aplicar_passo_eecc),
    ("passo_ant", aplicar_passo_ant),
    ("passo_persona", aplicar_passo_persona_mes),  # solo el mes en curso
    ("gcc", aplicar_gcc),
    ("calendario", aplicar_calendario),
    ("gcc_imagen", aplicar_gcc_imagen),
]


def actualizar(pptx_bytes, datos):
    """Devuelve (bytes_pptx_actualizado, lista_de_mensajes)."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    log = []
    for clave, fn in APLICADORES:
        if clave not in datos:
            continue
        try:
            fn(prs, datos[clave], log)
        except Exception as e:  # una lámina con problemas no bloquea las demás
            log.append(f"⚠ {clave}: {type(e).__name__}: {e}")
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), log
