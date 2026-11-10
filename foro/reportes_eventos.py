"""
Láminas «REPORTES POR EMPRESA» y «REPORTES Y GRAVEDAD POR EMPRESA» del foro.

Entrada: el Excel de eventos (hoja «Export», el mismo «data al DD-MM assurance.xlsx» que usa Assurance).
- calcular(excel_bytes, nombre=None) -> dict `d` con todos los números.
- (los textos y la aplicación a las láminas están más abajo)

Corre en Python normal y en el navegador (Pyodide): solo stdlib y openpyxl.
"""
import io
import re
import unicodedata
from collections import Counter
from datetime import date, datetime

VERSION = "1 (10-11 láminas reportes por empresa)"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]
CORTO = [m[:3].upper() for m in MESES]
G3 = ["JEJ", "COX", "PUMA"]          # orden de filas de la lámina 15
TIPOS = ["ACC", "CAP", "HAP", "CBP", "HBP"]   # orden de filas de la tabla de la lámina 16
CODIGOS = ("HBP", "HAP", "CBP", "CAP", "ACC")


def _norm(t):
    t = unicodedata.normalize("NFD", str(t or ""))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").strip().upper()


def _fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19] if "%H" in f else s[:10], f).date()
        except ValueError:
            pass
    return None


def _grupo(emp):
    n = _norm(emp)
    if n.startswith("PUMA"):
        return "PUMA"
    if n.startswith("COX"):
        return "COX"
    if n.startswith("J.E.J") or n.startswith("JEJ"):
        return "JEJ"
    return None


def _nombre_otra(emp):
    return re.sub(r"\s+", " ", _norm(emp)) or "(SIN EMPRESA)"


def _tipo_excel(cl, pot):
    c, p = _norm(cl), _norm(pot)
    if c == "ACCIDENTE":
        return "ACC"
    if c == "CUASI ACCIDENTE":
        return "CAP" if p == "ALTO POTENCIAL" else "CBP"
    if c == "HALLAZGO":
        return "HAP" if p == "ALTO POTENCIAL" else "HBP"
    return None


def _codigo_titulo(t):
    m = re.match(r"^\W*(HBP|HAP|CBP|CAP|ACC)(?![A-Z])", _norm(t))
    return m.group(1) if m else None


def _r(x):
    """Redondeo half-up (como Excel/Power BI), no el del banquero."""
    return int(x + 0.5)


def leer_eventos(excel_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    hoja = next((n for n in wb.sheetnames if _norm(n) == "EXPORT"), None)
    if hoja is None:
        raise ValueError(f"el Excel no tiene hoja 'Export' (hojas: {wb.sheetnames})")
    filas = wb[hoja].iter_rows(values_only=True)
    enc = [_norm(c) for c in next(filas, [])]
    pedidas = {"fecha": ("FECHA EVENTO", "FECHA"), "emp": ("EMPRESA",), "cl": ("CLASIFICACION", "CLASIFICACION SECUNDARIA"),
               "pot": ("POTENCIALIDAD",), "tit": ("TITULO", "TITULO EVENTO"), "riesgo": ("RIESGO INVOLUCRADO", "RIESGO")}
    idx = {}
    for k, nombres in pedidas.items():
        for n in nombres:
            if n in enc:
                idx[k] = enc.index(n)
                break
    falt = [k for k in pedidas if k not in idx and k != "riesgo"]
    if falt:
        raise ValueError(f"faltan columnas en la hoja Export: {falt}")
    out = []
    for f in filas:
        if not f or all(v is None or str(v).strip() == "" for v in f):
            continue
        g = lambda k: f[idx[k]] if (k in idx and idx[k] < len(f)) else None
        out.append({"fecha": _fecha(g("fecha")), "emp": str(g("emp") or "").strip(), "cl": g("cl"), "pot": g("pot"),
                    "tit": str(g("tit") or "").strip(), "riesgo": str(g("riesgo") or "").strip() if "riesgo" in idx else None})
    return out


def calcular(excel_bytes, nombre=None, anio=None):
    """Devuelve el diccionario `d` con todo lo que necesitan las dos láminas."""
    ev = [e for e in leer_eventos(excel_bytes) if e["fecha"] and _tipo_excel(e["cl"], e["pot"])]
    if not ev:
        raise ValueError("el Excel no tiene eventos con fecha y clasificación")
    if anio is None:
        anio = max(e["fecha"] for e in ev).year
    ev = [e for e in ev if e["fecha"].year == int(anio)]
    corte = max(e["fecha"] for e in ev)
    mes_max = corte.month
    meses = list(range(1, mes_max + 1))
    mensual = {g: [0] * mes_max for g in G3}
    ev15 = {g: {c: 0 for c in CODIGOS} for g in G3}
    grav = {t: {"PUMA": 0, "COX": 0, "JEJ": 0, "OTRAS": 0} for t in TIPOS}
    otras = Counter()
    acc = []
    riesgos = {"PUMA": {}, "COX": {}, "RESTO": {}}
    hay_riesgo = any(e["riesgo"] is not None for e in ev)
    for e in ev:
        g = _grupo(e["emp"])
        if hay_riesgo:
            rg = riesgos[g if g in ("PUMA", "COX") else "RESTO"]
            rr = rg.setdefault(e["riesgo"] or "(sin riesgo)", Counter())
            rr[_tipo_excel(e["cl"], e["pot"])] += 1
        t = _tipo_excel(e["cl"], e["pot"])
        grav[t][g or "OTRAS"] += 1
        if g:
            mensual[g][e["fecha"].month - 1] += 1
            c = _codigo_titulo(e["tit"])
            if c:
                ev15[g][c] += 1
        else:
            otras[_nombre_otra(e["emp"])] += 1
        if t == "ACC":
            acc.append({"mes": e["fecha"].month, "empresa": g or _nombre_otra(e["emp"]),
                        "pot": "Alto Potencial" if _norm(e["pot"]) == "ALTO POTENCIAL" else "Bajo Potencial"})
    total_mes = [sum(mensual[g][i] for g in G3) for i in range(mes_max)]
    pct = {g: [(_r(100 * mensual[g][i] / total_mes[i]) if total_mes[i] else 0) for i in range(mes_max)] for g in G3}
    tot_g = {g: sum(mensual[g]) for g in G3}
    fuente = None
    m = re.search(r"(\d{1,2})[-_.](\d{1,2})", nombre or "")
    if m and 1 <= int(m.group(2)) <= 12:
        fuente = f"{int(m.group(1))} {MESES[int(m.group(2)) - 1][:3]} {anio}"
    return {
        "anio": int(anio), "mes_max": mes_max, "dia_corte": corte.day, "meses": meses,
        "g3": G3, "mensual": mensual, "total_mes": total_mes, "pct": pct, "tot_g": tot_g, "tot3": sum(tot_g.values()),
        "ev": ev15, "grav": grav,
        "riesgos": ({k: {r: dict(c) for r, c in v.items()} for k, v in riesgos.items()} if hay_riesgo else None),
        "otras": sorted(otras.items(), key=lambda kv: (-kv[1], kv[0])),
        "acc": sorted(acc, key=lambda a: (a["mes"], _orden_emp(a["empresa"]))),
        "fuente": fuente or f"{corte.day} {MESES[corte.month - 1][:3]} {anio}",
    }


# ------------------------------------------------------------------ textos
# (redacción generada con OmniRoute y corregida/validada contra los textos reales del foro 12-11)
def _miles(n):
    return f"{n:,}".replace(",", ".")


def _pc(n, total):
    return _r(100 * n / total) if total else 0


def _lista(items):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " y " + items[-1]


def _orden_emp(e):
    """PUMA, COX y JEJ primero (como en las láminas); el resto alfabético."""
    return (G3[::-1].index(e) if e in G3 else 9, e)


def textos(d):
    mn = d["mes_max"] - 1
    mes_max = MESES[mn]
    tot3 = d["tot3"]
    gs = sorted(d["g3"], key=lambda g: -d["tot_g"][g])
    G1, G2, G3_ = gs
    p = {g: _pc(d["tot_g"][g], tot3) for g in gs}
    # pico y último mes COMPLETO (el último mes con datos es parcial)
    comp = d["total_mes"][:mn]
    if comp:
        mx = max(comp)
        ip = comp.index(mx)
        ut = comp[-1]
        mes_p, mes_u = MESES[ip], MESES[mn - 1]
        verbo, corto = (("bajaron", "bajó") if ut < mx else ("subieron", "subió") if ut > mx else ("se mantuvieron", "se mantuvo"))
    else:
        mx = ut = ip = 0
        mes_p = mes_u = verbo = corto = ""
    ev = {c: sum(d["ev"][g][c] for g in d["g3"]) for c in CODIGOS}
    otros = ev["HAP"] + ev["CBP"] + ev["CAP"] + ev["ACC"]
    jej = d["mensual"]["JEJ"]
    jm = [MESES[i] for i, v in enumerate(jej) if v > 0]
    jv = [v for v in jej if v > 0]
    acc = d["acc"]
    na = len(acc)
    nA = sum(1 for a in acc if a["pot"] == "Alto Potencial")
    nB = na - nA
    # ---------- lámina 15
    l15 = [(f"{G1} recibe {p[G1]}% de los reportes: ",
            f"{d['tot_g'][G1]} de {_miles(tot3)} entre enero y {mes_max}. {G2} recibe {p[G2]}% ({d['tot_g'][G2]}) y {G3_} {p[G3_]}% ({d['tot_g'][G3_]}).")]
    if comp:
        if verbo == "se mantuvieron":
            ini = f"Los reportes mensuales se mantuvieron en {mx}: "
        else:
            ini = f"Los reportes mensuales {verbo} de {mx} en {mes_p} a {ut} en {mes_u}: "
        l15.append((ini, f"{G1} pasó de {d['mensual'][G1][ip]} a {d['mensual'][G1][mn - 1]} y {G2} de {d['mensual'][G2][ip]} a {d['mensual'][G2][mn - 1]}."))
    l15.append((f"Los eventos HBP son {_pc(ev['HBP'], tot3)}% del total ({_miles(ev['HBP'])}): ",
                f"los eventos CBP, CAP, ACC y HAP suman {_miles(otros)} ({_pc(otros, tot3)}%), con HAP {ev['HAP']}, CBP {ev['CBP']} y CAP {ev['CAP']}. "
                + (f"ACC {ev['ACC']}." if ev["ACC"] else "ACC no registra eventos.")))
    if not jm:
        l15.append((f"JEJ no recibió reportes entre enero y {mes_max}.", ""))
    else:
        rango = (f"con {jv[0]} reporte{'s' if jv[0] != 1 else ''} por mes" if min(jv) == max(jv) else f"con {min(jv)} a {max(jv)} reportes por mes")
        l15.append((f"JEJ recibió reportes solo en {len(jm)} de {d['mes_max']} meses: ", f"{_lista(jm)}, {rango}."))
    if not acc:
        l15.append(("Sin accidentes: ", f"no hay accidentes entre enero y {mes_max}."))
    else:
        por_mes = {}
        for a in acc:
            por_mes.setdefault(a["mes"], []).append(a["empresa"])
        det = [f"{MESES[m - 1]} ({len(es)}: {_lista(sorted(set(es), key=_orden_emp))})" for m, es in sorted(por_mes.items())]
        if na == 1:
            fp = "Es de alto potencial." if nA else "Es de bajo potencial."
        elif nA == 0:
            fp = f"Los {na} son de bajo potencial."
        elif nB == 0:
            fp = f"Los {na} son de alto potencial."
        else:
            fp = f"{nA} de alto potencial y {nB} de bajo potencial."
        l15.append((f"{na} accidente{'s' if na != 1 else ''} en {len(por_mes)} mes{'es' if len(por_mes) != 1 else ''}, marcados en rojo: ", f"{_lista(det)}. {fp}"))
    # ---------- lámina 16
    gr = d["grav"]
    cols = list(gr["ACC"].keys())
    ct = {c: sum(gr[t][c] for t in gr) for c in cols}
    ca = {c: gr["CAP"][c] + gr["HAP"][c] for c in cols}
    talto, ttot = sum(ca.values()), sum(ct.values())
    gx = max(cols, key=lambda c: (ca[c] / ct[c]) if ct[c] else -1)
    l16 = [(f"{G1} recibe {p[G1]}% de los reportes ({d['tot_g'][G1]} de {_miles(tot3)}): ",
            f"el total mensual {corto} de {mx} en {mes_p} a {ut} en {mes_u}." if corto != "se mantuvo" else f"el total mensual se mantuvo en {mx}.")]
    if not acc:
        l16.append(("Sin accidentes: ", f"ninguno entre enero y {mes_max}."))
    else:
        if na == 1:
            pre = "1 accidente, de alto potencial: " if nA else "1 accidente, de bajo potencial: "
        elif nA == 0:
            pre = f"{na} accidentes, todos de bajo potencial: "
        elif nB == 0:
            pre = f"{na} accidentes, todos de alto potencial: "
        else:
            pre = f"{na} accidentes, {nA} de alto y {nB} de bajo potencial: "
        cnt, meses_e = {}, {}
        for a in acc:
            cnt[a["empresa"]] = cnt.get(a["empresa"], 0) + 1
            meses_e.setdefault(a["empresa"], set()).add(a["mes"])
        det = [f"{e} {cnt[e]} ({_lista(MESES[m - 1] for m in sorted(meses_e[e]))})" for e in sorted(cnt, key=lambda e: (-cnt[e], _orden_emp(e)))]
        l16.append((pre, _lista(det) + "."))
    l16.append((f"Alto potencial suma {talto} de {_miles(ttot)} eventos ({_pc(talto, ttot)}%): ",
                f"{'Otras' if gx == 'OTRAS' else gx} tiene la mayor proporción, {ca[gx]} de {ct[gx]} ({_pc(ca[gx], ct[gx])}%)."))
    hb = {g: gr["HBP"][g] for g in d["g3"]}
    h1, h2 = sorted(hb, key=lambda g: (-hb[g], g))[:2]
    thb = sum(gr["HBP"].values())
    l16.append((f"HBP es {_pc(thb, ttot)}% de los eventos ({_miles(thb)}): ", f"{h1} {hb[h1]} y {h2} {hb[h2]}."))
    l16.append(("JEJ reportó todos los meses.", "") if len(jm) == d["mes_max"] else (f"JEJ reportó en solo {len(jm)} de {d['mes_max']} meses: ", f"{_lista(jm)}."))
    fuera = sorted({a["empresa"] for a in acc if a["empresa"] not in d["g3"]})
    leyenda = (f"empresa con accidente ({_lista(fuera)} {'no figura' if len(fuera) == 1 else 'no figuran'} en la tabla)     " if fuera else "empresa con accidente     ")
    return {
        "sub15": f"% de reportes recibidos por cada empresa sobre el total mensual de COX, PUMA y JEJ, enero a {mes_max} {d['anio']}  ·  Eventos acumulados por tipo",
        "sub16": f"Reportes y eventos por empresa, enero a {mes_max} {d['anio']} ({mes_max[:3]}. al {d['dia_corte']})  ·  Accidentes en rojo",
        "escala_corte": (f"{mes_max.capitalize()} ", f"con corte al {d['dia_corte']} de {mes_max}"),
        "nota16": ("acc.: accidente. CAP / CBP: cuasi accidente de alto / bajo potencial. HAP / HBP: hallazgo de alto / bajo potencial. "
                   f"% alto = (CAP + HAP) / total. Otras: {_lista(x[0] for x in d['otras'])}. Fuente: Assurance, {d['fuente']}."),
        "lectura15": l15, "lectura16": l16, "leyenda_empresa": leyenda,
    }


# ------------------------------------------------------------------ aplicación a las láminas
import copy                                   # noqa: E402
from math import ceil                         # noqa: E402
from lxml import etree                        # noqa: E402
from pptx.dml.color import RGBColor           # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
ROJO, TEAL = "C0392B", "108898"
OFFSET_GUIA = 101740   # distancia entre el punto del gráfico y el final de la línea guía (medida en la lámina original)


def _lamina(prs, titulo):
    for s in prs.slides:
        for sh in s.shapes:
            if sh.name == "Titulo" and sh.has_text_frame and _norm(sh.text_frame.text) == _norm(titulo):
                return s
    return None


def _forma(s, nombre):
    for sh in s.shapes:
        if sh.name == nombre:
            return sh
    raise KeyError(f"no encontré la forma «{nombre}»")


def _poner_runs(par, textos_):
    """Pone textos_ en los runs del párrafo (uno por run); si faltan runs, los copia del último."""
    runs = list(par.runs)
    while len(runs) < len(textos_):
        nuevo = copy.deepcopy(runs[-1]._r)
        runs[-1]._r.addnext(nuevo)
        runs = list(par.runs)
    for r, t in zip(runs, textos_):
        r.text = t
    for r in runs[len(textos_):]:
        r._r.getparent().remove(r._r)


def _viñetas(shape, lista):
    """Lectura clave: el primer párrafo es el título; el resto son viñetas (negrita + normal)."""
    pars = list(shape.text_frame.paragraphs)
    while len(pars) - 1 < len(lista):
        nuevo = copy.deepcopy(pars[-1]._p)
        pars[-1]._p.addnext(nuevo)
        pars = list(shape.text_frame.paragraphs)
    for p_ in pars[1 + len(lista):]:
        p_._p.getparent().remove(p_._p)
    for par, (neg, normal) in zip(list(shape.text_frame.paragraphs)[1:], lista):
        _poner_runs(par, [neg, normal])


def _celda(cell, lineas, fondo=None, color=None, negrita=None):
    pars = cell.text_frame.paragraphs
    for i, t in enumerate(lineas):
        runs = pars[i].runs
        runs[0].text = t
        if color:
            runs[0].font.color.rgb = RGBColor.from_string(color)
        if negrita is not None and i == 0:
            runs[0].font.bold = negrita
    if fondo:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(fondo)


def _borde(cell, lado, w, color):
    pr = cell._tc.tcPr
    el = pr.find(A + lado) if pr is not None else None
    if el is None:
        return
    el.set("w", str(w))
    clr = el.find(f"{A}solidFill/{A}srgbClr")
    if clr is not None:
        clr.set("val", color)


def _estilo_pct(p):
    if p >= 60:
        return "108898", "FFFFFF", True
    if p >= 40:
        return "4EC7BE", "0B4F58", True
    if p >= 20:
        return "A8E2DE", "0B4F58", False
    if p >= 1:
        return "DDF3F1", "0B4F58", False
    return "EEF1F3", "8A97A0", False


def _desglose(e):
    partes = [(k, e[k]) for k in ("HAP", "CBP", "CAP", "ACC") if e[k]]
    partes.sort(key=lambda kv: -kv[1])
    return " · ".join(f"{k} {v}" for k, v in partes)


def _lamina15(prs, d, T, log):
    s = _lamina(prs, "REPORTES POR EMPRESA")
    if s is None:
        return log.append("⚠ Reportes por empresa: no encontré la lámina «REPORTES POR EMPRESA»")
    import foro_engine
    tabla = next(sh for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table).table
    enc = [c.text.strip().upper() for c in tabla.rows[0].cells]
    i_hbp = enc.index("EVENTOS HBP")
    n = len(d["meses"])
    foro_engine._ajustar_columnas_meses(tabla, 1, i_hbp - 1, n)
    i_hbp = 1 + n
    acc_meses = {a["mes"] for a in d["acc"]}
    acc_emp = {(a["empresa"], a["mes"]) for a in d["acc"]}
    for k, m in enumerate(d["meses"]):
        _celda(tabla.cell(0, 1 + k), [CORTO[m - 1]], ROJO if m in acc_meses else TEAL)
    filas = {}
    for r in range(1, len(tabla.rows)):
        nombre = tabla.cell(r, 0).text_frame.paragraphs[0].text.strip()
        if nombre in G3:
            filas[nombre] = r
    for g, r in filas.items():
        runs0 = tabla.cell(r, 0).text_frame.paragraphs[1].runs
        runs0[0].text = f"{d['tot_g'][g]} reportes recibidos"
        for k in range(n):
            pc = d["pct"][g][k]
            fondo, color, neg = _estilo_pct(pc)
            c = tabla.cell(r, 1 + k)
            _celda(c, [f"{pc}%", str(d["mensual"][g][k])], fondo, color, neg)
            c.text_frame.paragraphs[1].runs[0].font.color.rgb = RGBColor.from_string(color)
            for lado in ("lnL", "lnR", "lnT", "lnB"):
                _borde(c, lado, 12700, "FFFFFF")
        e = d["ev"][g]
        tabla.cell(r, i_hbp).text_frame.paragraphs[0].runs[0].text = str(e["HBP"])
        otros = e["HAP"] + e["CBP"] + e["CAP"] + e["ACC"]
        c = tabla.cell(r, i_hbp + 1)
        c.text_frame.paragraphs[0].runs[0].text = str(otros)
        c.text_frame.paragraphs[1].runs[0].text = _desglose(e)
    # contorno rojo: empresa con accidente en ese mes (y el lado vecino, para que no lo tape el borde blanco)
    for g, r in filas.items():
        for k, m in enumerate(d["meses"]):
            if (g, m) not in acc_emp:
                continue
            c = tabla.cell(r, 1 + k)
            for lado in ("lnL", "lnR", "lnT", "lnB"):
                _borde(c, lado, 28575, ROJO)
            _borde(tabla.cell(r, k), "lnR", 28575, ROJO)
            _borde(tabla.cell(r, 2 + k), "lnL", 28575, ROJO)
            if r > 1:
                _borde(tabla.cell(r - 1, 1 + k), "lnB", 28575, ROJO)
            if r + 1 < len(tabla.rows):
                _borde(tabla.cell(r + 1, 1 + k), "lnT", 28575, ROJO)
    poner = lambda nombre: _forma(s, nombre).text_frame.paragraphs[0]
    _poner_runs(poner("Subtitulo"), [T["sub15"]])
    esc = poner("Escala")
    runs = list(esc.runs)
    i_emp = next(i for i, r in enumerate(runs) if r.text.startswith("empresa con accidente"))
    runs[i_emp].text = T["leyenda_empresa"]
    runs[-2].text, runs[-1].text = T["escala_corte"]
    _viñetas(_forma(s, "Lectura"), T["lectura15"])
    log.append(f"Reportes por empresa (lámina 15) ✔ ({n} meses, corte {d['dia_corte']}-{d['mes_max']:02d})")


def _ymax(d):
    pico = max(max(v) for v in d["mensual"].values())
    return max(50, ceil(pico * 1.85 / 50) * 50)


def _pts(lit, valores, es_texto):
    for pt in lit.findall(C + "pt"):
        lit.remove(pt)
    lit.find(C + "ptCount").set("val", str(len(valores)))
    for i, v in enumerate(valores):
        pt = etree.SubElement(lit, C + "pt", idx=str(i))
        etree.SubElement(pt, C + "v").text = str(v)


def _sin_ext(el):
    for x in el.findall(C + "extLst"):
        el.remove(x)
    return el


def _grafico(s, d, ymax):
    gf = next(sh for sh in s.shapes if getattr(sh, "has_chart", False) and sh.has_chart)
    cs = gf.chart._chartSpace
    n = len(d["meses"])
    cats = [MESES[m - 1][:3].capitalize() for m in d["meses"]]
    comp = d["total_mes"][:n - 1]
    ip = comp.index(max(comp)) if comp else 0
    ilc = n - 2
    acc_idx = {a["mes"] - 1 for a in d["acc"]}
    for ser in cs.findall(f".//{C}ser"):
        nombre = ser.find(f"{C}tx/{C}v").text
        vals = d["mensual"][nombre]
        _pts(ser.find(f"{C}cat/{C}strLit"), cats, True)
        _pts(ser.find(f"{C}val/{C}numLit"), vals, False)
        color = ser.find(f"{C}spPr/{A}ln/{A}solidFill/{A}srgbClr").get("val")
        dpts = ser.findall(C + "dPt")
        plantilla_dpt = copy.deepcopy(dpts[0])
        ref = dpts[-1]
        tam = plantilla_dpt.find(f".//{C}size").get("val")
        for x in dpts:
            ser.remove(x)
        dl = ser.find(C + "dLbls")
        lbls = {int(x.find(C + "idx").get("val")): x for x in dl.findall(C + "dLbl")}
        orden = sorted(lbls)
        plant = {"last": lbls[orden[-1]]}
        if nombre == "PUMA":
            plant.update(first=lbls[orden[0]], peak=lbls[orden[1]], lastc=lbls[orden[2]])
        elif nombre == "COX":
            plant.update(first=lbls[orden[0]], peak=lbls[orden[1]])
        for x in lbls.values():
            dl.remove(x)
        # marcadores
        puntos = {}
        if nombre == "JEJ":
            puntos[vals.index(max(vals))] = (tam, color)
        else:
            puntos[0] = (tam, color); puntos[ip] = (tam, color); puntos[n - 1] = (tam, color)
            if nombre == "PUMA":
                puntos[ilc] = (tam, color)
        puntos[n - 1] = (tam, color)
        if nombre == "PUMA":
            for i in acc_idx:
                puntos[i] = ("11", ROJO)
        anchor = ser.find(C + "marker")
        for i in sorted(puntos, reverse=True):
            dp = copy.deepcopy(plantilla_dpt)
            _sin_ext(dp)
            dp.find(C + "idx").set("val", str(i))
            dp.find(f".//{C}size").set("val", puntos[i][0])
            dp.find(f".//{A}srgbClr").set("val", puntos[i][1])
            anchor.addnext(dp)
        # etiquetas
        etq = {n - 1: ("last", f"{vals[n - 1]} {nombre}")}
        for rol, i in (("peak", ip), ("lastc", ilc), ("first", 0)):
            if rol in plant and i not in etq and 0 <= i < n:
                etq[i] = (rol, str(vals[i]))
        for i in sorted(etq, reverse=True):
            rol, texto = etq[i]
            x = _sin_ext(copy.deepcopy(plant[rol]))
            x.find(C + "idx").set("val", str(i))
            x.find(f".//{A}t").text = texto
            dl.insert(0, x)
    for ax in cs.findall(f".//{C}valAx"):
        mx = ax.find(f"{C}scaling/{C}max")
        if mx is not None:
            mx.set("val", str(ymax))
    return gf


def _avisos(prs, s, gf, d, ymax):
    """Líneas guía y cuadros «N acc.» sobre los meses con accidente (se reubican según el gráfico)."""
    cs = gf.chart._chartSpace
    lay = cs.find(f".//{C}plotArea/{C}layout/{C}manualLayout")
    lx, ly, lw, lh = (float(lay.find(C + k).get("val")) for k in ("x", "y", "w", "h"))
    n = len(d["meses"])
    izq = gf.left + lx * gf.width
    ancho = lw * gf.width
    arr = gf.top + ly * gf.height
    alto = lh * gf.height
    guias = [sh for sh in s.shapes if sh.name.startswith("Guía accidente")]
    etqs = [sh for sh in s.shapes if sh.name.startswith("Etiqueta accidente")]
    if not guias or not etqs:
        return
    t_guia = copy.deepcopy(guias[0]._element)
    t_etq = copy.deepcopy(etqs[0]._element)
    top_etq = etqs[0].top
    ancla = etqs[-1]._element
    for sh in guias + etqs:
        if sh._element is not ancla:
            sh._element.getparent().remove(sh._element)
    nuevos = []
    por_mes = {}
    for a in d["acc"]:
        por_mes.setdefault(a["mes"], []).append(a["empresa"])
    puma = d["mensual"]["PUMA"]
    for m in sorted(por_mes):
        i = m - 1
        emps = sorted(set(por_mes[m]), key=_orden_emp)
        k = len(emps)
        lineas = [f"{len(por_mes[m])} acc."] + [e + ("," if j < k - 2 else " y" if j == k - 2 else "") for j, e in enumerate(emps)]
        x = int(izq + (i + 0.5) * ancho / n)
        w = int(max(812800, 130000 * max(len(t) for t in lineas[1:])))
        h = 482600 + 228600 * (len(lineas) - 2)
        e = copy.deepcopy(t_etq)
        pars = e.findall(f".//{A}p")
        modelo0, modelo1 = pars[0], pars[1]
        for p_ in pars:
            p_.getparent().remove(p_)
        txb = e.find(f"{P}txBody")
        for j, t in enumerate(lineas):
            p_ = copy.deepcopy(modelo0 if j == 0 else modelo1)
            p_.findall(f".//{A}t")[0].text = t
            txb.append(p_)
        e.find(f"{P}nvSpPr/{P}cNvPr").set("name", f"Etiqueta accidente {i}")
        off, ext = e.find(f"{P}spPr/{A}xfrm/{A}off"), e.find(f"{P}spPr/{A}xfrm/{A}ext")
        off.set("x", str(x - w // 2)); off.set("y", str(top_etq)); ext.set("cx", str(w)); ext.set("cy", str(h))
        g = copy.deepcopy(t_guia)
        y_punto = arr + (1 - puma[i] / ymax) * alto
        goff, gext = g.find(f"{P}spPr/{A}xfrm/{A}off"), g.find(f"{P}spPr/{A}xfrm/{A}ext")
        goff.set("x", str(x)); goff.set("y", str(top_etq + h)); gext.set("cy", str(max(int(y_punto - OFFSET_GUIA - (top_etq + h)), 50000)))
        g.find(f"{P}nvCxnSpPr/{P}cNvPr").set("name", f"Guía accidente {i}")
        nuevos += [g, e]
    ids = [int(x.get("id")) for x in s._element.iter(f"{P}cNvPr")]
    nid = max(ids) + 1
    for el in nuevos:
        for c in el.iter(f"{P}cNvPr"):
            c.set("id", str(nid)); nid += 1
        ancla.addprevious(el)
    ancla.getparent().remove(ancla)


def _lamina16(prs, d, T, log):
    s = _lamina(prs, "REPORTES Y GRAVEDAD POR EMPRESA")
    if s is None:
        return
    tabla = next(sh for sh in s.shapes if getattr(sh, "has_table", False) and sh.has_table).table
    cols = [c.text.strip() for c in tabla.rows[0].cells][1:]          # PUMA, COX, JEJ, Otras, Total
    clave = {"PUMA": "PUMA", "COX": "COX", "JEJ": "JEJ", "OTRAS": "OTRAS"}
    gr = d["grav"]
    fil = {tabla.cell(r, 0).text.strip(): r for r in range(1, len(tabla.rows))}
    def val(t, c):
        return sum(gr[t].values()) if c.upper() == "TOTAL" else gr[t][clave[c.upper()]]
    for nombre, t in (("Accidentes", "ACC"), ("CAP", "CAP"), ("HAP", "HAP"), ("CBP", "CBP"), ("HBP", "HBP")):
        for j, c in enumerate(cols):
            v = val(t, c)
            cel = tabla.cell(fil[nombre], 1 + j)
            runs = cel.text_frame.paragraphs[0].runs
            runs[0].text = _miles(v)
            if t == "ACC":
                runs[0].font.bold = v > 0
                runs[0].font.color.rgb = RGBColor.from_string(ROJO if v > 0 else "6B7780")
    for j, c in enumerate(cols):
        tot = sum(val(t, c) for t in TIPOS)
        tabla.cell(fil["Total"], 1 + j).text_frame.paragraphs[0].runs[0].text = _miles(tot)
        alto = val("CAP", c) + val("HAP", c)
        tabla.cell(fil["% alto"], 1 + j).text_frame.paragraphs[0].runs[0].text = f"{_pc(alto, tot)}%"
    _poner_runs(_forma(s, "Subtitulo").text_frame.paragraphs[0], [T["sub16"]])
    _poner_runs(_forma(s, "Nota").text_frame.paragraphs[0], [T["nota16"]])
    _viñetas(_forma(s, "Lectura"), T["lectura16"])
    ymax = _ymax(d)
    gf = _grafico(s, d, ymax)
    _avisos(prs, s, gf, d, ymax)
    log.append(f"Reportes y gravedad (lámina 16) ✔ ({len(d['acc'])} accidentes)")


# ------------------------------------------------------------------ lámina «TEMAS MÁS REPORTADOS POR EMPRESA»
OTRO_SEG = "OTRO RIESGO DE SEGURIDAD"
CORTO_RIESGO = {"INTERACCION CON ENERGIA ELECTRICA": "Energía eléctrica", "OPERACION DE VEHICULOS LIVIANOS": "Vehículos livianos",
                "CONSTRUCCION DE ZANJAS Y EXCAVACIONES": "Zanjas y excavaciones", "OTRO RIESGO MEDIOAMBIENTAL": "Medio ambiente"}
MINUS_RIESGO = {"Trabajos en altura": "altura", "Operaciones de izaje": "izaje", "Zanjas y excavaciones": "zanjas",
                "Energía eléctrica": "energía eléctrica", "Vehículos livianos": "vehículos livianos", "Medio ambiente": "medio ambiente"}
SERIES16 = ["HAP", "CAP", "HBP", "CBP", "ACC"]


def _etq_riesgo(r):
    return CORTO_RIESGO.get(_norm(r), r.strip())


def _top_riesgos(rg, maxfilas=6):
    """Top 5 sin «Otro riesgo de seguridad»; si hay empate en el 5.º puesto entran todos (marcados ¹)."""
    lista = [(r, c) for r, c in rg.items() if _norm(r) != OTRO_SEG and _norm(r) != "(SIN RIESGO)"]
    lista.sort(key=lambda rc: (-sum(rc[1].values()), _norm(rc[0])))
    if len(lista) <= 5:
        return [(_etq_riesgo(r), c, False) for r, c in lista]
    corte = sum(lista[4][1].values())
    empate = sum(lista[5][1].values()) == corte
    out = []
    for i, (r, c) in enumerate(lista):
        n = sum(c.values())
        if i < 5 or (empate and n == corte):
            out.append((_etq_riesgo(r), c, empate and n == corte))
    return out[:maxfilas]


def _tot_riesgo(c):
    return sum(c.values())


def _textos_temas(d):
    R = d["riesgos"]
    top = {g: _top_riesgos(R[g]) for g in ("PUMA", "COX")}
    n = {g: d["tot_g"][g] for g in ("PUMA", "COX")}
    ap = lambda c: c.get("HAP", 0) + c.get("CAP", 0)
    mes = CORTO[d["mes_max"] - 1].lower()
    sub = f"Top 5 por riesgo involucrado, ene–{mes} {d['anio']} · Nº de eventos, sin “Otro riesgo de seguridad”"
    l = []
    p1, c1 = top["PUMA"][0], top["COX"][0]
    if p1[0] == c1[0]:
        l.append((f"{p1[0]} es el tema nº 1 en las dos empresas: ", f"{_tot_riesgo(p1[1])} eventos en PUMA y {_tot_riesgo(c1[1])} en COX."))
    else:
        l.append(("El tema nº 1 difiere entre empresas: ", f"PUMA {p1[0].lower()} ({_tot_riesgo(p1[1])}) y COX {c1[0].lower()} ({_tot_riesgo(c1[1])})."))
    partes, nombres = [], []
    for g in ("PUMA", "COX"):
        sel = [(e, c) for e, c, _ in top[g] if _tot_riesgo(c) >= 10 and ap(c) * 100 >= 60 * _tot_riesgo(c)][:2]
        for e, c in sel:
            nom = MINUS_RIESGO.get(e, e.lower())
            if nom not in nombres:
                nombres.append(nom)
        if sel:
            partes.append(f"{g} " + " y ".join(f"{ap(c)}/{_tot_riesgo(c)} {MINUS_RIESGO.get(e, e.lower())}" for e, c in sel))
    if partes:
        l.append((f"Alto potencial pesa en {_lista(nombres)}: ", "; ".join(partes) + "."))
    else:
        l.append(("Alto potencial no domina ningún tema: ", "ningún tema del top supera el 60% de alto potencial."))
    ma = {g: next((c for r, c in R[g].items() if _norm(r) == "OTRO RIESGO MEDIOAMBIENTAL"), None) for g in ("PUMA", "COX")}
    if any(ma.values()):
        todo_hbp = all(set(c) <= {"HBP"} for c in ma.values() if c)
        nums = " y ".join(f"{_tot_riesgo(c)} en {g}" for g, c in ma.items() if c)
        l.append(("Medio ambiente es todo HBP: " if todo_hbp else "Medio ambiente: ", f"{nums.replace(' en ', ' eventos en ', 1) if ' y ' not in nums else nums.replace(' en PUMA', ' eventos en PUMA', 1)}."))
    otro = {g: _tot_riesgo(next((c for r, c in R[g].items() if _norm(r) == OTRO_SEG), {})) for g in ("PUMA", "COX")}
    pp, pc = _pc(otro["PUMA"], n["PUMA"]), _pc(otro["COX"], n["COX"])
    l.append((f"{pp}% de PUMA y {pc}% de COX no tiene riesgo específico: ", "quedó como “Otro riesgo de seguridad”."))
    resto = R["RESTO"]
    nres = sum(_tot_riesgo(c) for c in resto.values())
    tops = sorted([(r, _tot_riesgo(c)) for r, c in resto.items() if _norm(r) not in (OTRO_SEG, "(SIN RIESGO)")], key=lambda rn: (-rn[1], _norm(rn[0])))[:3]
    nombres_r = [f"{MINUS_RIESGO.get(_etq_riesgo(r), _etq_riesgo(r).lower())} {v}" for r, v in tops]
    l.append((f"Resto de empresas ({nres} eventos): ", (", ".join(nombres_r[:-1]) + " e " + nombres_r[-1] if len(nombres_r) > 1 else "".join(nombres_r)) + "."))
    return {"sub": sub, "lectura": l, "cap": {g: f"{g}, {n[g]} eventos" for g in n}, "top": top}


def _borrar_lbl(ser, modelo=None):
    dl = ser.find(C + "dLbls")
    if dl is None and modelo is not None:     # serie sin etiquetas (p. ej. Accidente en 0): copiar el formato de otra
        dl = copy.deepcopy(modelo)
        for x in dl.findall(C + "dLbl"):
            dl.remove(x)
        ser.find(C + "cat").addprevious(dl)
    for x in dl.findall(C + "dLbl"):
        dl.remove(x)
    return dl


def _grafico_temas(shape, filas, maximo):
    cs = shape.chart._chartSpace
    cats = [e + ("\u00b9" if emp else "") for e, _, emp in filas]
    plantilla_tot = None
    modelo = next((x.find(C + "dLbls") for x in cs.findall(f".//{C}ser") if x.find(f"{C}tx/{C}v").text == "HBP"), None)
    for ser in cs.findall(f".//{C}ser"):
        nombre = ser.find(f"{C}tx/{C}v").text
        if nombre == "Total":
            dl0 = ser.find(C + "dLbls").find(C + "dLbl")
            plantilla_tot = copy.deepcopy(dl0)
        if nombre == "Accidente":
            clave = "ACC"
        else:
            clave = nombre
        if nombre == "Total":
            vals = [4] * len(filas)
        else:
            vals = [c.get(clave, 0) for _, c, _ in filas]
        _pts(ser.find(f"{C}cat/{C}strLit"), cats, True)
        _pts(ser.find(f"{C}val/{C}numLit"), vals, False)
        dl = _borrar_lbl(ser, modelo)
        if nombre == "Total":
            for i, (_, c, _) in enumerate(filas):
                x = _sin_ext(copy.deepcopy(plantilla_tot))
                x.find(C + "idx").set("val", str(i))
                x.find(f".//{A}t").text = str(sum(c.values()))
                dl.insert(i, x)
        else:
            k = 0
            for i, v in enumerate(vals):
                if v < 2:
                    x = etree.Element(C + "dLbl")
                    etree.SubElement(x, C + "idx").set("val", str(i))
                    etree.SubElement(x, C + "delete").set("val", "1")
                    dl.insert(k, x)
                    k += 1
    for ax in cs.findall(f".//{C}valAx"):
        mx = ax.find(f"{C}scaling/{C}max")
        if mx is not None:
            mx.set("val", str(maximo))
    shape.height = 2376140 + (len(filas) - 5) * 268508


def _lamina16_temas(prs, d, T, log):
    s = _lamina(prs, "TEMAS MÁS REPORTADOS POR EMPRESA")
    if s is None:
        return
    if not d.get("riesgos"):
        return log.append("⚠ Temas más reportados: el Excel no trae la columna «RIESGO INVOLUCRADO»; lámina sin actualizar")
    X = _textos_temas(d)
    _poner_runs(_forma(s, "Subtitulo").text_frame.paragraphs[0], [X["sub"]])
    _poner_runs(_forma(s, "Caption linea").text_frame.paragraphs[0], [X["cap"]["PUMA"]])
    _poner_runs(_forma(s, "Caption tabla").text_frame.paragraphs[0], [X["cap"]["COX"]])
    _viñetas(_forma(s, "Lectura"), X["lectura"])
    maximo = max(sum(c.values()) for g in X["top"] for _, c, _ in X["top"][g]) + 6
    _grafico_temas(_forma(s, "Grafico PUMA"), X["top"]["PUMA"], maximo)
    _grafico_temas(_forma(s, "Grafico COX"), X["top"]["COX"], maximo)
    log.append("Temas más reportados por empresa (lámina 16) ✔")


def aplicar(prs, d, log):
    T = textos(d)
    for f in (_lamina15, _lamina16, _lamina16_temas):
        try:
            f(prs, d, T, log)
        except Exception as e:
            log.append(f"⚠ {f.__name__.strip('_')}: {type(e).__name__}: {e}")
