"""Lee el informe «Cumplimiento de Focos» de JEJ (PPTX) y devuelve lo necesario para el bloque Ejecución:
focos verificados, resultado por foco, fotos con su leyenda. Corre en el navegador (Pyodide); no sube nada.

leer_informe(pptx_bytes) -> dict (JSON serializable)
"""
import io
import re
import base64
import datetime

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image

PULG = 914400
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
MES3 = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
RE_ID = re.compile(r"\b(ACC|CAAP|HAP|HBP|CBP)\s*(?:ID\s*)?(\d{6})\b")


def _textos(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append({"t": sh.text_frame.text.strip(), "x": (sh.left or 0) / PULG, "y": (sh.top or 0) / PULG})
    return out


def _jpeg(blob):
    im = Image.open(io.BytesIO(blob)).convert("RGB")
    if im.size[0] < 250 or im.size[1] < 170:
        return None
    if im.size[0] > 1000:
        im = im.resize((1000, max(1, int(im.size[1] * 1000 / im.size[0]))))
    b = io.BytesIO()
    im.save(b, "JPEG", quality=82)
    return {"w": im.size[0], "h": im.size[1], "src": "data:image/jpeg;base64," + base64.b64encode(b.getvalue()).decode()}


def _leyenda(pic, textos):
    px, py = (pic.left or 0) / PULG, (pic.top or 0) / PULG
    mejor = None
    for t in textos:
        if len(t["t"]) < 10 or t["y"] < py or abs(t["x"] - px) > 0.45:
            continue
        if mejor is None or t["y"] < mejor["y"]:
            mejor = t
    return mejor["t"] if mejor else ""


def _tema(instr):
    """«Foco N°1 — Controles eléctricos en tableros: verificar…» -> «Controles eléctricos en tableros»."""
    m = re.search(r"Foco\s*N°\s*\d+\s*[—–-]\s*([^:]+):", instr)
    return m.group(1).strip() if m else ""


def _lema(texto_portada):
    m = re.search(r"(\d{1,2})\s*al\s*(\d{1,2})\s*de\s*([A-Za-záéíóú]+)(?:\s*de\s*(\d{4}))?", texto_portada, re.I)
    if not m:
        return ""
    d1, d2, mes = int(m.group(1)), int(m.group(2)), MESES.get(m.group(3).lower())
    if not mes:
        return ""
    rango = f"{d1:02d}–{d2:02d} {MES3[mes - 1]}"
    try:
        a = int(m.group(4)) if m.group(4) else datetime.date.today().year
        if datetime.date(a, mes, d1).weekday() in (4, 5) and datetime.date(a, mes, d2).weekday() == 6:
            return f"GESTIÓN EN SEGURIDAD · FIN DE SEMANA {rango}"
    except Exception:
        pass
    return f"GESTIÓN EN SEGURIDAD · {rango}"


def _resumen(verifs, ids, textos_pie):
    n = len(verifs)
    partes = [f"{n} verificaciones" if n != 1 else "1 verificación"]
    pie = " ".join(textos_pie).lower()
    if ids:
        partes.append(", ".join(ids))
    if "mejora" in pie and "regularizada" not in pie.replace("mejora regularizada", ""):
        partes.append("mejoras en curso")
    if "regularizada" in pie:
        partes.append("OM regularizada en terreno")
    if not ids and "mejora" not in pie:
        partes.append("conforme")
    return " · ".join(partes)


NUM_PALABRA = {"un": 1, "uno": 1, "dos": 2, "tres": 3, "cuatro": 4, "cinco": 5, "seis": 6, "siete": 7, "ocho": 8}
RE_NOMBRE_ITO = re.compile(r"\s*(?:\+|/| y )\s*")


def _rango(texto):
    m = re.search(r"(\d{1,2})\s*al\s*(\d{1,2})\s*de\s*([A-Za-záéíóú]+)(?:\s*de\s*(\d{4}))?", texto, re.I)
    if not m:
        return ""
    mes = MESES.get(m.group(3).lower())
    if not mes:
        return ""
    anio = m.group(4) or str(datetime.date.today().year)
    return f"{int(m.group(1)):02d}–{int(m.group(2)):02d} {MES3[mes - 1].lower()} {anio}"


def _tabla_resumen(slide):
    """Filas de la tabla «Resumen de Verificaciones» (lámina 2): [{verif, ito, frente, herramienta, resultado}]."""
    tx = [t for t in _textos(slide) if t["y"] >= 2.0]
    columnas = [(0.3, "verif"), (4.0, "ito"), (6.3, "frente"), (8.4, "herramienta"), (10.6, "resultado")]
    def col(x):
        c = None
        for lim, nombre in columnas:
            if x >= lim - 0.05:
                c = nombre
        return c
    anclas = sorted((t for t in tx if col(t["x"]) == "verif" and re.match(r"(Foco\s*\d+|Rutinaria)", t["t"])), key=lambda t: t["y"])
    filas = []
    for a in anclas:
        f = {"verif": a["t"], "ito": "", "frente": "", "herramienta": "", "resultado": ""}
        for t in tx:
            c = col(t["x"])
            if c and c != "verif" and abs(t["y"] - a["y"]) <= 0.14 and not f[c]:
                f[c] = t["t"]
        filas.append(f)
    return filas


def _persona(n):
    p = re.sub(r"\(.*?\)", "", n).strip().lower().split()
    return (p[0] + " " + p[-1]) if len(p) > 1 else ""


def resumen_ejecutivo(slides, portada, turno, focos):
    """Datos de la lámina «Resumen ejecutivo Assurance» (indicadores + una fila por foco)."""
    filas_t = _tabla_resumen(slides[1]) if len(slides) > 1 else []
    if not filas_t:
        return {}
    cierre = " ".join(t["t"] for t in _textos(slides[1]))
    grupos = {}
    for f in filas_t:
        m = re.match(r"Foco\s*(\d+)", f["verif"])
        clave = int(m.group(1)) if m else 99
        grupos.setdefault(clave, []).append(f)

    filas, ids_h, personas, n_verif = [], [], {}, len(filas_t)
    for clave in sorted(grupos):
        g = grupos[clave]
        evid, frentes, itos, herr, estados = [], [], [], [], []
        for f in g:
            nombre = re.sub(r"^(Foco\s*\d+|Rutinaria)\s*·\s*", "", f["verif"])
            nombre = re.sub(r"\s*\(L[aá]m\.?\s*\d+\)\s*$", "", nombre).strip()
            res = f["resultado"].strip()
            evid.append(f"{nombre}: {res}" if res else nombre)
            estados.append(res)
            fr = f["frente"].replace(" — ", " · ").strip()
            if fr and fr not in frentes:
                frentes.append(fr)
            for n in RE_NOMBRE_ITO.split(f["ito"]):
                n = n.strip()
                if n and n not in itos:
                    itos.append(n)
                if n and not n.upper().startswith("ITO") and _persona(n):
                    personas[_persona(n)] = n
            for h in re.split(r"\s*\+\s*", f["herramienta"]):
                h = h.strip()
                if h and h not in herr:
                    herr.append(h)
            for tipo, num in RE_ID.findall(res):
                if f"{tipo} {num}" not in ids_h:
                    ids_h.append(f"{tipo} {num}")
        pie = " ".join(estados).lower()
        hall = [i for e in estados for i in [" ".join(x) for x in RE_ID.findall(e)] if i]
        if hall:
            estado, color = "Hallazgo reportado\n" + ", ".join(hall), "ambar"
        elif "mejora" in pie and "regularizada" not in pie:
            estado, color = "Conforme con mejoras\nen ejecución", "ambar"
        elif "regularizada" in pie:
            estado, color = "Conforme con\noportunidad de mejora", "ambar"
        else:
            estado, color = "Conforme", "verde"
        if clave == 99:
            titulo_foco, ambito = "Rutinaria", re.sub(r"^Rutinaria\s*·\s*", "", g[0]["verif"]).replace(" (Lám. 13)", "")
            ambito = re.sub(r"\s*\(L[aá]m\.?\s*\d+\)\s*$", "", ambito)
        else:
            info = next((x for x in focos if x["n"] == clave), None)
            tema = (info or {}).get("tema", "")
            titulo_foco = f"Foco {clave}" + (f" · {info['nivel']}" if info else "")
            ambito = tema if len(tema) <= 58 else tema[:58].rsplit(" ", 1)[0]
        filas.append({
            "foco": titulo_foco, "ambito": ambito,
            "frente": " / ".join(frentes), "ito": ", ".join(itos),
            "herramienta": "\n".join(herr),
            "evidencia": "; ".join(evid) + ".",
            "estado": estado, "estado_color": color,
        })

    n_cubiertos = len([c for c in grupos if c != 99])
    total = n_cubiertos
    m = re.search(r"Los\s+(\w+)\s+focos", cierre)
    if m and m.group(1).lower() in NUM_PALABRA:
        total = max(total, NUM_PALABRA[m.group(1).lower()])
    pct = round(100 * n_cubiertos / total) if total else 0
    haps = [i for i in ids_h if i.startswith(("HAP", "ACC", "CAAP"))]
    hbps = [i for i in ids_h if i.startswith(("HBP", "CBP"))]
    partes = []
    for etq, lista in (("HAP", [i for i in ids_h if i.startswith("HAP")]), ("HBP", [i for i in ids_h if i.startswith("HBP")]),
                       ("CBP", [i for i in ids_h if i.startswith("CBP")])):
        if lista:
            partes.append(f"{len(lista)} {etq}")
    hall_txt = " + ".join(partes) if partes else "Sin hallazgos"
    rango = _rango(portada)
    return {
        "titulo": "RESUMEN EJECUTIVO ASSURANCE",
        "subtitulo": "Minera Antucoya · GPRO" + (f" · Turno {turno}" if turno else "") + (f" | Verificación: {rango}" if rango else ""),
        "kpis": [
            {"titulo": "COBERTURA DE FOCOS", "valor": f"{n_cubiertos} de {total} cubiertos · {pct}%", "color": "verde" if pct == 100 else "ambar"},
            {"titulo": "VERIFICACIONES EN TERRENO", "valor": f"{n_verif} verificaciones realizadas", "color": "azul"},
            {"titulo": "HALLAZGOS REPORTADOS", "valor": hall_txt, "color": "ambar" if ids_h else "negro"},
            {"titulo": "DESPLIEGUE ITO", "valor": f"{len(personas)} responsables en terreno", "color": "negro"},
        ],
        "filas": filas,
        "pie": "JEJ Ingeniería · Inspección Técnica de Obras — Minera Antucoya",
    }


DOC = re.compile(r"check ?list|lista de|artp|\bart\b|registro|certificado|\btnc\b|procedimiento|plan de operaciones|alerta de seguridad|tarjeta|firmad", re.I)
TERRENO = re.compile(r"tablero|generador|jaula|maniobra|andamio|detector|torre|acceso|roseta|eslinga|grillete|extensiones|equipo|camión|segregaci", re.I)


def _sugeridas(focos, fotos, maximo=4):
    """Una foto de terreno (no documentos) por foco, repartiendo hasta completar el máximo."""
    por_foco = []
    for f in focos:
        cand = sorted((fotos[i] for i in f["fotos"]),
                      key=lambda p: (-(5 if DOC.search(p["leyenda"]) else 0) + (2 if TERRENO.search(p["leyenda"]) else 0), p["w"] * p["h"]),
                      reverse=True)
        por_foco.append([p["id"] for p in cand])
    sel = []
    ronda = 0
    while len(sel) < maximo and any(ronda < len(c) for c in por_foco):
        for c in por_foco:
            if ronda < len(c) and len(sel) < maximo:
                sel.append(c[ronda])
        ronda += 1
    return sel


def leer_informe(pptx_bytes):
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides = list(prs.slides)
    n_total = len(slides)
    portada = " ".join(t["t"] for t in _textos(slides[0])) if slides else ""
    m = re.search(r"Turno\s*(\d-\d)", portada)
    turno = m.group(1) if m else ""
    entrante = {"1-3": "2-4", "2-4": "1-3"}.get(turno, "")

    focos, fotos, paginas = {}, [], []
    for i, s in enumerate(slides, 1):
        tx = _textos(s)
        paginas.append({"num": i, "lineas": [t["t"] for t in sorted(tx, key=lambda q: (round(q["y"], 1), q["x"]))]})
        encab = next((t["t"] for t in tx if re.search(r"FOCO\s*N°\s*(\d+)\s*[—–-]\s*N[12]", t["t"], re.I)), None)
        if not encab or i in (1, n_total):
            continue
        n = int(re.search(r"N°\s*(\d+)", encab).group(1))
        nivel = "N1" if re.search(r"N1", encab) else "N2"
        instr = next((t["t"] for t in tx if t["t"].startswith(f"Foco N°{n}") and ":" in t["t"] and t["t"] != encab), "")
        pie = next((t["t"] for t in tx if t["y"] >= 7.1 and t["x"] >= 5.0), "")
        f = focos.setdefault(n, {"n": n, "nivel": nivel, "tema": _tema(instr), "verifs": 0, "ids": [], "pies": [], "fotos": []})
        if not f["tema"]:
            f["tema"] = _tema(instr)
        f["verifs"] += 1
        f["pies"].append(pie)
        for tipo, num in RE_ID.findall(" ".join(t["t"] for t in tx if t["y"] > 1.2)):
            etiqueta = f"{tipo} {num}"
            if etiqueta not in f["ids"]:
                f["ids"].append(etiqueta)
        for sh in s.shapes:
            if sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            try:
                j = _jpeg(sh.image.blob)
            except Exception:
                continue
            if not j:
                continue
            j.update({"id": len(fotos), "pag": i, "foco": n, "leyenda": _leyenda(sh, tx)})
            fotos.append(j)
            f["fotos"].append(j["id"])

    if not focos:
        return {"ok": False, "error": "No reconozco la estructura del informe (no encontré láminas «Foco N°…»).", "paginas": paginas}

    orden = [focos[k] for k in sorted(focos)]
    items = []
    for f in orden:
        tema = f["tema"] or f"Foco N°{f['n']}"
        items.append(f"{tema[:1].upper() + tema[1:]}: " + _resumen(range(f["verifs"]), f["ids"], f["pies"]))
    intro = f"Chequeo de focos entregados por turno {turno} para turno {entrante} GPRO." if turno and entrante else "Chequeo de focos entregados."
    return {
        "ok": True, "turno": turno, "entrante": entrante, "lema": _lema(portada), "intro": intro, "items": items,
        "focos": [{"n": f["n"], "nivel": f["nivel"], "tema": f["tema"], "fotos": f["fotos"]} for f in orden],
        "fotos": fotos, "paginas": paginas,
        "sugeridas": _sugeridas([focos[k] for k in sorted(focos)], fotos),
        "resumen": resumen_ejecutivo(slides, portada, turno, [{"n": f["n"], "nivel": f["nivel"], "tema": f["tema"]} for f in orden]),
    }
