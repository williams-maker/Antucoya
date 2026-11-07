"""
Motor del Resumen de Gestión de Eventos | PUMA + COX.

Recibe el Excel de eventos (bytes, hoja "Export") y devuelve un .xlsx con:
- hoja "Resumen": título, nota, 6 tarjetas y tabla por empresa (Total, Cerrados, Abiertos, % cierre,
  Avance) con FÓRMULAS VIVAS sobre la hoja Export, como la referencia, más el valor ya calculado
  en cada fórmula (para que la vista previa de Teams/navegador no salga vacía);
- hoja "Export": los eventos usados (después de los filtros), con las 6 columnas de la referencia.

El formato (fuentes, colores, bordes, anchos, altos, combinadas y formato condicional de AVANCE)
está incrustado en _ESTILOS_JSON y se extrajo de la referencia con pruebas/extraer_estilos.py.

Corre igual en Python normal y en el navegador (Pyodide 0.27.2): solo stdlib, openpyxl y lxml.
"""
import io
import json
import re
import unicodedata
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta

import openpyxl
from lxml import etree
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

VERSION = "1 (06-11 resumen PUMA + COX, fórmulas vivas con valores)"

# <ESTILOS>
_ESTILOS_JSON = r'''{"grilla":{"1":[0,1,1,1,1,1,1,1,1,1,1,1,1],"2":[2,1,1,1,1,1,1,1,1,1,1,1,1],"3":[1,1,1,1,1,1,1,1,1,1,1,1,1],"4":[3,1,4,5,4,6,4,7,4,8,4,9,4],"5":[10,1,4,11,4,12,4,13,4,14,4,15,4],"6":[16,1,4,17,4,18,4,19,4,20,4,21,4],"7":[1,1,1,1,1,1,1,1,1,1,1,1,1],"8":[22,1,1,1,23,1,24,1,25,1,26,1,26],"cuerpo_impar":[27,1,1,1,28,1,28,1,28,1,29,1,30],"cuerpo_par":[31,1,1,1,32,1,32,1,32,1,33,1,30],"total":[22,1,1,1,34,1,34,1,34,1,35,1,30]},"alto":{"1":42.0,"2":24.0,"3":9.75,"4":31.5,"5":43.5,"6":25.5,"7":15.75,"8":25.5,"cuerpo":21.0,"total":24.0},"anchos":[[1,12,13.42578125],[13,13,23.0]],"merges_cabecera":["A1:M1","A2:M2","A4:C4","D4:E4","F4:G4","H4:I4","J4:K4","L4:M4","A5:C5","D5:E5","F5:G5","H5:I5","J5:K5","L5:M5","A6:C6","D6:E6","F6:G6","H6:I6","J6:K6","L6:M6","A8:D8","E8:F8","G8:H8","I8:J8","K8:L8"],"merges_cuerpo":["A:D","E:F","G:H","I:J","K:L"],"merges_total":["A:D","E:F","G:H","I:J","K:L"],"cf":[{"type":"cellIs","operator":"greaterThanOrEqual","formula":["1"],"priority":2,"bg":"FF2EAA52"},{"type":"cellIs","operator":"greaterThanOrEqual","formula":["0.5"],"priority":3,"bg":"FFF0A030"},{"type":"cellIs","operator":"lessThan","formula":["0.5"],"priority":4,"bg":"FFD1243F"}],"vista":{"showGridLines":false},"pagina":{"orientation":"landscape","horizontalDpi":300,"verticalDpi":300},"export":{"anchos":[[1,1,22.0],[2,5,18.0],[6,6,16.0]],"freeze":"A2","encabezados":["Empresa","Clasificación","Potencialidad","Acción Correctiva","Validación Evento","Cierre Evento"],"estilo_encabezado":36,"estilo_dato":37},"estilos":[{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",24.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":[null,null,false,0.0],"border":[null,null,null,null],"fill":null,"font":["Calibri",11.0,false,false,{"theme":1,"tint":0.0}],"numfmt":"General"},{"align":["center","top",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",11.0,false,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":[null,null,false,0.0],"border":[null,["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":null,"font":["Calibri",11.0,false,false,{"theme":1,"tint":0.0}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FF228B54"},{"rgb":"FF2EAA52"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFF0A030"},{"rgb":"FFFF8080"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FF2AA3B8"},{"rgb":"FF2EAA52"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFD1243F"},{"rgb":"FF993366"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",true,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FF6B3FB5"},{"rgb":"FF333399"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFEEF3FA"},{"rgb":"FFE8F6F9"}],"font":["Arial",24.0,true,false,{"rgb":"FF17375E"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFE9F6EE"},{"rgb":"FFE8F6F9"}],"font":["Arial",24.0,true,false,{"rgb":"FF228B54"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFFEF5E6"},{"rgb":"FFFDEBEE"}],"font":["Arial",24.0,true,false,{"rgb":"FFF0A030"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFE8F6F9"},{"rgb":"FFE9F6EE"}],"font":["Arial",24.0,true,false,{"rgb":"FF2AA3B8"}],"numfmt":"0.00%"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFFDEBEE"},{"rgb":"FFF1EAFA"}],"font":["Arial",24.0,true,false,{"rgb":"FFD1243F"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFF1EAFA"},{"rgb":"FFEEF3FA"}],"font":["Arial",24.0,true,false,{"rgb":"FF6B3FB5"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFEEF3FA"},{"rgb":"FFE8F6F9"}],"font":["Arial",10.0,true,false,{"rgb":"FF17375E"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFE9F6EE"},{"rgb":"FFE8F6F9"}],"font":["Arial",10.0,true,false,{"rgb":"FF228B54"}],"numfmt":"0.00%"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFFEF5E6"},{"rgb":"FFFDEBEE"}],"font":["Arial",10.0,true,false,{"rgb":"FFF0A030"}],"numfmt":"0.00%"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFE8F6F9"},{"rgb":"FFE9F6EE"}],"font":["Arial",10.0,true,false,{"rgb":"FF2AA3B8"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFFDEBEE"},{"rgb":"FFF1EAFA"}],"font":["Arial",10.0,true,false,{"rgb":"FFD1243F"}],"numfmt":"0.00%"},{"align":["center","center",false,0.0],"border":[["thin",{"rgb":"FFFFFFFF"}],["thin",{"rgb":"FFFFFFFF"}],null,null],"fill":["solid",{"rgb":"FFF1EAFA"},{"rgb":"FFEEF3FA"}],"font":["Arial",10.0,true,false,{"rgb":"FF6B3FB5"}],"numfmt":"0.00%"},{"align":["left","center",false,1.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF2563A8"},{"rgb":"FF3366FF"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF228B54"},{"rgb":"FF2EAA52"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFF0A030"},{"rgb":"FFFF8080"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":["left","center",false,1.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFFFFFFF"},{"rgb":"FFFEF5E6"}],"font":["Arial",10.0,true,false,{"rgb":"FF17375E"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFFFFFFF"},{"rgb":"FFFEF5E6"}],"font":["Arial",10.0,false,false,null],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFFFFFFF"},{"rgb":"FFFEF5E6"}],"font":["Arial",10.0,false,false,null],"numfmt":"0.00%"},{"align":[null,null,false,0.0],"border":[null,null,null,null],"fill":null,"font":["Calibri",11.0,false,false,{"theme":1,"tint":0.0}],"numfmt":";;;"},{"align":["left","center",false,1.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFEEF3FA"},{"rgb":"FFE8F6F9"}],"font":["Arial",10.0,true,false,{"rgb":"FF17375E"}],"numfmt":"General"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFEEF3FA"},{"rgb":"FFE8F6F9"}],"font":["Arial",10.0,false,false,null],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FFEEF3FA"},{"rgb":"FFE8F6F9"}],"font":["Arial",10.0,false,false,null],"numfmt":"0.00%"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"#,##0"},{"align":["center","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",10.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"0.00%"},{"align":["left","center",false,0.0],"border":[null,null,null,null],"fill":["solid",{"rgb":"FF17375E"},{"rgb":"FF333333"}],"font":["Arial",11.0,true,false,{"rgb":"FFFFFFFF"}],"numfmt":"General"},{"align":[null,null,false,0.0],"border":[null,null,null,null],"fill":null,"font":["Arial",10.0,false,false,null],"numfmt":"General"}]}'''
# </ESTILOS>
_E = json.loads(_ESTILOS_JSON)

OPCIONES = {
    "gerencia": "Proyectos",   # filtra la columna "Gerencia" si existe (contiene el texto); None = no filtra
    "tipo": None,              # {"columna": "CLAS_PRINC", "excluir": [...]} o {"columna": ..., "incluir": [...]}
    "filtros_texto": None,     # texto de filtros en la nota; None = "Gerencia Proyectos, año {anio}, tipo no laboral"
    "grupo": {"titulo": "PUMA + COX", "empresas": ["PUMA", "COX S.A"]},             # tarjeta 4
    "empresa": {"titulo": "J.E.J. INGENIERÍA", "empresas": ["J.E.J. INGENIERÍA"]},  # tarjeta 5
    "filas_formula": 5000,     # alcance de los rangos de las fórmulas (se amplía si hay más eventos)
}
# columnas de la hoja Export de salida (mismo orden que la referencia) → nombres aceptados en la entrada
COLUMNAS = [("empresa", ("Empresa", "Empresa Contratista", "Contratista", "EECC")), ("clasif", ("Clasificación", "Clasificación Secundaria", "Clasificación Evento", "Tipo Evento")), ("pot", ("Potencialidad", "Potencialidad Evento", "Potencial")),
            ("accion", ("Acción Correctiva", "Acciones Correctivas", "Acción", "Accion Correctiva")), ("validacion", ("Validación Evento", "Validado", "Validación", "Validacion Evento")),
            ("cierre", ("Cierre Evento", "Cierre del Evento", "Cierre"))]
EXTRA = {"fecha": ("Fecha Evento", "Fecha del Evento", "Fecha de Evento", "Fecha"), "gerencia": ("Gerencia",), "fcierre": ("Fecha cierre evento", "Fecha de cierre", "Fecha cierre")}
OBLIGATORIAS = ("empresa", "cierre")
SIN_EMPRESA = "(Sin empresa)"


# ----------------------------------------------------------------- utilidades
def _norm(t):
    t = unicodedata.normalize("NFD", str(t or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t.replace("_", " ")).strip().upper()


def _txt(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def _fecha(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 1 < v < 100000:
        return date(1899, 12, 30) + timedelta(days=int(v))
    s = str(v).strip()[:10]
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def _cerrado(v):
    return _norm(v) == "SI"


def _q(s):
    """Texto literal para una fórmula (comillas dobladas y comodines de COUNTIF escapados)."""
    s = str(s).replace("~", "~~").replace("*", "~*").replace("?", "~?")
    return '"' + s.replace('"', '""') + '"'


def _pct(a, b):
    return a / b if b else 0


def _fmt_pct(x):
    return f"{x * 100:.2f}".replace(".", ",") + "%"


def _miles(n):
    return f"{n:,}".replace(",", ".")


# ------------------------------------------------------------ lectura y filtros
def _leer(excel_bytes, log):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    hoja = next((n for n in wb.sheetnames if _norm(n) == "EXPORT"), None)
    if hoja is None:
        raise ValueError(f"el Excel no tiene hoja 'Export' (hojas: {wb.sheetnames})")
    filas = wb[hoja].iter_rows(values_only=True)
    enc = [_norm(c) for c in next(filas, [])]
    idx = {}
    for clave, nombres in COLUMNAS + list(EXTRA.items()):
        for n in nombres:
            if _norm(n) in enc:
                idx[clave] = enc.index(_norm(n))
                break
    falt = [dict(COLUMNAS)[c][0] for c in OBLIGATORIAS if c not in idx]
    if falt:
        raise ValueError(f"faltan columnas obligatorias en la hoja Export: {falt}")
    for clave, nombres in COLUMNAS:
        if clave not in idx and clave not in OBLIGATORIAS:
            log.append(f"⚠ Columna '{nombres[0]}' no encontrada: queda vacía en la hoja Export")
    regs = []
    for fila in filas:
        if fila is None or all(v is None or str(v).strip() == "" for v in fila):
            continue
        g = lambda k: fila[idx[k]] if k in idx and idx[k] < len(fila) else None
        r = {k: _txt(g(k)) for k, _ in COLUMNAS}
        r["fecha"] = _fecha(g("fecha"))
        r["gerencia"] = _txt(g("gerencia"))
        r["fcierre"] = _fecha(g("fcierre")) if "fcierre" in idx else None
        r["_tipo"] = None
        regs.append(r)
    return regs, idx, (enc, wb[hoja])


def _filtrar(regs, idx, anio, op, enc, log):
    n0 = len(regs)
    if op.get("gerencia") and "gerencia" in idx:
        g = _norm(op["gerencia"])
        regs = [r for r in regs if g in _norm(r["gerencia"])]
        log.append(f"✔ Filtro Gerencia contiene '{op['gerencia']}': {len(regs)} de {n0}")
    elif op.get("gerencia"):
        log.append("⚠ Sin columna 'Gerencia': no se filtró por gerencia (se asume Export ya filtrado)")
    if "fecha" in idx:
        con_fecha = [r for r in regs if r["fecha"]]
        if len(con_fecha) < len(regs):
            log.append(f"⚠ {len(regs) - len(con_fecha)} evento(s) sin 'Fecha Evento' válida: quedan fuera del filtro por año")
        if anio is None and con_fecha:
            anio = max(r["fecha"] for r in con_fecha).year
            log.append(f"✔ Año no indicado: se usa {anio} (fecha de evento más reciente)")
        if anio is not None:
            regs = [r for r in con_fecha if r["fecha"].year == int(anio)]
            log.append(f"✔ Filtro año {anio}: {len(regs)} evento(s)")
    else:
        log.append("⚠ Sin columna 'Fecha Evento': no se filtró por año (se asume Export ya filtrado)")
    t = op.get("tipo")
    if t:
        col = _norm(t.get("columna"))
        if col in enc:
            i = enc.index(col)
            inc = {_norm(x) for x in t.get("incluir") or []}
            exc = {_norm(x) for x in t.get("excluir") or []}
            antes = len(regs)
            regs = [r for r in regs if (not inc or _norm(r["_raw"][i]) in inc) and _norm(r["_raw"][i]) not in exc]
            log.append(f"✔ Filtro tipo ({t.get('columna')}): {len(regs)} de {antes}")
        else:
            log.append(f"⚠ Sin columna '{t.get('columna')}': no se filtró por tipo")
    return regs, anio


# ------------------------------------------------------------------ cálculo
def calcular(excel_bytes, anio=None, opciones=None):
    """Cifras del resumen (las mismas que muestran las fórmulas). Devuelve (datos, log)."""
    op = dict(OPCIONES, **(opciones or {}))
    log = []
    regs, idx, (enc, hoja) = _leer(excel_bytes, log)
    if op.get("tipo"):  # el filtro por tipo necesita la fila cruda
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        filas = [f for f in filas if f and not all(v is None or str(v).strip() == "" for v in f)]
        for r, f in zip(regs, filas):
            r["_raw"] = list(f) + [None] * (len(enc) - len(f))
    regs, anio = _filtrar(regs, idx, anio, op, enc, log)
    if not regs:
        log.append(f"⚠ No hay eventos{f' del año {anio}' if anio else ''} después de los filtros: el resumen sale en cero")
    sin_emp = sum(1 for r in regs if not r["empresa"])
    for r in regs:
        r["empresa"] = r["empresa"] or SIN_EMPRESA
    if sin_emp:
        log.append(f"⚠ {sin_emp} evento(s) sin Empresa: se agrupan como '{SIN_EMPRESA}'")
    vacios = sum(1 for r in regs if not r["cierre"])
    if vacios:
        log.append(f"⚠ {vacios} evento(s) con 'Cierre Evento' vacío: cuentan como abiertos")
    otros = Counter(r["cierre"] for r in regs if r["cierre"] and not _cerrado(r["cierre"]))
    if otros:
        log.append(f"✔ 'Cierre Evento' distinto de Sí (abiertos): {dict(otros)}")
    if "fcierre" in idx:
        sin_fc = sum(1 for r in regs if _cerrado(r["cierre"]) and not r["fcierre"])
        if sin_fc:
            log.append(f"⚠ {sin_fc} evento(s) con Cierre = Sí pero sin 'Fecha cierre evento' (se cuentan cerrados)")
    # empresas: COUNTIF no distingue mayúsculas → agrupar igual
    grupos = {}
    for r in regs:
        grupos.setdefault(r["empresa"].casefold(), []).append(r)
    empresas = []
    for k, rs in grupos.items():
        nombre = Counter(r["empresa"] for r in rs).most_common(1)[0][0]
        cer = sum(1 for r in rs if _cerrado(r["cierre"]))
        empresas.append({"empresa": nombre, "total": len(rs), "cerrados": cer, "abiertos": len(rs) - cer,
                         "pct": _pct(cer, len(rs))})
    empresas.sort(key=lambda e: (-e["total"], e["empresa"].casefold()))
    total = len(regs)
    cerrados = sum(e["cerrados"] for e in empresas)

    def tarjeta(cfg):
        claves = {x.casefold() for x in cfg["empresas"]}
        sel = [e for e in empresas if e["empresa"].casefold() in claves]
        t = sum(e["total"] for e in sel)
        c = sum(e["cerrados"] for e in sel)
        faltan = [x for x in cfg["empresas"] if x.casefold() not in {e["empresa"].casefold() for e in empresas}]
        if faltan:
            log.append(f"⚠ Tarjeta '{cfg['titulo']}': sin eventos de {faltan}")
        return {"titulo": cfg["titulo"], "empresas": list(cfg["empresas"]), "total": t, "cerrados": c,
                "abiertos": t - c, "pct": _pct(c, t)}

    hay_cp = "clasif" in idx and "pot" in idx
    cap_hap = sum(1 for r in regs if _norm(r["pot"]) == "ALTO POTENCIAL"
                  and _norm(r["clasif"]) in ("CUASI ACCIDENTE", "HALLAZGO")) if hay_cp else None
    if not hay_cp:
        log.append("⚠ Falta 'Clasificación' o 'Potencialidad': la tarjeta CAP + HAP queda sin dato")
    if anio is None:
        anio = date.today().year
        log.append(f"⚠ Año no determinado desde el Excel: la nota usa {anio}")
    d = {"anio": int(anio), "total": total, "cerrados": cerrados, "abiertos": total - cerrados,
         "pct_cerrados": _pct(cerrados, total), "pct_abiertos": _pct(total - cerrados, total),
         "grupo": tarjeta(op["grupo"]), "empresa": tarjeta(op["empresa"]),
         "cap_hap": cap_hap, "pct_cap_hap": _pct(cap_hap or 0, total) if hay_cp else None,
         "empresas": empresas, "registros": regs, "hay_cp": hay_cp}
    log.append(f"✔ Total {_miles(total)} · cerrados {_miles(cerrados)} ({_fmt_pct(d['pct_cerrados'])}) · "
               f"abiertos {_miles(total - cerrados)} · {len(empresas)} empresa(s)")
    return d, log


# ------------------------------------------------------------------ escritura
_CACHE = {}


def _color(c):
    if not c:
        return None
    if "rgb" in c:
        return Color(rgb=c["rgb"])
    if "theme" in c:
        return Color(theme=c["theme"], tint=c.get("tint") or 0.0)
    if "indexed" in c:
        return Color(indexed=c["indexed"])
    return None


def _objetos(sid):
    if sid not in _CACHE:
        e = _E["estilos"][sid]
        name, sz, b, i, col = e["font"]
        fill = PatternFill(fill_type=e["fill"][0], fgColor=_color(e["fill"][1]), bgColor=_color(e["fill"][2])) \
            if e["fill"] else PatternFill()
        lados = [Side(style=s[0], color=_color(s[1])) if s else Side() for s in e["border"]]
        h, v, w, ind = e["align"]
        _CACHE[sid] = (Font(name=name, sz=sz, b=b, i=i, color=_color(col)), fill,
                       Border(left=lados[0], right=lados[1], top=lados[2], bottom=lados[3]),
                       Alignment(horizontal=h, vertical=v, wrap_text=w, indent=ind), e["numfmt"])
    return _CACHE[sid]


def _estilo(cell, sid):
    cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = _objetos(sid)


def _anchos(ws, lista):
    for mn, mx, w in lista:
        for c in range(mn, mx + 1):
            ws.column_dimensions[get_column_letter(c)].width = w


def _hoja_export(wb, regs):
    ws = wb.create_sheet("Export")
    enc = _E["export"]["encabezados"]
    claves = [k for k, _ in COLUMNAS]
    se, sd = _E["export"]["estilo_encabezado"], _E["export"]["estilo_dato"]
    for j, t in enumerate(enc, 1):
        _estilo(ws.cell(1, j, t), se)
    for i, r in enumerate(regs, 2):
        for j, k in enumerate(claves, 1):
            _estilo(ws.cell(i, j, r[k]), sd)
    _anchos(ws, _E["export"]["anchos"])
    ws.freeze_panes = _E["export"]["freeze"]
    ws.auto_filter.ref = f"A1:{get_column_letter(len(enc))}{max(1, len(regs)) + 1}"
    return ws


def _hoja_resumen(ws, d, op, log):
    n = len(d["registros"])
    fin = max(int(op.get("filas_formula") or 5000), n + 1)
    R = lambda col: f"Export!${col}$2:${col}${fin}"
    EMP, CLA, POT, CIE = R("A"), R("B"), R("C"), R("F")
    si = lambda crit: f"COUNTIFS({EMP},{crit},{CIE},\"Si\")+COUNTIFS({EMP},{crit},{CIE},\"Sí\")"
    cnt = lambda crit: f"COUNTIF({EMP},{crit})"
    vals = {}

    def put(ref, valor, formula=None):
        ws[ref] = ("=" + formula) if formula else valor
        if formula:
            vals[ref] = valor

    # filas 1-8 con su estilo
    G = _E["grilla"]
    for r in range(1, 9):
        for c in range(1, 14):
            _estilo(ws.cell(r, c), G[str(r)][c - 1])
    for m in _E["merges_cabecera"]:
        ws.merge_cells(m)
    g, e = d["grupo"], d["empresa"]
    put("A1", f"RESUMEN GESTIÓN DE EVENTOS | {g['titulo']}")
    filtros = op.get("filtros_texto") or f"Gerencia Proyectos, año {d['anio']}, tipo no laboral"
    put("A2", f"Fuente: hoja Export · Filtros: {filtros} · Cierre = columna 'Cierre Evento' = Sí")
    for ref, t in (("A4", "TOTAL DE EVENTOS REGISTRADOS"), ("D4", "EVENTOS CERRADOS"), ("F4", "EVENTOS ABIERTOS"),
                   ("H4", g["titulo"]), ("J4", e["titulo"]), ("L4", "CAP + HAP")):
        put(ref, t)
    put("A5", d["total"], f"COUNTA({EMP})")
    put("D5", d["cerrados"], f"COUNTIF({CIE},\"Si\")+COUNTIF({CIE},\"Sí\")")
    put("F5", d["abiertos"], "$A$5-D5")
    tot_g = "+".join(cnt(_q(x)) for x in g["empresas"])
    cer_g = "+".join(si(_q(x)) for x in g["empresas"])
    put("H5", g["pct"], f"IFERROR(({cer_g})/({tot_g}),0)")
    tot_e = "+".join(cnt(_q(x)) for x in e["empresas"])
    cer_e = "+".join(si(_q(x)) for x in e["empresas"])
    put("J5", e["pct"], f"IFERROR(({cer_e})/({tot_e}),0)")
    # la tarjeta 5 de la referencia era un conteo (ACCIDENTES); ahora es % de cierre → formato % de la tarjeta 4
    ws["J5"].number_format = ws["H5"].number_format
    if d["hay_cp"]:
        put("L5", d["cap_hap"], f"COUNTIFS({POT},\"Alto Potencial\",{CLA},\"Cuasi accidente\")"
                                f"+COUNTIFS({POT},\"Alto Potencial\",{CLA},\"Hallazgo\")")
        put("L6", d["pct_cap_hap"], "IFERROR(L5/$A$5,0)")
    else:
        put("L5", "—")
        put("L6", "—")
    put("A6", "Base: hoja Export")
    put("D6", d["pct_cerrados"], "IFERROR(D5/$A$5,0)")
    put("F6", d["pct_abiertos"], "IFERROR(F5/$A$5,0)")
    put("H6", f"{g['total']} reportes", f"({tot_g})&\" reportes\"")
    put("J6", f"{e['abiertos']} de {e['total']} abiertos", f"(({tot_e})-({cer_e}))&\" de \"&({tot_e})&\" abiertos\"")
    for ref, t in (("A8", "EMPRESA"), ("E8", "TOTAL EVENTOS"), ("G8", "CERRADOS"), ("I8", "ABIERTOS"),
                   ("K8", "% CIERRE"), ("M8", "AVANCE")):
        put(ref, t)
    # tabla por empresa
    r0 = 9
    for i, emp in enumerate(d["empresas"]):
        r = r0 + i
        estilos = G["cuerpo_impar"] if i % 2 == 0 else G["cuerpo_par"]
        for c in range(1, 14):
            _estilo(ws.cell(r, c), estilos[c - 1])
        crit = _q(emp["empresa"]) if re.search(r"[*?~]", emp["empresa"]) else f"A{r}"
        put(f"A{r}", emp["empresa"])
        put(f"E{r}", emp["total"], f"COUNTIF({EMP},{crit})")
        put(f"G{r}", emp["cerrados"], si(crit))
        put(f"I{r}", emp["abiertos"], f"E{r}-G{r}")
        put(f"K{r}", emp["pct"], f"IFERROR(G{r}/E{r},0)")
        put(f"M{r}", emp["pct"], f"K{r}")
        for m in _E["merges_cuerpo"]:
            a, b = m.split(":")
            ws.merge_cells(f"{a}{r}:{b}{r}")
        ws.row_dimensions[r].height = _E["alto"]["cuerpo"]
    rt = r0 + len(d["empresas"])
    for c in range(1, 14):
        _estilo(ws.cell(rt, c), G["total"][c - 1])
    put(f"A{rt}", "TOTAL")
    if d["empresas"]:
        for col in "EGI":
            put(f"{col}{rt}", sum(x[{"E": "total", "G": "cerrados", "I": "abiertos"}[col]] for x in d["empresas"]),
                f"SUM({col}{r0}:{col}{rt - 1})")
    else:
        for col in "EGI":
            put(f"{col}{rt}", 0)
    put(f"K{rt}", d["pct_cerrados"], f"IFERROR(G{rt}/E{rt},0)")
    put(f"M{rt}", d["pct_cerrados"], f"K{rt}")
    for m in _E["merges_total"]:
        a, b = m.split(":")
        ws.merge_cells(f"{a}{rt}:{b}{rt}")
    ws.row_dimensions[rt].height = _E["alto"]["total"]
    for r in range(1, 9):
        ws.row_dimensions[r].height = _E["alto"][str(r)]
    # AVANCE: valor oculto (;;;) + relleno por formato condicional, igual que la referencia
    for regla in _E["cf"]:
        ws.conditional_formatting.add(
            f"M{r0}:M{rt}",
            Rule(type=regla["type"], operator=regla["operator"], formula=regla["formula"],
                 dxf=DifferentialStyle(fill=PatternFill(bgColor=regla["bg"]))))
    _anchos(ws, _E["anchos"])
    ws.sheet_view.showGridLines = _E["vista"]["showGridLines"]
    pg = _E["pagina"]
    ws.page_setup.orientation = pg.get("orientation")
    ws.page_setup.horizontalDpi = pg.get("horizontalDpi")
    ws.page_setup.verticalDpi = pg.get("verticalDpi")
    return vals


def _inyectar_valores(xlsx, vals):
    """Agrega el valor calculado (<v>) a cada fórmula de la hoja Resumen (sheet1.xml)."""
    zin = zipfile.ZipFile(io.BytesIO(xlsx))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                root = etree.fromstring(data)
                ns = root.nsmap.get(None)
                for c in root.iter(f"{{{ns}}}c"):
                    ref = c.get("r")
                    if c.find(f"{{{ns}}}f") is None or ref not in vals:
                        continue
                    for v in c.findall(f"{{{ns}}}v"):
                        c.remove(v)
                    v = etree.SubElement(c, f"{{{ns}}}v")
                    val = vals[ref]
                    if isinstance(val, str):
                        c.set("t", "str")
                        v.text = val
                    else:
                        c.attrib.pop("t", None)
                        v.text = repr(float(val)) if isinstance(val, float) else str(int(val))
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    return out.getvalue()


def nombre_archivo(fecha=None):
    f = _fecha(fecha) or date.today()
    return f"Resumen_Gestion_de_Eventos_{f.strftime('%d-%m-%Y')}.xlsx"


def generar_resumen(excel_bytes, anio=None, opciones=None):
    """Excel de eventos (bytes) → (xlsx_bytes, log). opciones: ver OPCIONES."""
    op = dict(OPCIONES, **(opciones or {}))
    d, log = calcular(excel_bytes, anio, op)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"
    vals = _hoja_resumen(ws, d, op, log)
    _hoja_export(wb, d["registros"])
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    log.insert(0, f"✔ Nombre sugerido: {nombre_archivo(op.get('fecha'))}")
    log.append(f"✔ Hoja Resumen: {len(d['empresas'])} empresa(s), {len(vals)} fórmulas con valor calculado")
    return _inyectar_valores(buf.getvalue(), vals), log
