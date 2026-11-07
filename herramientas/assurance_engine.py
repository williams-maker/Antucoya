"""Motor determinista del Assurance ITO (formato v1.3).

- analizar(excel_bytes, params) -> dict: eventos del período, indicadores 14 días, abiertos,
  AREA_GEO, posibles revalorizaciones y HBP sin tema (no redacta ni asigna temas).
- calcular_nivel_y_prioridad(focos) -> dict: nivel N1/N2 y orden de prioridad (punto 4).
- generar(plantilla_bytes, plan) -> (pptx_bytes, log): rellena la lámina 1 con un plan ya decidido
  (esquema en ASSURANCE_ESQUEMA.md) y elimina la lámina 2.
- nombre_archivo(turno, vigencia) -> str (punto 13).

Corre igual en Python normal y en el navegador (Pyodide 0.27.2): solo stdlib, openpyxl,
python-pptx y lxml; sin disco ni red.
"""

import copy
import io
import re
import unicodedata
from datetime import date, datetime, timedelta
import openpyxl

VERSION = "2 (07-11 prioridad de los 7 días del turno + registro e historial por mes)"
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
NIVELES = {
    "N1": ("N1 · CRITICIDAD ALTA", "C62828"),
    "N2": ("N2 · CRITICIDAD INTERMEDIA", "E67E22"),
}

GRAVEDADES = {
    "ACC": (6, "ACC"),
    "ACC-DM": (5, "ACC-DM"),
    "CAAP": (4, "CAAP"),
    "CA": (4, "CA"),
    "CBP": (3, "CBP"),
    "HAP": (2, "HAP"),
    "HBP": (1, "HBP"),
}

SENALES_REGEX = {
    "linea_de_fuego": r"\b(linea de fuego|retroceso|aculat|marcha atras|en movimiento|atropell|giro de(l)? (equipo|camion|grua|maquina|vehiculo))\b",
    "carga_suspendida": r"\b(carga suspendida|bajo (la )?carga|debajo de (la )?carga|carga izada|area de influencia de la carga|radio de (giro|la carga))\b",
    "energia": r"\b(energizad|energia (peligrosa|electrica|residual|cero)|tension|voltaje|\d+ ?kv|sin bloqueo|presion|presurizad|vapor|termica)\b",
    "altura": r"\b(altura|sin arnes|linea de vida|caida a distinto nivel|sin baranda|borde de)\b",
    "zanja": r"\b(zanja|excavacion)\b",
    "atmosfera": r"\b(atmosfera|espacio confinado|gas(es)?|oxigeno|monoxido)\b",
}

def _norm(t):
    """Normaliza texto: sin tildes, mayúsculas, guiones/guion bajo a espacio, espacios colapsados, strip."""
    if t is None:
        return ""
    if not isinstance(t, str):
        t = str(t)
    nfd = unicodedata.normalize("NFD", t)
    sin_tildes = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    res = sin_tildes.upper()
    res = res.replace("-", " ").replace("_", " ")
    res = re.sub(r"\s+", " ", res).strip()
    return res

def _fecha(v):
    """Convierte un valor a objeto date o None."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            base = date(1899, 12, 30)
            return base + timedelta(days=int(v))
        except Exception:
            return None
    if isinstance(v, str):
        v_str = v.strip()
        if not v_str:
            return None
        v_clean = v_str.split()[0] if " " in v_str else v_str
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(v_clean, fmt).date()
            except ValueError:
                continue
        for fmt in (
            "%d-%m-%Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S",
            "%Y-%m-%d %H:%M:%S",
            "%d-%m-%Y %H:%M",
            "%d/%m/%Y %H:%M",
            "%Y-%m-%d %H:%M",
        ):
            try:
                return datetime.strptime(v_str, fmt).date()
            except ValueError:
                continue
    return None

def _rango_texto(d1, d2, prefijo=""):
    """Genera texto de rango de fechas (ej: 10 al 23 SEP)."""
    m1 = MESES[d1.month - 1]
    m2 = MESES[d2.month - 1]
    d1_str = f"{d1.day:02d}"
    d2_str = f"{d2.day:02d}"
    if m1 == m2:
        rango = f"{d1_str} al {d2_str} {m1}"
    else:
        rango = f"{d1_str} {m1} al {d2_str} {m2}"
    if prefijo:
        return f"{prefijo} {rango}"
    return rango

def _validar_params(params):
    """Valida los parámetros de entrada."""
    if not isinstance(params, dict):
        return {"ok": False, "version": VERSION, "errores": ["params debe ser un diccionario"]}
    errores = []
    turno = params.get("turno")
    if not turno:
        errores.append("falta turno")
    else:
        turno_norm = _norm(str(turno)).replace(" ", "")
        if turno_norm in ("13", "1/3", "1 3"):
            params["turno"] = "1-3"
        elif turno_norm in ("24", "2/4", "2 4"):
            params["turno"] = "2-4"
        else:
            errores.append(f"turno inválido: {turno}")

    vigencia = params.get("vigencia")
    if not vigencia:
        errores.append("falta vigencia")
    else:
        if isinstance(vigencia, dict):
            d_desde = _fecha(vigencia.get("desde"))
            d_hasta = _fecha(vigencia.get("hasta"))
            if not d_desde or not d_hasta:
                errores.append("vigencia dict incompleta o inválida")
            elif d_desde > d_hasta:
                errores.append("vigencia desde mayor que hasta")
            else:
                vigencia["desde"] = d_desde.strftime("%Y-%m-%d")
                vigencia["hasta"] = d_hasta.strftime("%Y-%m-%d")
                vigencia["texto"] = _rango_texto(d_desde, d_hasta, "")
        elif isinstance(vigencia, str):
            parts = re.split(r"\s+al\s+|\s+hasta\s+", vigencia, flags=re.IGNORECASE)
            if len(parts) == 2:
                d_desde = _fecha(parts[0].strip())
                d_hasta = _fecha(parts[1].strip())
                if not d_desde or not d_hasta:
                    errores.append(f"vigencia texto inválida: {vigencia}")
                elif d_desde > d_hasta:
                    errores.append("vigencia desde mayor que hasta")
                else:
                    params["vigencia"] = {
                        "desde": d_desde.strftime("%Y-%m-%d"),
                        "hasta": d_hasta.strftime("%Y-%m-%d"),
                        "texto": _rango_texto(d_desde, d_hasta, ""),
                    }
            else:
                errores.append(f"formato vigencia texto inválido: {vigencia}")
        else:
            errores.append("vigencia tipo inválido")

    fcierre = params.get("fecha_cierre")
    if not fcierre:
        errores.append("falta fecha_cierre")
    else:
        d_cierre = _fecha(fcierre)
        if not d_cierre:
            errores.append(f"fecha_cierre inválida: {fcierre}")
        else:
            params["fecha_cierre"] = d_cierre.strftime("%Y-%m-%d")

    if errores:
        return {"ok": False, "version": VERSION, "errores": errores}
    return None

def _get_effective_type(ev, log):
    """Determina el tipo efectivo considerando tipo_rev válido."""
    base = ev.get("tipo")
    rev = ev.get("tipo_rev")
    if not rev:
        return base
    rev_norm = _norm(rev)
    valid = False
    if base == "HBP" and rev_norm in ("CBP", "HAP", "CA", "CAAP"):
        valid = True
    elif base == "CBP" and rev_norm in ("CA", "CAAP"):
        valid = True
    elif base == "HAP" and rev_norm in ("CA", "CAAP"):
        valid = True

    if valid:
        return "CAAP" if rev_norm == "CA" else rev_norm
    else:
        log.append(f"⚠ tipo_rev inválido '{rev}' para evento {ev.get('id')} con base {base}, se ignora.")
        return base

def analizar(excel_bytes, params):
    """Analiza el archivo Excel de eventos y genera las estructuras de datos del Assurance."""
    log = []
    params = copy.deepcopy(params) if isinstance(params, dict) else params  # no modificar el dict del llamador
    val_err = _validar_params(params)
    if val_err:
        return val_err

    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    except Exception as e:
        return {"ok": False, "version": VERSION, "errores": [f"Error al abrir archivo Excel: {str(e)}"]}

    sheet_found = None
    for name in wb.sheetnames:
        if _norm(name) == "EXPORT":
            sheet_found = name
            break
    if not sheet_found:
        return {"ok": False, "version": VERSION, "errores": ["Hoja 'Export' no encontrada en el libro Excel"]}

    sheet = wb[sheet_found]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return {"ok": False, "version": VERSION, "errores": ["El Excel está vacío"]}

    header_row = rows[0]
    col_map = {}
    for idx, h in enumerate(header_row):
        if h is not None:
            col_map[_norm(str(h))] = idx

    def find_col(aliases):
        for alias in aliases:
            if alias in col_map:
                return col_map[alias]
        return None

    c_id = find_col(["ID DEL EVENTO", "ID EVENTO", "ID", "N EVENTO", "N° EVENTO", "NRO EVENTO", "NUMERO EVENTO", "NUMERO DE EVENTO", "CODIGO EVENTO", "ID SAP"])
    c_fecha = find_col(["FECHA EVENTO", "FECHA DEL EVENTO", "FECHA DE EVENTO", "FECHA OCURRENCIA", "FECHA DE OCURRENCIA", "FECHA"])
    c_empresa = find_col(["EMPRESA", "EMPRESA CONTRATISTA", "EMPRESA REPORTANTE", "CONTRATISTA", "EECC"])
    c_clasif = find_col(["CLASIFICACION", "CLASIFICACION SECUNDARIA", "CLASIFICACION EVENTO", "CLASIFICACION DEL EVENTO", "TIPO EVENTO", "TIPO DE EVENTO"])
    c_pot = find_col(["POTENCIALIDAD", "POTENCIALIDAD EVENTO", "POTENCIALIDAD DEL EVENTO", "POTENCIAL"])
    c_titulo = find_col(["TITULO", "TITULO EVENTO", "TITULO DEL EVENTO", "NOMBRE EVENTO", "NOMBRE DEL EVENTO", "ASUNTO"])
    c_desc = find_col(["DESCRIPCION", "DESCRIPCION EVENTO", "DESCRIPCION DEL EVENTO", "DETALLE EVENTO", "DETALLE DEL EVENTO", "DETALLE", "RELATO"])
    c_riesgo = find_col(["RIESGO INVOLUCRADO", "RIESGO INVOLUCREDO", "RIESGO INVOLUCRADO EVENTO", "RIESGO"])
    c_cc = find_col(["CONTROL CRITICO", "CONTROL CRITICO ASOCIADO", "CONTROL CRITICO INVOLUCRADO", "CC"])
    c_area = find_col(["AREA GEO", "AREAGEO", "AREA GEOGRAFICA", "AREA GEOGRAFICA EVENTO", "ZONA GEOGRAFICA"])
    c_fcierre = find_col(["FECHA CIERRE EVENTO", "FECHA CIERRE DEL EVENTO", "FECHA DE CIERRE EVENTO", "FECHA DE CIERRE", "FECHA CIERRE"])
    c_cierre = find_col(["CIERRE EVENTO", "CIERRE DEL EVENTO", "EVENTO CERRADO", "ESTADO CIERRE", "CIERRE"])
    c_ger = find_col(["GERENCIA"])
    if c_ger is not None:
        log.append("✔ Filtro Gerencia: solo filas cuya Gerencia contiene 'Proyectos'")

    errores_cols = []
    if c_id is None:
        errores_cols.append("falta columna ID del Evento")
    if c_fecha is None:
        errores_cols.append("falta columna Fecha Evento")
    if c_clasif is None:
        errores_cols.append("falta columna Clasificación")
    if c_pot is None:
        errores_cols.append("falta columna Potencialidad")

    if errores_cols:
        return {"ok": False, "version": VERSION, "errores": errores_cols}

    if c_empresa is None:
        log.append("⚠ columna Empresa no encontrada")
    if c_titulo is None:
        log.append("⚠ columna Título no encontrada")
    if c_desc is None:
        log.append("⚠ columna Descripción no encontrada")
    if c_riesgo is None:
        log.append("⚠ columna Riesgo Involucrado no encontrada")
    if c_cc is None:
        log.append("⚠ columna Control Crítico no encontrada")
    if c_area is None:
        log.append("⚠ columna AREA_GEO no encontrada")
    if c_fcierre is None:
        log.append("⚠ columna Fecha cierre evento no encontrada")

    d_cierre = datetime.strptime(params["fecha_cierre"], "%Y-%m-%d").date()
    actual_desde = d_cierre - timedelta(days=13)
    actual_hasta = d_cierre
    prev_desde = d_cierre - timedelta(days=27)
    prev_hasta = d_cierre - timedelta(days=14)

    todos_eventos = []
    sin_fecha = []
    sin_tipo = []

    for r_idx, row in enumerate(rows[1:], start=2):
        if all(cell is None for cell in row):
            continue

        raw_id = row[c_id] if c_id is not None and c_id < len(row) else None
        if raw_id is None:
            continue
        if isinstance(raw_id, (int, float)):
            ev_id = str(int(raw_id))
        else:
            ev_id = str(raw_id).strip()

        if c_ger is not None and "PROYECTOS" not in _norm(row[c_ger] if c_ger < len(row) else ""):
            continue
        raw_fecha = row[c_fecha] if c_fecha is not None and c_fecha < len(row) else None
        f_obj = _fecha(raw_fecha)
        if not f_obj:
            sin_fecha.append(ev_id)
            log.append(f"⚠ Evento ID {ev_id} sin Fecha Evento válida o vacía.")
            continue

        empresa = str(row[c_empresa]).strip() if c_empresa is not None and c_empresa < len(row) and row[c_empresa] is not None else None
        clasif = row[c_clasif] if c_clasif is not None and c_clasif < len(row) else None
        pot = row[c_pot] if c_pot is not None and c_pot < len(row) else None
        titulo = row[c_titulo] if c_titulo is not None and c_titulo < len(row) and row[c_titulo] is not None else ""
        desc = row[c_desc] if c_desc is not None and c_desc < len(row) and row[c_desc] is not None else ""
        riesgo = row[c_riesgo] if c_riesgo is not None and c_riesgo < len(row) and row[c_riesgo] is not None else ""
        control_critico = row[c_cc] if c_cc is not None and c_cc < len(row) and row[c_cc] is not None else ""
        area_geo = str(row[c_area]).strip() if c_area is not None and c_area < len(row) and row[c_area] is not None else "SIN AREA"
        raw_fcierre = row[c_fcierre] if c_fcierre is not None and c_fcierre < len(row) else None
        fcierre_obj = _fecha(raw_fcierre)
        cierre_evento_val = row[c_cierre] if c_cierre is not None and c_cierre < len(row) else None

        c_norm = _norm(clasif)
        p_norm = _norm(pot)

        tipo = None
        dm = False
        if c_norm == "ACCIDENTE":
            tipo = "ACC"
            if re.search(r"\bDM\b", str(titulo)):
                dm = True
        elif c_norm == "CUASI ACCIDENTE":
            if p_norm == "ALTO POTENCIAL":
                tipo = "CAAP"
            elif p_norm == "BAJO POTENCIAL":
                tipo = "CBP"
        elif c_norm == "HALLAZGO":
            if p_norm == "ALTO POTENCIAL":
                tipo = "HAP"
            elif p_norm == "BAJO POTENCIAL":
                tipo = "HBP"

        if tipo is None:
            sin_tipo.append({"id": ev_id, "clasificacion": str(clasif), "potencialidad": str(pot)})
            log.append(f"⚠ Evento ID {ev_id} con clasificación o potencialidad no reconozida.")

        m_pref = re.match(r"^(ACC|CAAP|CA|CBP|HAP|HBP)(?:\s|-)", str(titulo), re.IGNORECASE)
        prefijo_titulo = m_pref.group(1).upper() if m_pref else None
        prefijo_distinto = bool(prefijo_titulo and prefijo_titulo != tipo)

        cc_flag = "(C.C.)" if "(C.C.)" in str(control_critico) or "(C.C.)" in _norm(control_critico) else ""
        cdp_flag = bool(re.search(r"\bCdP\b", str(titulo), re.IGNORECASE))

        evento = {
            "id": ev_id,
            "fecha": f_obj.strftime("%Y-%m-%d"),
            "empresa": empresa,
            "clasificacion": str(clasif) if clasif is not None else "",
            "potencialidad": str(pot) if pot is not None else "",
            "tipo": tipo,
            "dm": dm,
            "titulo": str(titulo),
            "descripcion": str(desc),
            "riesgo": str(riesgo),
            "control_critico": str(control_critico),
            "cc": cc_flag,
            "area_geo": area_geo,
            "fecha_cierre": fcierre_obj.strftime("%Y-%m-%d") if fcierre_obj else None,
            "cierre_evento": str(cierre_evento_val) if cierre_evento_val is not None else None,
            "abierto": False,
            "cdp": cdp_flag,
            "prefijo_titulo": prefijo_titulo,
            "prefijo_distinto": prefijo_distinto,
            "tema": None,
        }
        todos_eventos.append((f_obj, evento))

    # Turno que prepara el Assurance: los últimos `dias_turno` días hasta la fecha de cierre
    dias_turno = int(params.get("dias_turno") or 7)
    turno_desde = d_cierre - timedelta(days=dias_turno - 1)
    # IDs ya usados en focos de Assurance guardados (registros leídos con leer_registro)
    ya_usados = {}
    for reg in params.get("anteriores") or []:
        for f in (reg or {}).get("focos", []):
            for i in f.get("ids", []):
                ya_usados.setdefault(str(i), []).append(reg.get("nombre") or f"Assurance {reg.get('turno')} {reg.get('fecha_cierre')}")

    # Filtrar eventos del período actual
    eventos_actual = []
    eventos_anterior = []
    cubierto_check = False

    for f_obj, ev in todos_eventos:
        if f_obj <= prev_hasta and f_obj >= prev_desde:
            eventos_anterior.append(ev)
        if f_obj <= d_cierre:
            cubierto_check = True

        if actual_desde <= f_obj <= actual_hasta:
            # Determinar abierto
            is_abierto = False
            if ev["fecha_cierre"] is None:
                is_abierto = True
            else:
                fc_date = datetime.strptime(ev["fecha_cierre"], "%Y-%m-%d").date()
                if fc_date > d_cierre:
                    is_abierto = True
            if ev["cierre_evento"] is not None:
                if _norm(ev["cierre_evento"]) != "SI":
                    is_abierto = True
            ev["abierto"] = is_abierto
            # 14 días = 7 del turno anterior + 7 del turno que prepara el Assurance (los últimos 7)
            ev["en_turno"] = f_obj >= turno_desde
            ev["bloque"] = "turno" if ev["en_turno"] else "turno_anterior"
            ev["en_assurance_anterior"] = ya_usados.get(ev["id"], [])
            eventos_actual.append(ev)

    # Orden actual: fecha desc, luego id
    eventos_actual.sort(key=lambda x: (x["fecha"], x["id"]), reverse=True)

    cubierto = any(f_obj <= prev_desde for f_obj, _ in todos_eventos)

    conteo = {"ACC": 0, "CAAP": 0, "HAP": 0, "CBP": 0, "HBP": 0}
    for ev in eventos_actual:
        t = ev["tipo"]
        if t in conteo:
            conteo[t] += 1

    total_n1n2 = conteo["ACC"] + conteo["CAAP"] + conteo["HAP"] + conteo["CBP"]

    conteo_ant = {"ACC": 0, "CAAP": 0, "HAP": 0, "CBP": 0, "HBP": 0}
    for ev in eventos_anterior:
        t = ev["tipo"]
        if t in conteo_ant:
            conteo_ant[t] += 1
    total_n1n2_ant = conteo_ant["ACC"] + conteo_ant["CAAP"] + conteo_ant["HAP"] + conteo_ant["CBP"]

    if not cubierto:
        flecha = None
        comparacion_texto = "[DATO FALTANTE]"
        log.append("⚠ El Excel no cubre el período anterior de 14 días.")
    else:
        if total_n1n2 > total_n1n2_ant:
            flecha = "▲"
        elif total_n1n2 < total_n1n2_ant:
            flecha = "▼"
        else:
            flecha = "="
        comparacion_texto = f"{total_n1n2_ant} {flecha}"

    suma_tipos = sum(conteo.values())
    total_eventos_periodo = len(eventos_actual)
    verif_ok = (suma_tipos == total_eventos_periodo) and (len(sin_tipo) == 0)
    verif_detalle = f"Suma de tipos ({suma_tipos}) coincide con total del período ({total_eventos_periodo}) y sin eventos sin tipo."

    rango_actual_txt = "del " + _rango_texto(actual_desde, actual_hasta, "")
    texto_franja = f"{rango_actual_txt} Total N1+N2 {total_n1n2} · ACC {conteo['ACC']} · CAAP {conteo['CAAP']} · HAP {conteo['HAP']} · CBP {conteo['CBP']} · vs. 14 días anteriores: {comparacion_texto}"

    # Abiertos resumen
    abiertos_list = [ev for ev in eventos_actual if ev["abierto"]]
    abiertos_ids = [ev["id"] for ev in abiertos_list]
    por_empresa = {}
    por_area = {}
    for ev in abiertos_list:
        emp = ev["empresa"] or "SIN EMPRESA"
        por_empresa.setdefault(emp, []).append(ev["id"])
        ar = ev["area_geo"]
        por_area.setdefault(ar, []).append(ev["id"])

    grupos_abiertos = []
    # Agrupar por empresa y area_geo
    grupos_dict = {}
    for ev in abiertos_list:
        emp = ev["empresa"] or "SIN EMPRESA"
        ar = ev["area_geo"]
        key = (emp, ar)
        if key not in grupos_dict:
            grupos_dict[key] = {"ids": [], "titulos": []}
        grupos_dict[key]["ids"].append(ev["id"])
        grupos_dict[key]["titulos"].append(ev["titulo"])

    for (emp, ar), g_data in grupos_dict.items():
        grupos_abiertos.append({
            "empresa": emp,
            "area_geo": ar,
            "ids": g_data["ids"],
            "titulos": g_data["titulos"],
            "tema": None,
        })

    abiertos_resumen = {
        "total": len(abiertos_list),
        "ids": abiertos_ids,
        "por_empresa": por_empresa,
        "por_area": por_area,
        "grupos": grupos_abiertos,
    }

    # Areas
    areas_por_evento = {ev["id"]: ev["area_geo"] for ev in eventos_actual}
    area_counts = {}
    for ev in eventos_actual:
        ar = ev["area_geo"]
        area_counts[ar] = area_counts.get(ar, 0) + 1
    areas_unicas = [{"area_geo": ar, "n": n} for ar, n in sorted(area_counts.items(), key=lambda x: (-x[1], x[0]))]
    areas_resumen = {"por_evento": areas_por_evento, "unicas": areas_unicas}

    # Revalorización posible
    revalorizacion_posible = []
    for ev in eventos_actual:
        if ev["tipo"] in ("CBP", "HAP", "HBP"):
            desc = ev["descripcion"]
            if not desc:
                continue
            oraciones = re.split(r"\.\s+|\;\s+|\n+", desc)
            senales_encontradas = set()
            frases_encontradas = []
            for ora in oraciones:
                ora_clean = ora.strip()
                if not ora_clean:
                    continue
                ora_norm = _norm(ora_clean)
                for cat, pat in SENALES_REGEX.items():
                    if re.search(pat, ora_norm, re.IGNORECASE):
                        senales_encontradas.add(cat)
                        if ora_clean not in frases_encontradas:
                            frases_encontradas.append(ora_clean)
            if senales_encontradas:
                revalorizacion_posible.append({
                    "id": ev["id"],
                    "tipo": ev["tipo"],
                    "senales": list(senales_encontradas),
                    "frases": frases_encontradas,
                    "estado": "posible revalorización",
                })

    hbp_sin_tema = [ev for ev in eventos_actual if ev["tipo"] == "HBP" and ev["tema"] is None]

    d_vig_desde = datetime.strptime(params["vigencia"]["desde"], "%Y-%m-%d").date()
    d_vig_hasta = datetime.strptime(params["vigencia"]["hasta"], "%Y-%m-%d").date()
    m1 = MESES[d_vig_desde.month - 1]
    m2 = MESES[d_vig_hasta.month - 1]
    mes_str = m1 if m1 == m2 else f"{m1}-{m2}"
    nombre_archivo = f"ASSURANCE_TURNO_{params['turno']}_{d_vig_desde.day:02d}-{d_vig_hasta.day:02d}_{mes_str}.pptx"

    encabezado = {
        "titulo": f"ASSURANCE TURNO {params['turno']}",
        "periodo": params["vigencia"]["texto"],
        "turno": "☒ 1-3 ☐ 2-4" if params["turno"] == "1-3" else "☐ 1-3 ☒ 2-4",
    }

    log.append("✔ Análisis de Excel completado exitosamente.")

    return {
        "ok": True,
        "version": VERSION,
        "params": {
            "turno": params["turno"],
            "vigencia": params["vigencia"],
            "fecha_cierre": params["fecha_cierre"],
        },
        "encabezado": encabezado,
        "nombre_archivo": nombre_archivo,
        "indicadores": {
            "periodo": {"desde": actual_desde.strftime("%Y-%m-%d"), "hasta": actual_hasta.strftime("%Y-%m-%d"), "texto": rango_actual_txt},
            "conteo": conteo,
            "total_n1n2": total_n1n2,
            "anterior": {
                "desde": prev_desde.strftime("%Y-%m-%d"),
                "hasta": prev_hasta.strftime("%Y-%m-%d"),
                "texto": "del " + _rango_texto(prev_desde, prev_hasta, ""),
                "conteo": conteo_ant,
                "total_n1n2": total_n1n2_ant,
                "cubierto": cubierto,
            },
            "flecha": flecha,
            "comparacion_texto": comparacion_texto,
            "verificacion": {"ok": verif_ok, "suma_tipos": suma_tipos, "total": total_eventos_periodo, "detalle": verif_detalle},
            "texto_franja": texto_franja,
        },
        "turno": {
            "desde": turno_desde.strftime("%Y-%m-%d"), "hasta": d_cierre.strftime("%Y-%m-%d"), "dias": dias_turno,
            "texto": "del " + _rango_texto(turno_desde, d_cierre, ""),
            "conteo": {t: sum(1 for e in eventos_actual if e["en_turno"] and e["tipo"] == t) for t in conteo},
            "conteo_turno_anterior": {t: sum(1 for e in eventos_actual if not e["en_turno"] and e["tipo"] == t) for t in conteo},
            "ya_en_assurance_anterior": sorted({e["id"] for e in eventos_actual if e["en_assurance_anterior"]}),
        },
        "eventos": eventos_actual,
        "abiertos": abiertos_resumen,
        "areas": areas_resumen,
        "revalorizacion_posible": revalorizacion_posible,
        "hbp_sin_tema": hbp_sin_tema,
        "sin_tipo": sin_tipo,
        "sin_fecha": sin_fecha,
        "log": log,
    }

def calcular_nivel_y_prioridad(focos, prioridad_turno=True):
    """Calcula el nivel, criticidad y prioridad de una lista de focos.

    prioridad_turno (decisión de Williams, 07-11): el nivel N1/N2 y el resto del punto 4 no cambian, pero dentro
    de cada nivel los focos con eventos del turno que prepara el Assurance (eventos con "en_turno": True, los
    últimos 7 días) van antes que los focos solo con eventos del turno anterior. En N2 los focos de solo HBP
    siguen después de los que tienen un evento N1/N2. Sin la marca "en_turno" en los eventos, no cambia nada."""
    log = []
    focos_copia = copy.deepcopy(focos)
    validos_n1 = []
    validos_n2 = []
    excluidos = []

    for foco in focos_copia:
        evs = foco.get("eventos", [])
        eff_types = []
        n1n2_count = 0
        hbp_count = 0
        max_grav_score = 0
        max_grav_name = "HBP"
        has_dm = False
        max_fecha = ""

        for ev in evs:
            et = _get_effective_type(ev, log)
            eff_types.append(et)
            if et in ("ACC", "CAAP", "CA", "CBP", "HAP"):
                n1n2_count += 1
            elif et == "HBP":
                hbp_count += 1

            if ev.get("dm"):
                has_dm = True

            f_ev = ev.get("fecha", "")
            if f_ev > max_fecha:
                max_fecha = f_ev

            # Gravedad score
            if et == "ACC":
                score = 5 if has_dm else 6
                g_name = "ACC-DM" if has_dm else "ACC"
            elif et in ("CAAP", "CA"):
                score = 4
                g_name = "CAAP"
            elif et == "CBP":
                score = 3
                g_name = "CBP"
            elif et == "HAP":
                score = 2
                g_name = "HAP"
            else:
                score = 1
                g_name = "HBP"

            if score > max_grav_score:
                max_grav_score = score
                max_grav_name = g_name

        total_eventos = len(evs)
        foco["n1n2"] = n1n2_count
        foco["hbp"] = hbp_count
        foco["total_eventos"] = total_eventos
        foco["gravedad_max"] = max_grav_name
        foco["gravedad_max_score"] = max_grav_score
        foco["max_fecha"] = max_fecha
        clave = [e for e, t in zip(evs, eff_types) if t != "HBP"] or evs   # solo HBP: cuentan sus HBP
        foco["del_turno"] = any(e.get("en_turno") for e in clave)
        foco["_t"] = 0 if (prioridad_turno and foco["del_turno"]) else 1

        # Reiteración en el turno (decisión de Williams, 07-11; plantilla v1.4 punto 3): un foco nace de 2 o más
        # eventos del mismo tema (N1/N2 o HBP) con al menos uno en los 7 días del turno; los del turno anterior
        # cuentan para reiterar. Excepción: un ACC/CAAP/CA del turno genera foco aunque esté aislado.
        # Sin marca "en_turno" en los eventos, todos se consideran del turno.
        en_t = [e.get("en_turno", True) is not False for e in evs]
        alto_turno = any(t in ("ACC", "CAAP", "CA") and x for t, x in zip(eff_types, en_t))
        if not any(en_t):
            excluidos.append({**foco, "motivo": "sin eventos en los 7 días del turno"})
            continue
        if n1n2_count > 0 and not alto_turno and total_eventos < 2:
            excluidos.append({**foco, "motivo": "evento aislado: sin reiteración del tema en el turno (punto 3)"})
            continue

        # Nivel rules
        if n1n2_count == 0:
            if hbp_count >= 3:
                nivel = "N2"
                solo_hbp = True
                motivo_nivel = f"N2: solo HBP ({hbp_count})"
            else:
                excluidos.append({**foco, "motivo": "solo HBP con menos de 3 (punto 3)"})
                continue
        else:
            solo_hbp = False
            has_high_type = any(t in ("ACC", "CAAP", "CA") for t in eff_types)
            if has_high_type or n1n2_count >= 2:
                nivel = "N1"
                if has_high_type:
                    high_t = next(t for t in eff_types if t in ("ACC", "CAAP", "CA"))
                    motivo_nivel = f"N1: {high_t}"
                else:
                    motivo_nivel = f"N1: reiteración ({n1n2_count} eventos N1/N2)"
            else:
                nivel = "N2"
                motivo_nivel = f"N2: 1 evento N1/N2"

        foco["nivel"] = nivel
        foco["nivel_texto"] = NIVELES[nivel][0]
        foco["color"] = NIVELES[nivel][1]
        foco["solo_hbp"] = solo_hbp
        foco["motivo_nivel"] = motivo_nivel

        if nivel == "N1":
            validos_n1.append(foco)
        else:
            validos_n2.append(foco)

    # Ordenar N1
    # (a) n1n2 >= 2 antes que n1n2 == 1; (b) mayor gravedad máxima; (c) más eventos totales; (d) evento más reciente primero.
    validos_n1.sort(
        key=lambda f: (
            0 if f["n1n2"] >= 2 else 1,
            -f["gravedad_max_score"],
            -f["total_eventos"],
            f["max_fecha"] == "", # False (0) for non-empty, True (1) for empty
            f["max_fecha"] # ascending, so we need desc -> wait, python sort keys are ascending. To make desc for strings, we can invert or use custom logic.
        ),
        reverse=False
    )
    # Correcting sort with proper tuples for desc (negating numeric, and using a proxy or custom key)
    # Let's write a clean sorting key function for N1:
    def n1_sort_key(f):
        multi = 0 if f["n1n2"] >= 2 else 1
        grav = -f["gravedad_max_score"]
        tot = -f["total_eventos"]
        # date desc: datetime.max - date or inverted string if ISO
        date_val = f["max_fecha"] or "0000-00-00"
        # To sort desc with min-heap/ascending sort key, we can use a helper or negative epoch if parsed, or tuple of negatives if numeric. For date string YYYY-MM-DD, we can negate or use 9999-12-31 minus date.
        return (multi, grav, tot, date_val)

    # Wait, date_val desc means larger date comes first. If we want ascending sort key to put larger date first, we can convert date to integer days or negative string.
    def n1_sort_tuple(f):
        multi = 0 if f["n1n2"] >= 2 else 1
        grav = -f["gravedad_max_score"]
        tot = -f["total_eventos"]
        d = f["max_fecha"]
        # Convert date to negative ordinal for ascending sort to yield descending order
        if d:
            try:
                dt_ord = -datetime.strptime(d, "%Y-%m-%d").toordinal()
            except Exception:
                dt_ord = 0
        else:
            dt_ord = 0
        return (f["_t"], multi, grav, tot, dt_ord)

    validos_n1.sort(key=n1_sort_tuple)

    # Ordenar N2
    # (a) focos con n1n2 >= 1 antes que solo_hbp; (b) más eventos totales (incluye HBP); (c) mayor gravedad máxima; (d) más reciente.
    # Entre focos solo_hbp: más HBP; empate -> cc_o_emergencia True primero; luego más reciente.
    def n2_sort_tuple(f):
        has_n1 = 0 if f["n1n2"] >= 1 else 1
        tot = -f["total_eventos"]
        grav = -f["gravedad_max_score"]
        hbp_cnt = -f["hbp"]
        cc_em = 0 if f.get("cc_o_emergencia") else 1
        d = f["max_fecha"]
        if d:
            try:
                dt_ord = -datetime.strptime(d, "%Y-%m-%d").toordinal()
            except Exception:
                dt_ord = 0
        else:
            dt_ord = 0
        return (has_n1, f["_t"], tot, grav, hbp_cnt, cc_em, dt_ord)

    validos_n2.sort(key=n2_sort_tuple)

    todos_ordenados = validos_n1 + validos_n2
    focos_finales = []
    fuera = []

    for idx, foco in enumerate(todos_ordenados, start=1):
        if idx <= 5:
            foco["prioridad"] = idx
            focos_finales.append(foco)
        else:
            foco_fuera = {**foco, "motivo": "supera el máximo de 5 focos"}
            fuera.append(foco_fuera)

    log.append(f"✔ Cálculo de niveles y prioridades completado: {len(focos_finales)} focos seleccionados, {len(fuera)} fuera, {len(excluidos)} excluidos.")

    return {
        "focos": focos_finales,
        "fuera": fuera,
        "excluidos": excluidos,
        "log": log,
    }


# ======================================================================
# generar: rellena la lámina 1 de la plantilla con un plan ya decidido
# ======================================================================
import json
from pptx import Presentation
from pptx.dml.color import RGBColor
from lxml import etree

NEGRO = "222222"          # texto normal de la plantilla
GRIS_GUIA = "7A8594"      # texto guía (placeholder) de la plantilla
BLANCO = "FFFFFF"
VERSION_REGLAS = "1.4"    # versión de "Instrucciones IA" que implementa este motor (1.4: prioridad del turno + guardado)
TIPOS_PLAN = ("ACC", "CAAP", "CA", "CBP", "HAP", "HBP")
TEMAS_PROHIBIDOS = {"OTRO", "OTRO RIESGO DE SEGURIDAD", "SIN INFORMACION"}
CIERRE_RECURRENTE = "Refuerzo transversal en todos los recorridos y al cierre de cada tarea."
DATO_FALTANTE = "[DATO FALTANTE]"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _turno(t):
    d = re.sub(r"\D", "", str(t or ""))
    if d == "13":
        return "1-3"
    if d == "24":
        return "2-4"
    raise ValueError(f"turno inválido: {t!r} (use '1-3' o '2-4')")


def _fecha_obl(v, nombre):
    d = _fecha(v)
    if not d:
        raise ValueError(f"{nombre} inválida: {v!r}")
    return d


def _instrucciones_de(prs):
    for sh in prs.slides[0].shapes:
        if sh.name == "Instrucciones IA" and sh.has_text_frame:
            texto = sh.text_frame.text
            m = re.search(r"\bv(\d+(?:\.\d+)+)", texto.split("\n", 1)[0])
            secciones = {}
            for bloque in re.split(r"\n(?=\d{1,2}\. [A-ZÁÉÍÓÚÑ])", texto):
                mm = re.match(r"(\d{1,2})\. ", bloque)
                if mm:
                    secciones[mm.group(1)] = bloque.strip()
            return {"version": m.group(1) if m else None, "texto": texto, "secciones": secciones}
    return None


MESES_LARGOS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
                "Octubre", "Noviembre", "Diciembre"]
REGISTRO_ASUNTO = "ASSURANCE_REGISTRO v1"


def ruta_guardado(turno, fecha_cierre, vigencia):
    """Dónde guardar el Assurance en Teams (General\\Assurance), por mes de la fecha de cierre, como el Foro
    de Seguridad; el archivo lleva el nombre de la regla 13 (decisión de Williams, 07-11).
    Ej.: turno 1-3, cierre 23-09-2026, vigencia 24-09..07-10 → {"carpeta": "09. Septiembre",
    "archivo": "ASSURANCE_TURNO_1-3_24-07_SEP-OCT.pptx", "ruta": "Assurance\\09. Septiembre\\ASSURANCE_TURNO_1-3_24-07_SEP-OCT.pptx"}"""
    d = _fecha_obl(fecha_cierre, "fecha_cierre")
    carpeta = f"{d.month:02d}. {MESES_LARGOS[d.month - 1]}"
    archivo = nombre_archivo(turno, vigencia)
    return {"carpeta": carpeta, "archivo": archivo, "ruta": f"Assurance\\{carpeta}\\{archivo}"}


REGISTRO_NS = "urn:antucoya:assurance-registro"
_RT_CXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
_RT_CXML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"


def _escribir_registro(prs, registro):
    """Guarda el registro como parte customXml (estándar de Office; no se ve en la lámina)."""
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    import uuid
    for rel in list(prs.part.rels.values()):          # reemplaza un registro previo
        if rel.reltype == _RT_CXML and REGISTRO_NS.encode() in rel.target_part.blob:
            prs.part.drop_rel(rel.rId)
    usados = {str(p.partname) for p in prs.part.package.iter_parts()}
    n = next(i for i in range(90, 999) if f"/customXml/item{i}.xml" not in usados)
    root = etree.Element(f"{{{REGISTRO_NS}}}registro", nsmap={None: REGISTRO_NS})
    root.text = json.dumps(registro, ensure_ascii=False)
    item = Part(PackURI(f"/customXml/item{n}.xml"), "application/xml", prs.part.package,
                etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))
    props = Part(PackURI(f"/customXml/itemProps{n}.xml"),
                 "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", prs.part.package,
                 ('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{%s}" '
                  'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/>'
                  '</ds:datastoreItem>' % str(uuid.uuid4()).upper()).encode())
    item.relate_to(props, _RT_CXML_PROPS)
    prs.part.relate_to(item, _RT_CXML)


def leer_registro(pptx_bytes):
    """Registro que `generar` deja dentro del pptx (turno, fechas, focos e IDs) o None si no lo trae.
    La lista de registros de los Assurance guardados se pasa a analizar() en params["anteriores"]."""
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        for rel in prs.part.rels.values():
            if rel.reltype == _RT_CXML and REGISTRO_NS.encode() in rel.target_part.blob:
                return json.loads(etree.fromstring(rel.target_part.blob).text)
    except Exception:
        return None
    return None


def leer_instrucciones(plantilla_bytes):
    """Reglas del cuadro oculto "Instrucciones IA" de la lámina 1 (fuente de verdad del formato).

    Devuelve {"version": "1.3", "texto": str, "secciones": {"1": "1. ENTRADAS Y DATOS …", …, "17": …}}
    o None si la plantilla no lo trae. Sirve para que la sesión que redacta los focos use las reglas
    del propio archivo y no una copia."""
    return _instrucciones_de(Presentation(io.BytesIO(plantilla_bytes)))


def nombre_archivo(turno, vigencia):
    """Punto 13: ASSURANCE_TURNO_[1-3 o 2-4]_[DD-DD]_[MES].pptx (meses distintos → 'SEP-OCT')."""
    d1 = _fecha_obl(vigencia.get("desde"), "vigencia.desde")
    d2 = _fecha_obl(vigencia.get("hasta"), "vigencia.hasta")
    m1, m2 = MESES[d1.month - 1], MESES[d2.month - 1]
    mes = m1 if m1 == m2 else f"{m1}-{m2}"
    return f"ASSURANCE_TURNO_{_turno(turno)}_{d1.day:02d}-{d2.day:02d}_{mes}.pptx"


def _set_run(run, texto, color=NEGRO, italic=False):
    """Escribe el texto del run; color/italic None = conservar el de la plantilla."""
    run.text = texto
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None:
        run.font.italic = italic


def _escribir(p, partes):
    """partes = [(texto, color, italic)] sobre los runs del párrafo (clona el último si faltan)."""
    runs = list(p.runs)
    if not runs:
        runs = [p.add_run()]
    while len(runs) < len(partes):
        runs[-1]._r.addnext(copy.deepcopy(runs[-1]._r))
        runs = list(p.runs)
    for r, (t, c, it) in zip(runs, partes):
        _set_run(r, t, c, it)
    for r in runs[len(partes):]:
        r._r.getparent().remove(r._r)


def _parrafos(tf, n):
    """Deja exactamente n párrafos (n >= 1), clonando el primero como modelo."""
    ps = list(tf.paragraphs)
    modelo = ps[0]._p
    while len(ps) < n:
        ps[-1]._p.addnext(copy.deepcopy(modelo))
        ps = list(tf.paragraphs)
    for p in ps[n:]:
        p._p.getparent().remove(p._p)
    return list(tf.paragraphs)


def _tema_prohibido(tema):
    return _norm((tema or "").strip().rstrip(".")) in TEMAS_PROHIBIDOS


def _etiqueta_evento(ev, log, nf):
    tipo = str(ev.get("tipo") or "").upper().strip()
    if tipo not in TIPOS_PLAN:
        log.append(f"⚠ Foco {nf}: tipo de evento desconocido {tipo!r}")
    ids = ev.get("ids") or ([ev["id"]] if ev.get("id") else [])
    s = f"{tipo} ID {' / '.join(str(i) for i in ids) if ids else DATO_FALTANTE}"
    if ev.get("rev"):
        if tipo in ("CBP", "HAP"):
            s += " (rev.)"
        else:
            log.append(f"⚠ Foco {nf}: 'rev' solo aplica a CBP/HAP (un revalorizado a CA se escribe tipo 'CA')")
    if ev.get("empresa"):
        s += f" – {ev['empresa']}"
    return s + ":"


def _llenar_fila(tbl, fila, prioridad, f, log):
    nivel = f["nivel"]
    texto_nivel, color = NIVELES[nivel]
    # c0 nivel / prioridad
    c0 = tbl.cell(fila, 0)
    ps = c0.text_frame.paragraphs
    _escribir(ps[0], [(str(prioridad), BLANCO, None)])
    _escribir(ps[1], [(texto_nivel, BLANCO, None)])
    c0.fill.solid()
    c0.fill.fore_color.rgb = RGBColor.from_string(color)
    # c1 foco
    tema = (f.get("tema") or "").strip()
    if not tema:
        tema = DATO_FALTANTE
        log.append(f"⚠ Foco {prioridad}: sin tema")
    elif not tema.endswith("."):
        tema += "."
    if _tema_prohibido(tema):
        log.append(f"⚠ Foco {prioridad}: tema prohibido {tema!r} (punto 3)")
    instr = (f.get("instruccion") or "").strip() or DATO_FALTANTE
    crit = (f.get("criterio") or "").strip() or DATO_FALTANTE
    if len(instr) > 350:
        log.append(f"⚠ Foco {prioridad}: instrucción de {len(instr)} caracteres (máx. ~350)")
    if len(crit) > 150:
        log.append(f"⚠ Foco {prioridad}: criterio de {len(crit)} caracteres (máx. ~150)")
    ps = _parrafos(tbl.cell(fila, 1).text_frame, 2)
    _escribir(ps[0], [(tema + " ", NEGRO, False), (instr, NEGRO, False)])
    # la etiqueta conserva su color de la plantilla (punto 13: solo cambia el color de la celda de nivel)
    _escribir(ps[1], [("Criterio de restricción: ", None, None), (crit, NEGRO, False)])
    # c2 eventos
    lineas, n_hbp = [], 0
    for ev in f.get("eventos") or []:
        texto = (ev.get("texto") or "").strip() or DATO_FALTANTE
        if len(texto) > 200:
            log.append(f"⚠ Foco {prioridad}: texto de evento de {len(texto)} caracteres (máx. 200)")
        if str(ev.get("tipo") or "").upper() == "HBP":
            n_hbp += len(ev.get("ids") or [1])
        lineas.append((_etiqueta_evento(ev, log, prioridad), " " + texto))
    if f.get("aislado") and lineas:
        lineas[-1] = (lineas[-1][0], lineas[-1][1] + " (evento aislado)")
    if n_hbp > 3:
        log.append(f"⚠ Foco {prioridad}: {n_hbp} HBP complementarios listados (máx. 3; usar 'hbp_mas')")
    if f.get("hbp_mas"):
        lineas.append(("", f"+ {int(f['hbp_mas'])} HBP más"))
    if not lineas and f.get("sin_eventos_texto"):
        lineas = [("", str(f["sin_eventos_texto"]))]
    if not lineas:
        lineas = [("", DATO_FALTANTE)]
        log.append(f"⚠ Foco {prioridad}: sin eventos")
    for p, (et, tx) in zip(_parrafos(tbl.cell(fila, 2).text_frame, len(lineas)), lineas):
        _escribir(p, [(et, None, None), (tx, NEGRO, False)])
    # c3 herramienta, c4 evidencia
    for col, clave in ((3, "herramienta"), (4, "evidencia")):
        val = (f.get(clave) or "").strip() or DATO_FALTANTE
        _escribir(_parrafos(tbl.cell(fila, col).text_frame, 1)[0], [(val, NEGRO, False)])


def _validar_niveles(focos, log):
    """Compara nivel y orden del plan con calcular_nivel_y_prioridad (solo avisa)."""
    try:
        pseudo = []
        for i, f in enumerate(focos):
            if f.get("fijo") or f.get("manual"):
                continue                      # el foco fijo/manual no se valida con las reglas de selección
            evs = []
            for ev in f.get("eventos") or []:
                t = str(ev.get("tipo") or "").upper()
                base = {"CA": "CAAP"}.get(t, t)
                for j, _id in enumerate(ev.get("ids") or [ev.get("id") or f"?{i}"]):
                    evs.append({"id": str(_id), "tipo": base, "dm": bool(ev.get("dm")),
                                "fecha": ev.get("fecha") or "", "en_turno": ev.get("en_turno")})
            evs += [{"id": f"+{i}-{k}", "tipo": "HBP", "fecha": ""} for k in range(int(f.get("hbp_mas") or 0))]
            pseudo.append({"tema": str(i), "eventos": evs})
        r = calcular_nivel_y_prioridad(pseudo)
        calc = {x["tema"]: x for x in r["focos"]}
        for i, f in enumerate(focos):
            if f.get("fijo") or f.get("manual"):
                continue
            c = calc.get(str(i))
            if c is None:
                log.append(f"⚠ Foco {i + 1}: según el punto 3/4 no califica como foco (revisar)")
            elif c["nivel"] != f["nivel"]:
                log.append(f"⚠ Foco {i + 1}: el plan dice {f['nivel']} y el punto 4 da {c['nivel']} ({c.get('motivo_nivel', '')})")
        orden = [int(x["tema"]) for x in r["focos"]]
        if orden != sorted(orden):
            log.append("⚠ El orden del plan no coincide con la prioridad del punto 4 "
                       f"(orden calculado: {[o + 1 for o in orden]}; con fechas faltantes el desempate puede variar)")
    except Exception as e:  # la validación nunca bloquea la generación
        log.append(f"⚠ No se pudo validar nivel/prioridad: {e}")


def _textos_forma(sh, excluir_ultima_fila=False):
    if sh.has_text_frame:
        return [sh.text_frame.text]
    if getattr(sh, "has_table", False) and sh.has_table:
        filas = list(sh.table.rows)
        if excluir_ultima_fila:
            filas = filas[:-1]
        return [c.text_frame.text for r in filas for c in r.cells]
    return []


def generar(plantilla_bytes, plan):
    """Rellena la lámina 1 con `plan` (esquema en ASSURANCE_ESQUEMA.md). Devuelve (bytes, log)."""
    log = []
    if not isinstance(plan, dict):
        raise ValueError("plan debe ser un dict")
    turno = _turno(plan.get("turno"))
    vig = plan.get("vigencia") or {}
    d1 = _fecha_obl(vig.get("desde"), "vigencia.desde")
    d2 = _fecha_obl(vig.get("hasta"), "vigencia.hasta")
    if d1 > d2:
        raise ValueError("vigencia.desde es posterior a vigencia.hasta")
    focos = plan.get("focos")
    if not isinstance(focos, list):
        raise ValueError("plan.focos debe ser una lista (puede ir vacía)")
    for i, f in enumerate(focos, 1):
        if f.get("nivel") not in NIVELES:
            raise ValueError(f"foco {i}: nivel debe ser 'N1' o 'N2'")
    nombre = nombre_archivo(turno, {"desde": d1.isoformat(), "hasta": d2.isoformat()})
    log.append(f"✔ Nombre: {nombre}")

    prs = Presentation(io.BytesIO(plantilla_bytes))
    s = prs.slides[0]
    F = {sh.name: sh for sh in s.shapes}
    faltan = [n for n in ("Text 0", "Text 3", "Text 5", "Text 14", "Table 0", "Text 20", "Text 36", "Text 39") if n not in F]
    if faltan:
        raise ValueError(f"la plantilla no tiene las formas {faltan} (¿es FORMATO_ASSURANCE v1.3?)")
    instr = F.get("Instrucciones IA")
    instr_xml = etree.tostring(instr._element) if instr is not None else None
    if instr is None:
        log.append("⚠ La plantilla no tiene el cuadro 'Instrucciones IA'")
    else:
        info = _instrucciones_de(prs)
        if info["version"] != VERSION_REGLAS:
            log.append(f"⚠ La plantilla trae Instrucciones IA v{info['version']}; este motor aplica las reglas "
                       f"v{VERSION_REGLAS}: revisar si cambió algún criterio")

    # 1. Encabezado (punto 12)
    _escribir(F["Text 0"].text_frame.paragraphs[0], [(f"ASSURANCE TURNO {turno}", None, None)])
    _escribir(F["Text 3"].text_frame.paragraphs[0], [(_rango_texto(d1, d2, ""), NEGRO, False)])
    r5 = F["Text 5"].text_frame.paragraphs[0].runs[0]
    r5.text = r5.text.replace("☒", "☐").replace(f"☐ {turno}", f"☒ {turno}", 1)
    log.append(f"✔ Encabezado: ASSURANCE TURNO {turno} · {_rango_texto(d1, d2, '')}")

    # 2. Indicadores 14 días (punto 2)
    ind = plan.get("indicadores") or {}
    runs = list(F["Text 14"].text_frame.paragraphs[0].runs)
    if len(runs) != 18:
        raise ValueError(f"la franja 'Text 14' tiene {len(runs)} runs (se esperaban 18)")
    conteo = ind.get("conteo") or {}
    val = lambda k: str(conteo[k]) if k in conteo else DATO_FALTANTE
    total = ind.get("total_n1n2")
    if total is not None and all(k in conteo for k in ("ACC", "CAAP", "HAP", "CBP")):
        suma = sum(int(conteo[k]) for k in ("ACC", "CAAP", "HAP", "CBP"))
        if suma != int(total):
            log.append(f"⚠ Indicadores: ACC+CAAP+HAP+CBP = {suma} ≠ total {total}")
    per = ((ind.get("periodo") or {}).get("texto")) or DATO_FALTANTE
    esp = runs[0].text[len(runs[0].text.rstrip()):]
    for i, v in ((0, per + esp), (2, str(total) if total is not None else DATO_FALTANTE), (5, val("ACC")),
                 (8, val("CAAP")), (11, val("HAP")), (14, val("CBP")),
                 (17, ind.get("comparacion_texto") or DATO_FALTANTE)):
        _set_run(runs[i], v, NEGRO, False)
    log.append("✔ Indicadores 14 días")

    # 3. Focos (puntos 3, 4, 5, 7, 8, 9)
    if len(focos) > 5:
        log.append(f"⚠ El plan trae {len(focos)} focos; se usan los 5 primeros")
        focos = focos[:5]
    _validar_niveles(focos, log)
    for f in focos:
        if f["nivel"] not in NIVELES:
            raise ValueError("nivel inválido")
    tsh = F["Table 0"]
    tbl = tsh.table
    if len(tbl.rows) != 7:
        raise ValueError(f"la tabla de focos tiene {len(tbl.rows)} filas (se esperaban 7)")
    for i, f in enumerate(focos, 1):
        _llenar_fila(tbl, i, i, f, log)
    trs = list(tbl._tbl.tr_lst)
    for tr in trs[len(focos) + 1:6]:          # filas de foco sobrantes; la fila 6 (SI PRO) nunca
        tbl._tbl.remove(tr)
    tsh.height = sum(int(tr.h) for tr in tbl._tbl.tr_lst)
    n1 = sum(1 for f in focos if f["nivel"] == "N1")
    log.append(f"✔ Focos: {len(focos)} ({n1} N1, {len(focos) - n1} N2) + fila SI PRO intacta")
    if not focos:
        log.append("⚠ Plan sin focos: solo queda la fila SI PRO")

    # 4. Recurrente (punto 10)
    rec = plan.get("recurrente")
    if rec is None:
        partes = ["", "", "", DATO_FALTANTE]
        log.append("⚠ Recurrente sin datos → [DATO FALTANTE]")
    elif "texto" in rec:
        n = int(rec.get("n") or 0)
        partes = [f"{n} ", "evento" if n == 1 else "eventos", " HBP — ", str(rec["texto"]).strip()]
    else:
        temas = list(rec.get("temas") or [])
        if len(temas) > 3:
            log.append(f"⚠ Recurrente: {len(temas)} temas; se muestran los 3 primeros")
            temas = temas[:3]
        for t in temas:
            if _tema_prohibido(t.get("tema")):
                log.append(f"⚠ Recurrente: tema prohibido {t.get('tema')!r}")
        if temas:
            n = sum(int(t["n"]) for t in temas)
            partes = [f"{n} ", "evento" if n == 1 else "eventos", " HBP — ",
                      " · ".join(f"{t['tema']} ({int(t['n'])})" for t in temas) + ". " + CIERRE_RECURRENTE]
        else:
            partes = ["", "", "", "Sin temas HBP recurrentes en el período."]
    _escribir(F["Text 20"].text_frame.paragraphs[0], [(t, BLANCO, False) for t in partes])
    log.append("✔ Recurrente")

    # 6. Tendencias (punto 11)
    tend = plan.get("tendencias") or {}
    for clave, forma_n, fmt in (
            ("no_cc", "Text 36", lambda x: f"{x.get('tema')} ({int(x.get('n', 0))})"),
            ("cc", "Text 39", lambda x: f"{x.get('tema')}: {x.get('control')} ({int(x.get('n', 0))})")):
        lst = tend.get(clave)
        if lst is None:
            lineas = [DATO_FALTANTE]
            log.append(f"⚠ Tendencias {clave}: sin datos → [DATO FALTANTE]")
        elif not lst:
            lineas = ["Sin tendencias en el período."]
        else:
            orden = sorted(lst, key=lambda x: -int(x.get("n", 0)))
            if len(orden) > 5:
                log.append(f"⚠ Tendencias {clave}: {len(orden)} viñetas; se muestran 5")
            for x in orden:
                if _tema_prohibido(x.get("tema")):
                    log.append(f"⚠ Tendencias {clave}: tema prohibido {x.get('tema')!r}")
            lineas = [fmt(x) for x in orden[:5]]
        for p, t in zip(_parrafos(F[forma_n].text_frame, len(lineas)), lineas):
            _escribir(p, [(t, NEGRO, False)])
    log.append("✔ Tendencias")

    # 7. Eliminar la lámina 2 (instructivo)
    if len(prs.slides) >= 2:
        lst = prs.slides._sldIdLst
        sid = lst[1]
        prs.part.drop_rel(sid.rId)
        lst.remove(sid)
        log.append("✔ Lámina 2 (instructivo) eliminada")
        if len(prs.slides) > 1:
            log.append(f"⚠ La plantilla tenía más de 2 láminas; quedan {len(prs.slides)}")
    else:
        log.append("⚠ La plantilla no tiene lámina 2")

    # 8. Controles finales (punto 13)
    if instr is not None:
        igual = etree.tostring(instr._element) == instr_xml
        log.append("✔ 'Instrucciones IA' intacta" if igual else "⚠ 'Instrucciones IA' cambió")
    for sh in s.shapes:
        if sh.name == "Instrucciones IA":
            continue
        for t in _textos_forma(sh, excluir_ultima_fila=(sh.name == "Table 0")):
            resto = t.replace(DATO_FALTANTE, "")
            if "[" in resto or "]" in resto:
                log.append(f"⚠ Quedan corchetes en '{sh.name}': {t.strip()[:60]!r}")
    chicos = set()
    for sh in s.shapes:
        if sh.name == "Instrucciones IA":
            continue
        for el in sh._element.iter(f"{{{A_NS}}}rPr", f"{{{A_NS}}}endParaRPr"):
            sz = el.get("sz")
            if sz and int(sz) < 600:
                chicos.add(sh.name)
    log.append(f"⚠ Letra menor a 6 pt en {sorted(chicos)}" if chicos else "✔ Letra mínima 6 pt")
    if any(d == DATO_FALTANTE or DATO_FALTANTE in d for sh in s.shapes if sh.name != "Instrucciones IA"
           for d in _textos_forma(sh)):
        log.append("⚠ Hay casillas con [DATO FALTANTE]: informar al usuario")
    prs.core_properties.title = nombre[:-5]
    # 9. Registro para el historial (carpeta Teams General\Assurance\NN. Mes\): se lee con leer_registro()
    fc = _fecha(plan.get("fecha_cierre"))
    if not fc:
        fc = d1 - timedelta(days=1)
        log.append(f"⚠ El plan no trae fecha_cierre: se usa {fc.strftime('%d-%m-%Y')} (día antes de la vigencia)")
    rg = ruta_guardado(turno, fc, {"desde": d1.isoformat(), "hasta": d2.isoformat()})
    registro = {"v": 1, "nombre": rg["archivo"][:-5], "turno": turno, "fecha_cierre": fc.isoformat(),
                "vigencia": {"desde": d1.isoformat(), "hasta": d2.isoformat()},
                "focos": [{"prioridad": i, "nivel": f["nivel"], "tema": f.get("tema"),
                           "ids": [str(x) for ev in f.get("eventos") or [] for x in (ev.get("ids") or [ev.get("id")]) if x]}
                          for i, f in enumerate(focos, 1)],
                "generado": datetime.now().isoformat(timespec="seconds"), "motor": VERSION}
    _escribir_registro(prs, registro)
    log.insert(1, f"✔ Guardar en Teams: {rg['ruta']}")

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), log
