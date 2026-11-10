"""Puente entre la página (JavaScript) y los motores. Todo entra y sale como texto JSON o bytes.
No contiene datos de Antucoya."""
import json
import resumen_engine as R
import assurance_engine as A
import catalogo_temas as T

VERSION_PUENTE = "1 (07-11 página Resumen + Assurance)"
_SALIDA = b""
_AN = None
_P = None


def versiones():
    return json.dumps({"puente": VERSION_PUENTE, "resumen": R.VERSION, "assurance": A.VERSION,
                       "reglas": getattr(A, "VERSION_REGLAS", "?")}, ensure_ascii=False)


def obtener_salida():
    """Bytes del último archivo generado (xlsx o pptx)."""
    return _SALIDA


def _humano(msg):
    m = str(msg)
    bajo = m.lower()
    if "zip" in bajo or "not a zip" in bajo or "bad magic" in bajo:
        return "El archivo no es un Excel .xlsx válido (¿es un .xls antiguo, un CSV o está dañado?). Guárdalo como .xlsx e inténtalo de nuevo."
    if "no sheet" in bajo or "worksheet" in bajo and "empty" in bajo:
        return "El Excel no tiene hojas con datos."
    return m


def _err(msg, log=None):
    msg = _humano(msg)
    return json.dumps({"ok": False, "errores": [str(msg)], "log": list(log or [])}, ensure_ascii=False)


# ------------------------------------------------------------------ Resumen
def resumen(excel_bytes):
    global _SALIDA
    try:
        out, log = R.generar_resumen(bytes(excel_bytes))
    except Exception as e:  # mensajes del motor ya son claros (hoja Export, columnas obligatorias)
        return _err(e)
    _SALIDA = out
    nombre = ""
    for l in log:
        if "Nombre sugerido:" in l:
            nombre = l.split("Nombre sugerido:", 1)[1].strip()
    return json.dumps({"ok": True, "log": log, "nombre": nombre or "Resumen_Gestion_de_Eventos.xlsx"},
                      ensure_ascii=False)


# ---------------------------------------------------------------- Assurance
def _armar(decisiones):
    r = T.armar_plan(_AN, _P, decisiones or None, A.calcular_nivel_y_prioridad)
    return json.dumps({"ok": True, "plan": r["plan"], "qa": r["qa"], "decisiones": r["decisiones"],
                       "log": list(_AN.get("log", [])) + list(r.get("log", [])),
                       "nombre": A.nombre_archivo(_P["turno"], _P["vigencia"]),
                       "ruta": _ruta()}, ensure_ascii=False, default=str)


def _ruta():
    try:
        return A.ruta_guardado(_P["turno"], _P["fecha_cierre"], _P["vigencia"])
    except Exception:
        return None


def analizar_assurance(excel_bytes, params_json):
    """params: turno, vigencia{desde,hasta}, fecha_cierre, inspector?, proyecto?, area?"""
    global _AN, _P
    try:
        params = json.loads(params_json)
        an = A.analizar(bytes(excel_bytes), dict(params))
    except Exception as e:
        return _err(e)
    if not an.get("ok"):
        return _err("; ".join(an.get("errores") or ["No se pudo leer el Excel"]), an.get("log"))
    _AN, _P = an, params
    try:
        return _armar({})
    except Exception as e:
        return _err(e, an.get("log"))


def rearmar(decisiones_json):
    if _AN is None:
        return _err("Primero sube el Excel y pulsa Analizar")
    try:
        return _armar(json.loads(decisiones_json))
    except Exception as e:
        return _err(e)


def generar_pptx(plan_json, plantilla_bytes):
    global _SALIDA
    try:
        pptx, log = A.generar(bytes(plantilla_bytes), json.loads(plan_json))
        if _AN is not None:                      # láminas "Gestión SSO GPRO" y "Reportabilidad y focos de turno"
            import laminas_gestion as LG
            datos = LG.calcular(_AN, _P or {}, json.loads(plan_json))
            pptx, log_lg = LG.agregar(pptx, datos)
            log = list(log) + list(datos.get("log") or []) + list(log_lg)
        prev = json.loads(plan_json).get("preventiva")
        if prev:                                    # láminas "Gestión preventiva ejecutada en terreno" (desde el informe JEJ)
            import lamina_preventiva as LV
            pptx, log_v = LV.agregar_preventiva(pptx, prev)
            log = list(log) + list(log_v)
        import lamina_pulso as LP                  # lámina manual "Seguridad | ¿Qué hicimos…?" (siempre al final)
        pptx, log_p = LP.agregar_pulso(pptx, json.loads(plan_json).get("pulso") or {})
        log = list(log) + list(log_p)
    except Exception as e:
        return _err(e)
    _SALIDA = pptx
    nombre = ""
    for l in log:
        if l.startswith("✔ Nombre:"):
            nombre = l.split(":", 1)[1].strip()
    return json.dumps({"ok": True, "log": log, "nombre": nombre}, ensure_ascii=False)


def cargar_conocimiento(json_texto):
    """Carga (o limpia con texto vacío) la base de conocimiento privada. No se guarda en el repo."""
    try:
        datos = json.loads(json_texto) if json_texto and json_texto.strip() else None
        n, avisos = T.cargar_conocimiento(datos)
        return json.dumps({"ok": True, "n": n, "avisos": avisos,
                           "meta": T.BASE_CONOCIMIENTO_META}, ensure_ascii=False)
    except Exception as e:
        return _err(e)


def leer_registro(pptx_bytes):
    """Registro (turno, fechas, focos e IDs) que un Assurance generado por esta herramienta lleva dentro, o 'null'."""
    try:
        return json.dumps(A.leer_registro(bytes(pptx_bytes)), ensure_ascii=False)
    except Exception:
        return "null"
