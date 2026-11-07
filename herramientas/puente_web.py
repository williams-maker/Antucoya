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
    except Exception as e:
        return _err(e)
    _SALIDA = pptx
    nombre = ""
    for l in log:
        if l.startswith("✔ Nombre:"):
            nombre = l.split(":", 1)[1].strip()
    return json.dumps({"ok": True, "log": log, "nombre": nombre}, ensure_ascii=False)
