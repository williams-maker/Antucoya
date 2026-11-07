"""Catálogo de temas y armado del plan de Assurance (determinista, solo stdlib).

Funciones públicas: normalizar, asignar_tema, asignar_todos, sugerir_revalorizacion,
resumir_evento y armar_plan. Los datos viven en catalogo_data.py (CATALOGO, REVAL).
Compatible con Pyodide 0.27.2: sin imports pesados ni acceso a disco o red.
"""

import copy
import re
import unicodedata
from functools import lru_cache

from catalogo_data import CATALOGO, REVAL

__all__ = ["UMBRAL", "RATIO_AMBIGUO", "normalizar", "asignar_tema", "asignar_todos",
           "sugerir_revalorizacion", "resumir_evento", "armar_plan"]

# ----------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------
UMBRAL = 2              # puntaje mínimo para aceptar un tema
RATIO_AMBIGUO = 0.75    # el segundo tema con >= 75% del primero vuelve ambiguo el caso
PESO_TITULO = 2
PESO_DESC = 1
MAX_INSTRUCCION = 350
MAX_HBP_LISTADOS = 3
MAX_FOCOS = 5

TIPOS_VALIDOS = ("ACC", "CAAP", "HAP", "CBP", "HBP")
GRAVEDAD = {"ACC": 6, "CAAP": 4, "CBP": 3, "HAP": 2, "HBP": 1}
TRANSICIONES = {"HBP": ("CBP", "HAP", "CA"), "CBP": ("CA",), "HAP": ("CA",)}

_NUNCA = re.compile(r"(?!x)x")

# ----------------------------------------------------------------------
# Utilidades de texto y regex
# ----------------------------------------------------------------------

def normalizar(texto):
    """Minúscula, sin tildes y con espacios colapsados."""
    if texto is None:
        return ""
    if not isinstance(texto, str):
        texto = str(texto)
    nfd = unicodedata.normalize("NFD", texto)
    sin = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin.lower()).strip()

@lru_cache(maxsize=4096)
def _rx(patron):
    """Compila un patrón (cacheado); un patrón inválido nunca calza."""
    try:
        return re.compile(patron, re.IGNORECASE)
    except (re.error, TypeError):
        return _NUNCA

def _rx_union(valor):
    """str, lista de str o patrón compilado -> patrón compilado (o None si vacío)."""
    if valor is None or valor == "" or valor == []:
        return None
    if hasattr(valor, "search"):
        return valor
    if isinstance(valor, (list, tuple, set)):
        partes = [f"(?:{p})" for p in valor if isinstance(p, str) and p]
        return _rx("|".join(partes)) if partes else None
    if isinstance(valor, str):
        return _rx(valor)
    return None

def _unir(items):
    """Une en español: 'a', 'a y b', 'a, b y c' (usa 'e' antes de palabras con 'i')."""
    items = [str(i) for i in items if str(i)]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    ultimo = items[-1]
    y = "e" if re.match(r"(?i)(i|hi(?![eau]))", ultimo) else "y"
    return ", ".join(items[:-1]) + f" {y} " + ultimo

def _oraciones(texto):
    """Divide un texto en oraciones (punto, salto de línea o punto y coma)."""
    if not texto:
        return []
    partes = re.split(r"(?<=[.!?])\s+|\n+|;\s+", str(texto).strip())
    return [p.strip() for p in partes if p and p.strip()]

# --- nombres de personas ------------------------------------------------
_CARGOS = (r"supervisor(?:a)?|operador(?:a)?|rigger|ayudante|trabajador(?:a)?|maestro|soldador|conductor|chofer|"
           r"capataz|prevencionista|ingeniero|electricista|mec[aá]nico|se[ñn]alero|gruero|eslingador|"
           r"jefe de (?:turno|[aá]rea|terreno|obra)|t[eé]cnico|andamiero|jornal|colaborador|inspector|"
           r"administrador de contrato")
_NOM = r"[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}"
_NOMBRE_COMPLETO = _NOM + r"(?:\s+(?:(?:de|del)\s+(?:la\s+|los\s+|las\s+)?)?" + _NOM + r"){0,3}"
_RX_CAPS = re.compile(r"\b" + _NOMBRE_COMPLETO)
_RX_TITULO = re.compile(r"\b(?:Sres?|Sra|Srta|Don|Doña|Dn)\.?\s+" + _NOMBRE_COMPLETO)
_RX_INICIAL = re.compile(r"\b[A-ZÁÉÍÓÚÑ]\.\s*" + _NOM)
_RX_CARGO_FIN = re.compile(r"(?i:\b(?:" + _CARGOS + r"))\s+(?:(?:don|sr\.?|sra\.?)\s+)?$")

_NOMBRES = set(normalizar(n) for n in (
    "juan jose luis carlos jorge pedro miguel manuel francisco victor cristian christian marcelo claudio sergio "
    "ricardo patricio rodrigo mauricio hector hugo oscar raul ramon roberto eduardo fernando gabriel daniel david "
    "diego pablo andres felipe nicolas matias sebastian javier ivan alejandro alex alberto alfredo angel antonio "
    "arturo benjamin boris camilo cesar cristobal dario edgardo elias emilio enrique ernesto esteban eugenio "
    "fabian franco gonzalo gustavo guillermo ignacio isaac jaime joaquin leonardo lorenzo marco marcos mario "
    "mauro maximiliano nelson nestor omar orlando osvaldo rene rolando rodolfo samuel tomas ulises waldo walter "
    "william williams yerko maria ana carmen carolina claudia constanza daniela camila fernanda francisca "
    "gabriela isabel jessica karen katherine laura lorena marcela margarita monica natalia paola patricia "
    "paulina sandra silvia sofia valentina veronica ximena yessica johana javiera macarena pamela andrea "
    "alejandra").split())

_NO_NOMBRE = set(normalizar(n) for n in (
    "empresa area zona sector turno planta proyecto faena mina camion grua equipo pluma excavadora cargador "
    "contrato contratista electrico mecanico civil seguridad produccion operaciones mantencion mantenimiento "
    "obra terreno montaje izaje altura andamios gerencia").split())

def _reemplazo_nombre(pre):
    """Qué escribir en lugar de un nombre, según lo que lo precede."""
    pre_s = pre.rstrip()
    if _RX_CARGO_FIN.search(pre):
        return ""                                   # el cargo ya está escrito: se elimina el nombre
    ultimo = re.findall(r"\S+", pre_s)[-1] if pre_s else ""
    if not pre_s or ultimo.endswith((",", ".")) or normalizar(ultimo) in ("que", "y", "e"):
        return "el trabajador"
    if normalizar(ultimo) in ("el", "al", "del", "un", "la", "este", "ese"):
        return "trabajador"
    return "personal"

def _sub_caps(m):
    s = m.group(0)
    toks = s.split()
    prim = normalizar(toks[0])
    pre = m.string[:m.start()]
    cargo_previo = bool(_RX_CARGO_FIN.search(pre))
    es_nombre = prim in _NOMBRES or (
        cargo_previo and len(toks) >= 2 and not any(normalizar(t) in _NO_NOMBRE for t in toks))
    return _reemplazo_nombre(pre) if es_nombre else s

def _sin_nombres(texto):
    """Reemplaza por el cargo (o elimina) los patrones de nombres de personas."""
    if not texto:
        return ""
    s = str(texto)
    s = _RX_TITULO.sub(lambda m: _reemplazo_nombre(m.string[:m.start()]), s)
    s = _RX_INICIAL.sub(lambda m: _reemplazo_nombre(m.string[:m.start()]), s)
    s = _RX_CAPS.sub(_sub_caps, s)
    s = re.sub(r"\(\s*(?:personal|el trabajador|trabajador)?\s*\)", "", s)
    s = re.sub(r"\s+([,.;:])", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()

_COLA = {"de", "del", "la", "el", "los", "las", "con", "y", "e", "en", "a", "al", "por", "para", "sin", "que",
         "o", "un", "una", "su", "sus", "se", "lo"}

def _limpiar_cola(s, quitar_conectores):
    s = s.rstrip(" ,;:.-–")
    while quitar_conectores:
        m = re.search(r"\s+(\w+)$", s)
        if m and normalizar(m.group(1)) in _COLA:
            s = s[:m.start()].rstrip(" ,;:.-–")
        else:
            break
    return s

def _cortar(s, n):
    """Corta en palabra completa a lo más n caracteres."""
    s = s.strip()
    cortado = False
    if len(s) > n:
        corte = s[:n]
        if s[n] != " " and " " in corte:
            corte = corte.rsplit(" ", 1)[0]
        s = corte
        cortado = True
    return _limpiar_cola(s, cortado)

def _minuscula_inicial(s):
    """Primera letra en minúscula, salvo siglas o palabras con dígitos."""
    if not s:
        return s
    w = re.match(r"\S+", s).group(0)
    letras = re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", w)
    if len(letras) > 1 and w.upper() == w:
        return s
    if any(c.isdigit() for c in w):
        return s
    return s[0].lower() + s[1:]

_MULETILLAS = [
    re.compile(r"^\s*siendo\s+las\s+\d{1,2}[:.h]\d{2}\s*(?:hrs?\.?|horas)?\s*,?\s*", re.I),
    re.compile(r"^\s*(?:durante|en)\s+(?:la|el|una|un)\s+(?:inspecci[oó]n|recorrido|ronda|visita|"
               r"verificaci[oó]n|auditor[ií]a)\b[^,]{0,60},\s*", re.I),
    re.compile(r"^\s*(?:adem[aá]s|asimismo|tambi[eé]n|por otro lado|por otra parte)\s*,?\s*", re.I),
    re.compile(r"^\s*se\s+(?:observ[aó]|detect[aó]|identific[aó]|constat[aó]|evidenci[aó]|aprecia|verific[aó]|"
               r"visualiz[aó]|advierte|comprueba|reporta|informa|nota)(?:n)?\s+que\s+", re.I),
    re.compile(r"^\s*se\s+(?:observ[aó]|detect[aó]|identific[aó]|constat[aó]|evidenci[aó]|aprecia|verific[aó]|"
               r"visualiz[aó]|advierte|comprueba)(?:n)?\s+", re.I),
    re.compile(r"^\s*(?:hallazgo|observaci[oó]n|descripci[oó]n|evento)\s*:\s*", re.I),
]

def _titulo_limpio(titulo, max_chars=60):
    """Título sin prefijo de tipo ni 'CdP', sin nombres, en minúscula inicial y corto."""
    t = str(titulo or "")
    t = re.sub(r"^\s*(?:ACC|CAAP|CA|CBP|HAP|HBP)\s*[-–—:.]*\s*", "", t, flags=re.I)
    t = re.sub(r"\bCdP\b[\s:.\-–]*", "", t, flags=re.I)
    t = _sin_nombres(t)
    t = re.sub(r"\s+", " ", t).strip(" -–:.,;")
    return _cortar(_minuscula_inicial(t), max_chars) if t else ""

def _frase_corta(texto, max_chars=160):
    """Frase de evidencia: sin nombres y con corte en palabra completa."""
    return _cortar(_sin_nombres(texto), max_chars)

def _evidencia(ev):
    """Frase de la descripción (primera oración con contenido) o el título."""
    for o in _oraciones(ev.get("descripcion")):
        if len(o.split()) >= 3:
            return _frase_corta(o)
    return _frase_corta(ev.get("titulo") or "")

# ----------------------------------------------------------------------
# Asignación de tema
# ----------------------------------------------------------------------

def asignar_tema(evento, catalogo=CATALOGO):
    """Puntúa cada tema con título (x2) y descripción (x1); nunca usa 'riesgo'."""
    titulo = normalizar(evento.get("titulo"))
    desc = normalizar(evento.get("descripcion"))
    ambos = titulo + " " + desc
    res = []
    for orden, t in enumerate(catalogo):
        excl = _rx_union(t.get("excluir"))
        if excl is not None and excl.search(ambos):
            continue
        puntaje, pal = 0, []
        for pat in t.get("palabras") or []:
            rx = _rx(pat)
            m_t = rx.search(titulo)
            m_d = rx.search(desc)
            if m_t:
                puntaje += PESO_TITULO
                pal.append(m_t.group(0))
            if m_d:
                puntaje += PESO_DESC
                if not m_t:
                    pal.append(m_d.group(0))
        for pat, pts in t.get("bonus") or []:
            m_b = _rx(pat).search(ambos)
            if m_b:
                puntaje += pts
                pal.append(m_b.group(0))
        if puntaje > 0:
            res.append((puntaje, orden, t["id"], pal))
    res.sort(key=lambda r: (-r[0], r[1]))
    candidatos = [(r[2], r[0]) for r in res[:5]]
    if not res or res[0][0] < UMBRAL:
        return {"tema_id": None, "puntaje": res[0][0] if res else 0, "candidatos": candidatos,
                "estado": "sin_tema", "palabras": list(dict.fromkeys(res[0][3])) if res else []}
    mejor = res[0]
    palabras = list(dict.fromkeys(p.strip() for p in mejor[3] if p.strip()))
    if len(res) > 1 and res[1][0] >= RATIO_AMBIGUO * mejor[0]:
        return {"tema_id": None, "puntaje": mejor[0], "candidatos": candidatos,
                "estado": "ambiguo", "palabras": palabras}
    return {"tema_id": mejor[2], "puntaje": mejor[0], "candidatos": candidatos,
            "estado": "ok", "palabras": palabras}

def _decision_tema(ev, res, catalogo):
    """Decisión visible para el usuario: evento sin tema o con tema ambiguo."""
    nombres = {t["id"]: t["tema"] for t in catalogo}
    ev_id = str(ev.get("id"))
    if res["estado"] == "ambiguo":
        mejor = res["candidatos"][0][1] if res["candidatos"] else 0
        ids = [i for i, p in res["candidatos"] if p >= RATIO_AMBIGUO * mejor][:4]
        clase = "tema_ambiguo"
        pregunta = f"El evento {ev_id} ({ev.get('tipo')}) calza con varios temas. ¿Cuál corresponde?"
    else:
        cand = [i for i, _ in res["candidatos"]]
        ids = cand + [t["id"] for t in catalogo if t["id"] not in cand]
        clase = "sin_tema"
        pregunta = f"El evento {ev_id} ({ev.get('tipo')}) no calza con ningún tema del catálogo. ¿A qué tema pertenece?"
    opciones = [{"valor": i, "etiqueta": nombres[i]} for i in ids if i in nombres]
    opciones.append({"valor": "ignorar", "etiqueta": "Ignorar (no forma foco ni tendencia)"})
    return {"id": ev_id, "clase": clase, "tipo": ev.get("tipo"), "titulo": ev.get("titulo") or "",
            "pregunta": pregunta, "opciones": opciones, "evidencia": _evidencia(ev), "defecto": "ignorar"}

def asignar_todos(eventos, manual=None, catalogo=CATALOGO):
    """Asigna tema a todos; `manual` {id: tema_id} gana siempre. Devuelve por_id y pendientes."""
    manual = {str(k): v for k, v in (manual or {}).items()}
    ids_cat = {t["id"] for t in catalogo}
    por_id, pendientes = {}, []
    for ev in eventos:
        eid = str(ev.get("id"))
        res = asignar_tema(ev, catalogo)
        if eid in manual:
            m = manual[eid]
            previo = res.get("tema_id")
            res = dict(res)
            res["manual"] = True
            if m in ids_cat:
                res["puntaje"] = dict(res["candidatos"]).get(m, 0)
                res["palabras"] = res["palabras"] if previo == m else []
                res["tema_id"], res["estado"] = m, "ok"
            else:
                res["tema_id"], res["estado"], res["ignorado"] = None, "sin_tema", True
        elif res["estado"] in ("sin_tema", "ambiguo"):
            pendientes.append(_decision_tema(ev, res, catalogo))
        por_id[eid] = res
    orden = {str(e.get("id")): GRAVEDAD.get(e.get("tipo"), 0) for e in eventos}
    pendientes.sort(key=lambda d: -orden.get(d["id"], 0))
    return {"por_id": por_id, "pendientes": pendientes}

# ----------------------------------------------------------------------
# Revalorización (solo sugerencia)
# ----------------------------------------------------------------------
_DEF_PERSONAS = (r"\b(trabajador(es|a|as)?|operador(es|a)?|rigger(s)?|supervisor(es|a)?|ayudante(s)?|maestro(s)?|"
                 r"soldador(es)?|mecanico(s)?|electrico(s)?|electricista(s)?|jornal(es)?|conductor(es|a)?|"
                 r"chofer(es)?|personal|persona(s)?|colaborador(es|a)?|prevencionista|capataz|senalero(s)?|"
                 r"andamiero(s)?|gruero(s)?|eslingador(es)?|obrero(s)?|tecnico(s)?|instalador(es)?|ingeniero|"
                 r"jefe de (turno|area|terreno|obra))\b")
_DEF_EXPOSICIONES = {
    "linea_de_fuego": (r"linea de fuego|retroceso|aculat|marcha atras|atropell|equipo en movimiento|"
                       r"vehiculo en movimiento|camion en movimiento|maquina en movimiento|"
                       r"(cruza\w*|transita\w*|circula\w*) (frente|delante|por delante) (de|al) (equipo|camion|vehiculo|maquina)"),
    "carga_suspendida": (r"carga suspendida|bajo (la |una )?carga|debajo de (la |una )?carga|carga izada|"
                         r"bajo (la )?pluma|bajo el gancho|area de influencia de la carga|radio de (giro|la carga)"),
    "energia": (r"energizad|energia (peligrosa|electrica|residual)|\d+ ?kv|sin bloqueo|presurizad|con presion|"
                r"partes? vivas?|contacto electrico|arco electrico|cables? (pelado|sin aislacion|energizado)"),
    "altura": (r"trabajos? en altura|en altura|sin arnes|sin linea de vida|caida a distinto nivel|sin baranda|"
               r"sin proteccion contra caidas|sin doble gancho|sin anclaje"),
    "zanja": (r"(zanja|excavacion)[^.]{0,80}(sin (entibacion|proteccion|permiso|autorizacion|talud|escala|baranda)|desprotegid)|"
              r"sin (entibacion|proteccion|permiso|autorizacion)[^.]{0,60}(zanja|excavacion)|"
              r"ingres\w* (a|en) (la |una )?(zanja|excavacion)|dentro de (la |una )?(zanja|excavacion)|"
              r"interior de (la |una )?(zanja|excavacion)"),
    "atmosfera": (r"espacio confinado|atmosfera (peligrosa|explosiva|deficiente)|sin (medicion|monitoreo) de "
                  r"(gases|atmosfera|oxigeno)|\bgas(es)?\b|monoxido|deficiencia de oxigeno|h2s|sulfhidrico"),
}
_DEF_TAREA = (r"\b(mientras|al momento de|en plena|en pleno|trabajando|trabajaba|trabajaban|realizaba|realizaban|"
              r"ejecutaba|ejecutaban|efectuaba|efectuaban|operando|operaba|izando|izaba|soldando|excavando|"
              r"transitaba|circulaba|desarrollaba|instalando|montando|desmontando|cargando|descargando|"
              r"retrocediendo|manipulando|en ejecucion|en curso|durante (el|la) (izaje|maniobra|trabajo|tarea|"
              r"montaje|excavacion|traslado|operacion))\b")
_DEF_POTENCIAL = (r"\b(podria(n)?|pudo|pudiera|pudieron|hubiera|habria|eventual(es|mente)?|potencial(mente)?|"
                  r"posible(s)?|de haber|en caso de|susceptible)\b|riesgo de|probabilidad")
_DEF_SIN_PERSONA = (r"sin (personal|personas|presencia de (personal|personas))|no (habia|se encontraba|se encontraban|"
                    r"hubo) (personal|personas|trabajadores)|sin exposicion|sin trabajos en curso|sin actividad|"
                    r"no se (estaban|estaba) ejecutando")
_DEF_EROSION = (r"sin (barandas?|arnes|linea de vida|bloqueo|candado|permiso|plan|rigger|radio|segregacion|"
                r"senaliz\w*|certificad\w*|inspeccion|check ?list|supervis\w*|autorizacion|entibacion|medicion|"
                r"extintor|vigia)|no (cuenta|contaba|poseia|tenia|implement\w*|se (verifico|realizo|ejecuto|aplico|"
                r"instalo))|ausencia de|falta de|incompleto|vencid\w*|deteriorad\w*|inoperante|omit\w*|no vigente|"
                r"sin vigencia|no estaba (implementado|instalado|habilitado)")
_DEF_INTERACCION = (r"se acerc\w+|cerca de|junto a|proximidad|transitaba|circulaba|pasaba|paso por|cruzo|cruzaba|"
                    r"ingreso|ingresaba|manipulaba|se ubicaba|a pocos metros|contacto con|se encontraba (cerca|junto|proximo)")
_DEF_CONDICION = (r"deficien\w*|dan[ao]d\w*|suelto|expuest\w*|desprotegid\w*|sin proteccion|derram\w*|obstruid\w*|"
                  r"inestable|resbal\w*|punzante|caliente|irregular|desnivel")
_ETIQUETA_EXPO = {
    "linea_de_fuego": "línea de fuego de equipo o vehículo en movimiento",
    "carga_suspendida": "carga suspendida",
    "energia": "energía peligrosa energizada",
    "altura": "altura sin protección contra caídas",
    "zanja": "zanja o excavación sin protección o autorización",
    "atmosfera": "atmósfera peligrosa o espacio confinado sin control",
}

_ALIAS_REVAL = {"personas": "persona", "potencial": "negacion_potencial", "control_erosionado": "control_critico_erosionado",
                "interaccion": "interaccion_cercana"}

def _reval(clave, defecto):
    v = None
    if isinstance(REVAL, dict):
        v = REVAL.get(clave) or REVAL.get(_ALIAS_REVAL.get(clave, clave))
    return _rx_union(v) or _rx_union(defecto)

def _exposiciones():
    v = REVAL.get("exposiciones") if isinstance(REVAL, dict) else None
    if not isinstance(v, dict) or not v:
        v = _DEF_EXPOSICIONES
    out = []
    for nombre, pat in v.items():
        if isinstance(pat, dict):
            pat = pat.get("regex")
        rx = _rx_union(pat)
        if rx is not None:
            out.append((nombre, rx))
    return out

_CC_INVALIDOS = {"", "otro", "otros", "sin informacion", "ninguno", "no aplica", "n/a", "na", "-",
                 "sin control", "none", "nan"}

def _control_cc(ev):
    """Nombre del control crítico específico del evento, o None (Otro / Sin Información / vacío)."""
    raw = str(ev.get("control_critico") or "")
    if not raw.strip():
        raw = str(ev.get("cc") or "")
    txt = re.sub(r"\(\s*C\.?\s*C\.?\s*\)", "", raw, flags=re.I)
    txt = re.split(r"[;\n|]", txt)[0]
    txt = re.sub(r"\s+", " ", txt).strip(" -–:.,")
    n = normalizar(txt)
    if n in _CC_INVALIDOS or n.startswith("otro"):
        return None
    return _cortar(txt, 60) or None

def _sugerencia(ev_id, de, a, fr, motivo, conf):
    return {"id": ev_id, "de": de, "a": a, "frase": _frase_corta(fr, 200), "motivo": motivo, "confianza": conf}

def sugerir_revalorizacion(evento):
    """Sugiere subir el tipo del evento (solo sugerencia; conservador). None si no hay base clara."""
    de = evento.get("tipo")
    if de not in TRANSICIONES:
        return None
    desc = (evento.get("descripcion") or "").strip()
    if not desc:
        return None
    ev_id = str(evento.get("id"))
    frases = _oraciones(desc)
    norm = [normalizar(f) for f in frases]
    texto_norm = " ".join(norm)
    rx_pers = _reval("personas", _DEF_PERSONAS)
    rx_tarea = _reval("tarea_en_curso", _DEF_TAREA)
    rx_pot = _reval("potencial", _DEF_POTENCIAL)
    rx_sinp = _reval("sin_persona", _DEF_SIN_PERSONA)
    rx_ero = _reval("control_erosionado", _DEF_EROSION)
    expo = _exposiciones()
    tarea_global = bool(rx_tarea.search(texto_norm))

    # 1) A CA: persona efectivamente expuesta a una de las 6 exposiciones
    mejor = None
    for k_fr, (fr, n) in enumerate(zip(frases, norm)):
        ctx_pers = rx_pers.search(n) or (k_fr > 0 and rx_pers.search(norm[k_fr - 1]))   # persona en la oración o en la anterior
        if rx_sinp.search(n) or not ctx_pers:
            continue
        exps = [nom for nom, rx in expo if rx.search(n)]
        if not exps:
            continue
        pot = bool(rx_pot.search(n))
        if pot and not tarea_global:
            continue                                  # "podría haber pasado": no se revaloriza
        alta = tarea_global and not pot
        cand = (alta, len(exps), fr, exps, pot)
        if mejor is None or (cand[0], cand[1]) > (mejor[0], mejor[1]):
            mejor = cand
    if mejor:
        alta, _, fr, exps, pot = mejor
        motivo = "persona expuesta a " + _unir([_ETIQUETA_EXPO.get(e, e.replace("_", " ")) for e in exps])
        motivo += "; tarea en curso" if tarea_global else "; sin tarea en curso explícita"
        if pot:
            motivo += "; redacción potencial"
        return _sugerencia(ev_id, de, "CA", fr, motivo, "alta" if alta else "baja")
    if de != "HBP":
        return None

    # 2) HBP -> HAP: control crítico no implementado o erosionado mientras se ejecutaba una tarea de riesgo crítico
    expo_global = [nom for nom, rx in expo if rx.search(texto_norm)]
    if expo_global and tarea_global:
        cc_esp = _control_cc(evento) is not None
        cand_alta = cand_baja = None
        for fr, n in zip(frases, norm):
            if not rx_ero.search(n):
                continue
            local = any(rx.search(n) for _, rx in expo) or bool(rx_tarea.search(n))
            pot = bool(rx_pot.search(n))
            if pot and not local:
                continue
            if not (cc_esp or local):
                continue
            if cc_esp and local and not pot:
                cand_alta = cand_alta or fr
            else:
                cand_baja = cand_baja or fr
        fr = cand_alta or cand_baja
        if fr:
            motivo = ("control crítico no implementado o erosionado durante tarea de riesgo crítico ("
                      + _unir([_ETIQUETA_EXPO.get(e, e.replace("_", " ")) for e in expo_global]) + ")")
            return _sugerencia(ev_id, de, "HAP", fr, motivo, "alta" if cand_alta else "baja")

    # 3) HBP -> CBP: una persona interactuó de cerca con la situación de riesgo
    rx_int = _reval("interaccion", _DEF_INTERACCION)
    rx_cond = _reval("condicion_riesgo", _DEF_CONDICION)
    cand_alta = cand_baja = None
    for fr, n in zip(frases, norm):
        if rx_sinp.search(n) or not rx_pers.search(n) or not rx_int.search(n):
            continue
        if not (rx_cond.search(n) or rx_ero.search(n)):
            continue
        pot = bool(rx_pot.search(n))
        tarea_loc = bool(rx_tarea.search(n))
        if pot and not tarea_loc:
            continue
        if tarea_loc and not pot:
            cand_alta = cand_alta or fr
        else:
            cand_baja = cand_baja or fr
    fr = cand_alta or cand_baja
    if fr:
        return _sugerencia(ev_id, de, "CBP", fr,
                           "persona interactuó de cerca con la situación de riesgo, sin exposición a riesgo crítico",
                           "alta" if cand_alta else "baja")
    return None

def _decision_rev(ev, sug):
    """Decisión visible para una revalorización sugerida y aún sin decidir."""
    de, a = sug["de"], sug["a"]
    return {"id": str(ev.get("id")), "clase": "revalorizacion", "tipo": ev.get("tipo"),
            "titulo": ev.get("titulo") or "",
            "pregunta": (f"¿Revalorizar el evento {ev.get('id')} de {de} a {a}? "
                         f"Motivo: {sug['motivo']} (confianza {sug['confianza']})."),
            "opciones": [{"valor": "mantener", "etiqueta": f"Mantener {de}"},
                         {"valor": a, "etiqueta": f"Revalorizar a {a}"}],
            "evidencia": sug["frase"], "defecto": "mantener"}

# ----------------------------------------------------------------------
# Resumen de evento
# ----------------------------------------------------------------------

def _tema_por_id(tema):
    if isinstance(tema, dict):
        return tema
    if isinstance(tema, str):
        for t in CATALOGO:
            if t.get("id") == tema:
                return t
    return None

def resumir_evento(evento, tema=None, max_chars=120):
    """Frase concreta (minúscula inicial, sin nombres) elegida de la descripción según el tema."""
    t = _tema_por_id(tema)
    desc = (evento.get("descripcion") or "").strip()
    try:
        import resumen_texto
        texto_n = normalizar((evento.get("titulo") or "") + " " + desc)
        palabras = []
        for p_ in ((t or {}).get("palabras") or []):
            m_ = _rx(p_).search(texto_n)
            if m_:
                palabras.append(m_.group(0).strip())
        frase = resumen_texto.resumir(_sin_nombres(str(evento.get("titulo") or "")), _sin_nombres(desc),
                                      palabras, max_chars=max(max_chars - 1, 1))
        if frase:
            frase = frase.strip().rstrip(".")
            return frase + "." if frase else ""
    except ImportError:
        pass
    elegida = ""
    orac = _oraciones(desc)
    if orac:
        rxs = [_rx(p) for p in ((t or {}).get("palabras") or [])]
        mejor = -1
        for o in orac:
            if len(o.split()) < 3:
                continue
            n = normalizar(o)
            pts = sum(1 for rx in rxs if rx.search(n))
            if pts > mejor:
                mejor, elegida = pts, o
        if not elegida:
            elegida = orac[0]
    if not elegida:
        elegida = re.sub(r"^\s*(?:ACC|CAAP|CA|CBP|HAP|HBP)\s*[-–—:.]*\s*", "", str(evento.get("titulo") or ""),
                         flags=re.I)
    s = _sin_nombres(elegida)
    for rx in _MULETILLAS:
        s = rx.sub("", s, count=1)
    s = re.sub(r"\s+", " ", s).strip(" ,;:-–")
    s = _minuscula_inicial(s)
    s = _cortar(s, max(max_chars - 1, 1))
    return s + "." if s else ""

# ----------------------------------------------------------------------
# Armado del plan
# ----------------------------------------------------------------------

def _tokens(texto):
    return {w for w in re.findall(r"[a-z0-9]+", normalizar(texto)) if len(w) > 3}

def _similares(a, b):
    if not a or not b:
        return False
    return len(a & b) / len(a | b) >= 0.6

def _top_areas(evs, n=3):
    """Las n áreas (area_geo exacto) más frecuentes; ignora vacías y 'SIN AREA'."""
    cont = {}
    for e in evs:
        a = (e.get("area_geo") or "").strip()
        if not a or normalizar(a) == "sin area":
            continue
        cont[a] = cont.get(a, 0) + 1
    return [a for a, _ in sorted(cont.items(), key=lambda x: (-x[1], x[0]))[:n]]

def _componer_instruccion(t, texto_norm):
    """Instrucción base + cláusulas extra que apliquen, sin pasar de 350 caracteres."""
    base = (t.get("instruccion") or "").strip()
    for item in t.get("extras") or []:
        try:
            patron, clausula = item[0], item[1]
        except (TypeError, IndexError, KeyError):
            continue
        rx = _rx_union(patron)
        if rx is None or not rx.search(texto_norm):
            continue
        cl = str(clausula).strip().rstrip(".;, ")
        if not cl or normalizar(cl) in normalizar(base):
            continue
        cand = base.rstrip(". ") + "; " + cl + "."
        if len(cand) <= MAX_INSTRUCCION:
            base = cand
    return base

def _nota_abiertos(ids, tipos):
    lead = "Evento" if len(ids) == 1 else "Eventos"
    tipo_txt = f" {next(iter(tipos))}" if len(tipos) == 1 else ""
    return f"({lead}{tipo_txt} {_unir(ids)}, {'abierto' if len(ids) == 1 else 'abiertos'} a la fecha de cierre.)"

# ---------------------------------------------------------------------------
# FOCO FIJO (prioridad 1, N1) hasta el cierre del año. Para cambiarlo o apagarlo basta editar este bloque.
# Los eventos de los temas indicados se listan como evidencia cuando existen; si no hay, la fila igual sale.
# ---------------------------------------------------------------------------
FOCO_FIJO = {
    "vigente_hasta": "2026-12-31",           # se aplica si la vigencia del turno empieza hasta esta fecha
    "tema_id": "tableros_electricos",         # tema del catálogo al que se asocia
    "temas_asociados": ("generadores_equipos_menores",),   # sus eventos también se listan en este foco
    "tema": "Controles eléctricos en tableros, generadores y torres.",
    "instruccion": ("Verificar antes de energizar o usar equipos eléctricos que el tablero tenga diferencial operativo "
                    "(sin puentes), que generadores y torres de iluminación tengan puesta a tierra y segregación, "
                    "que la alimentación pase por tablero de distribución y que la soldadora tenga inspección vigente; "
                    "el supervisor lo valida."),
    "criterio": "no se energiza ni se usa equipo con diferencial puenteado, sin tierra o sin tablero intermedio.",
    "herramienta": ("ASP + Inspección de Equipos + Acompañamiento ASP: revisión de tableros, generadores y torres "
                    "al inicio del turno; ASP antes de cada tarea eléctrica."),
    "evidencia": "Diferencial operativo / generador y torre con tierra / tablero intermedio / registro fotográfico",
    "sin_eventos": "Foco fijo del año: verificación permanente (sin eventos N1/N2 de este tema en el período).",
}


def _foco_fijo_activo(params):
    vig = (params or {}).get("vigencia") or {}
    desde = str(vig.get("desde") or "")[:10]
    return bool(desde) and desde <= FOCO_FIJO["vigente_hasta"]


def armar_plan(analisis, params, decisiones=None, calcular=None):
    """Arma el `plan` para `generar` a partir de la salida de `analizar`. Devuelve plan, qa, decisiones y log."""
    if not isinstance(analisis, dict) or not analisis.get("ok"):
        raise ValueError("analisis inválido o con ok=False")
    params = params or {}
    decisiones = decisiones or {}
    if calcular is None:
        try:
            from assurance_engine import calcular_nivel_y_prioridad as calcular
        except Exception as e:  # pragma: no cover
            raise ValueError(f"falta la función calcular: {e}")

    log, supuestos, faltantes = [], [], []
    catalogo = CATALOGO
    cat = {t["id"]: t for t in catalogo}
    orden_cat = {t["id"]: i for i, t in enumerate(catalogo)}

    eventos = list(analisis.get("eventos") or [])
    for e in eventos:
        e["id"] = str(e.get("id"))
    validos = [e for e in eventos if e.get("tipo") in TIPOS_VALIDOS]
    sin_tipo = [e["id"] for e in eventos if e.get("tipo") not in TIPOS_VALIDOS]
    by_id = {e["id"]: e for e in validos}
    if sin_tipo:
        faltantes.append(f"Eventos sin tipo reconocido (no se usan): {', '.join(sin_tipo)}")
    por_tipo = {t: sum(1 for e in validos if e["tipo"] == t) for t in TIPOS_VALIDOS}

    # --- decisiones del usuario -------------------------------------------
    dec_tema = {str(k): v for k, v in (decisiones.get("tema") or {}).items()}
    dec_rev = {}
    for k, v in (decisiones.get("revalorizar") or {}).items():
        u = str(v).strip()
        u = "CA" if u.upper() == "CAAP" else (u.upper() if u.upper() in ("CA", "CBP", "HAP") else u.lower())
        dec_rev[str(k)] = u
    sin_foco = {str(x) for x in (decisiones.get("sin_foco") or [])}

    # --- 1. tema para todos -----------------------------------------------
    asig = asignar_todos(validos, dec_tema, catalogo)
    por_ev = asig["por_id"]
    pend_tema = asig["pendientes"]
    n_ok = sum(1 for r in por_ev.values() if r["estado"] == "ok")
    log.append(f"✔ Temas asignados: {n_ok} de {len(validos)} eventos; "
               f"{sum(1 for p in pend_tema if p['clase'] == 'tema_ambiguo')} ambiguos, "
               f"{sum(1 for p in pend_tema if p['clase'] == 'sin_tema')} sin tema")
    for k in dec_tema:
        if k not in by_id:
            log.append(f"⚠ Decisión de tema para un ID que no está en el período: {k}")

    # --- 2. revalorizaciones ----------------------------------------------
    rev_aplicadas, rev_info, rev_pend, dec_rev_pend = {}, [], [], []
    for ev in validos:
        eid = ev["id"]
        sug = sugerir_revalorizacion(ev)
        dec = dec_rev.get(eid)
        if dec and dec != "mantener":
            if dec in TRANSICIONES.get(ev["tipo"], ()):
                rev_aplicadas[eid] = dec
                rev_info.append({"id": eid, "de": ev["tipo"], "a": dec,
                                 "frase": sug["frase"] if sug else _evidencia(ev),
                                 "sugerida_por_el_motor": bool(sug and sug["a"] == dec)})
            else:
                log.append(f"⚠ Revalorización {ev['tipo']}→{dec} del evento {eid} no es una transición válida: se ignora")
        elif dec == "mantener":
            continue
        elif sug:
            rev_pend.append({"id": eid, "de": sug["de"], "a": sug["a"], "confianza": sug["confianza"],
                             "motivo": sug["motivo"], "frase": sug["frase"]})
            dec_rev_pend.append(_decision_rev(ev, sug))
    for k in dec_rev:
        if k not in by_id:
            log.append(f"⚠ Decisión de revalorización para un ID que no está en el período: {k}")
    log.append(f"✔ Revalorizaciones: {len(rev_aplicadas)} aplicadas, {len(rev_pend)} sugeridas sin decidir (no se aplican)")

    def tipo_ef(ev):
        t = rev_aplicadas.get(ev["id"])
        if t:
            return "CAAP" if t == "CA" else t
        return ev["tipo"]

    def tipo_plan(ev):
        t = rev_aplicadas.get(ev["id"])
        if not t:
            return ev["tipo"], False
        return t, t in ("CBP", "HAP")

    # --- 3. agrupar por tema ----------------------------------------------
    grupos = {}
    for ev in validos:
        tid = por_ev.get(ev["id"], {}).get("tema_id")
        if not tid or ev["id"] in sin_foco:
            continue
        grupos.setdefault(tid, []).append(ev)
    if sin_foco:
        supuestos.append("Eventos marcados 'sin_foco' no forman foco ni Recurrente, pero sí cuentan en Tendencias: "
                         + ", ".join(sorted(sin_foco)))

    def ev_calc(ev):
        d = {"id": ev["id"], "tipo": ev["tipo"], "dm": bool(ev.get("dm")),
             "fecha": ev.get("fecha") or "", "en_turno": ev.get("en_turno", True)}
        if ev["id"] in rev_aplicadas:
            d["tipo_rev"] = rev_aplicadas[ev["id"]]
        return d

    def _tiene_n1n2(tid):
        return any(tipo_ef(e) != "HBP" for e in grupos[tid])

    fijo_activo = _foco_fijo_activo(params) and FOCO_FIJO["tema_id"] in cat
    FIJO = FOCO_FIJO["tema_id"]
    if fijo_activo:
        for otro in FOCO_FIJO["temas_asociados"]:
            if otro in grupos:
                grupos.setdefault(FIJO, []).extend(grupos.pop(otro))
    focos_in = [{"tema": tid, "eventos": [ev_calc(e) for e in grupos[tid]],
                 "cc_o_emergencia": bool(cat[tid].get("cc_o_emergencia"))}
                for tid in sorted(grupos, key=lambda x: orden_cat.get(x, 999))
                if _tiene_n1n2(tid) or not cat[tid].get("no_foco_solo_hbp")]
    res_calc = calcular(focos_in)
    log.extend(res_calc.get("log") or [])
    if fijo_activo:
        lista = list(res_calc.get("focos") or [])
        fx = next((f for f in lista if f["tema"] == FIJO), None)
        otros = [f for f in lista if f["tema"] != FIJO]
        if fx is None:
            fx = {"tema": FIJO, "eventos": [ev_calc(e) for e in grupos.get(FIJO, [])]}
        fx = {**fx, "nivel": "N1", "motivo_nivel": "Foco fijo hasta el 31-12-2026", "fijo": True}
        fuera = list(res_calc.get("fuera") or [])
        for f in otros[4:]:
            fuera.append({**f, "motivo": "supera el máximo de 5 focos (el foco 1 es el fijo)"})
        nuevos = [fx] + otros[:4]
        for i, f in enumerate(nuevos, 1):
            f["prioridad"] = i
        res_calc = {**res_calc, "focos": nuevos, "fuera": fuera}
        supuestos.append("El foco 1 es fijo hasta el 31-12-2026 (controles eléctricos); los demás se calculan como siempre, "
                         "con un máximo de 4.")

    # --- 4. construir los focos del plan ----------------------------------
    focos_plan, grupo_por_prioridad, qa_focos = [], {}, []
    ya_anteriores = []
    for f in res_calc.get("focos") or []:
        tid = f["tema"]
        t = cat[tid]
        if f.get("fijo"):
            t = {**t, "tema": FOCO_FIJO["tema"].rstrip("."), "instruccion": FOCO_FIJO["instruccion"],
                 "criterio": FOCO_FIJO["criterio"], "herramienta": FOCO_FIJO["herramienta"],
                 "evidencia": FOCO_FIJO["evidencia"], "extras": []}
        evs = [by_id[e["id"]] for e in f["eventos"]]
        nn = [e for e in evs if tipo_ef(e) != "HBP"]
        hb = [e for e in evs if tipo_ef(e) == "HBP"]
        nn.sort(key=lambda e: (e.get("fecha") or "", e["id"]), reverse=True)
        nn.sort(key=lambda e: -GRAVEDAD.get(tipo_ef(e), 0))
        hb.sort(key=lambda e: (e.get("fecha") or "", e["id"]), reverse=True)
        hb.sort(key=lambda e: 0 if e.get("en_turno", True) else 1)

        eventos_plan = []
        for e in nn:
            tp, rv = tipo_plan(e)
            texto = resumir_evento(e, t)
            if not texto:
                faltantes.append(f"Evento {e['id']} sin descripción ni título: texto vacío")
            eventos_plan.append({"tipo": tp, "ids": [e["id"]], "empresa": e.get("empresa"), "texto": texto,
                                 "fecha": e.get("fecha"), "dm": bool(e.get("dm")), "rev": rv,
                                 "en_turno": bool(e.get("en_turno", True))})
        # HBP complementarios: se juntan los que dicen lo mismo (misma empresa) y se listan máx. 3 IDs
        grupos_hbp = []
        for e in hb:
            texto = resumir_evento(e, t)
            tok = _tokens(texto)
            for g in grupos_hbp:
                if g["empresa"] == e.get("empresa") and (texto == g["texto"] or _similares(tok, g["tokens"])):
                    g["evs"].append(e)
                    break
            else:
                grupos_hbp.append({"empresa": e.get("empresa"), "texto": texto, "tokens": tok, "evs": [e]})
        usados, hbp_ids = 0, []
        for g in grupos_hbp:
            libres = MAX_HBP_LISTADOS - usados
            if libres <= 0:
                break
            sel = g["evs"][:libres]
            usados += len(sel)
            hbp_ids.extend(e["id"] for e in sel)
            if not g["texto"]:
                faltantes.append(f"Evento {sel[0]['id']} sin descripción ni título: texto vacío")
            eventos_plan.append({"tipo": "HBP", "ids": [e["id"] for e in sel], "empresa": g["empresa"],
                                 "texto": g["texto"], "fecha": max((e.get("fecha") or "") for e in sel),
                                 "dm": False, "rev": False,
                                 "en_turno": any(e.get("en_turno", True) for e in sel)})
        hbp_mas = len(hb) - usados

        texto_norm = normalizar(" ".join(f"{e.get('titulo') or ''} {e.get('descripcion') or ''}" for e in evs))
        instr = _componer_instruccion(t, texto_norm)
        foco = {"nivel": f["nivel"], "tema": t["tema"], "instruccion": instr, "criterio": t.get("criterio"),
                "eventos": eventos_plan, "herramienta": t.get("herramienta"), "evidencia": t.get("evidencia"),
                "tema_id": tid, "prioridad": f["prioridad"]}
        if f.get("fijo"):
            foco["fijo"] = True
            if not eventos_plan:
                foco["sin_eventos_texto"] = FOCO_FIJO["sin_eventos"]
        if hbp_mas > 0:
            foco["hbp_mas"] = hbp_mas
        if len(evs) == 1 and tipo_ef(evs[0]) in ("ACC", "CAAP"):
            foco["aislado"] = True
        focos_plan.append(foco)
        grupo_por_prioridad[f["prioridad"]] = evs
        for e in evs:
            if e.get("en_assurance_anterior"):
                ya_anteriores.append(e["id"])
        qa_focos.append({"prioridad": f["prioridad"], "tema": t["tema"], "nivel": f["nivel"],
                         "motivo_nivel": f.get("motivo_nivel"),
                         "n1n2": [e["id"] for e in nn], "hbp_listados": hbp_ids, "hbp_mas": max(hbp_mas, 0),
                         "hbp_total": len(hb)})
    plan_tids = {f["tema_id"] for f in focos_plan}
    log.append(f"✔ Focos armados: {len(focos_plan)}")

    # N1/N2 fuera con motivo
    def info_fuera(x):
        return {"tema": cat[x["tema"]]["tema"] if x.get("tema") in cat else x.get("tema"),
                "ids": [e["id"] for e in x.get("eventos") or []], "n1n2": x.get("n1n2", 0),
                "hbp": x.get("hbp", 0), "motivo": x.get("motivo")}
    fuera_todos = [info_fuera(x) for x in (res_calc.get("fuera") or [])] + \
                  [info_fuera(x) for x in (res_calc.get("excluidos") or [])]
    n1n2_fuera = [x for x in fuera_todos if x["n1n2"] > 0]
    solo_hbp_excl = [x for x in fuera_todos if x["n1n2"] == 0]
    if res_calc.get("fuera"):
        supuestos.append("Los temas que quedaron fuera por el máximo de 5 focos no cuentan como foco para Recurrente.")

    # --- 5. recurrente -----------------------------------------------------
    cont_hbp = {}
    for ev in validos:
        tid = por_ev.get(ev["id"], {}).get("tema_id")
        if tid and tid not in plan_tids and ev["id"] not in sin_foco and tipo_ef(ev) == "HBP":
            cont_hbp.setdefault(tid, []).append(ev)
    rec_tids = sorted((tid for tid, l in cont_hbp.items() if len(l) >= 2),
                      key=lambda x: (-len(cont_hbp[x]), orden_cat.get(x, 999)))[:3]
    recurrente = {"temas": [{"tema": cat[tid].get("nombre_corto") or cat[tid]["tema"].lower(),
                             "n": len(cont_hbp[tid])} for tid in rec_tids]}

    # --- 7. tendencias -----------------------------------------------------
    cnt_no, cnt_cc, sin_tema_n = {}, {}, 0
    for ev in validos:
        tid = por_ev.get(ev["id"], {}).get("tema_id")
        if not tid:
            sin_tema_n += 1
            continue
        ctrl = _control_cc(ev)
        if ctrl:
            cnt_cc[(tid, ctrl)] = cnt_cc.get((tid, ctrl), 0) + 1
        else:
            cnt_no[tid] = cnt_no.get(tid, 0) + 1
    no_cc = [{"tema": cat[tid]["tema"], "n": n} for tid, n in
             sorted(cnt_no.items(), key=lambda x: (-x[1], orden_cat.get(x[0], 999)))[:5]]
    cc = [{"tema": cat[tid]["tema"], "control": ctrl, "n": n} for (tid, ctrl), n in
          sorted(cnt_cc.items(), key=lambda x: (-x[1], orden_cat.get(x[0][0], 999), x[0][1]))[:5]]
    tendencias = {"no_cc": no_cc, "cc": cc}
    if sin_tema_n:
        supuestos.append(f"{sin_tema_n} evento(s) sin tema no se cuentan en Tendencias.")
    supuestos.append("El control crítico se toma del campo control_critico (el campo cc del motor solo trae la marca (C.C.)); "
                     "'Otro', 'Sin Información' o vacío cuentan como no CC.")

    # --- 8. indicadores y plan --------------------------------------------
    indicadores = copy.deepcopy(analisis.get("indicadores") or {})
    if not indicadores.get("comparacion_texto") or "[DATO FALTANTE]" in str(indicadores.get("comparacion_texto")):
        faltantes.append("Comparación con los 14 días anteriores no disponible (el Excel no cubre ese período)")
    base = analisis.get("params") or {}
    vig = base.get("vigencia") or params.get("vigencia")
    if isinstance(vig, dict):
        vig = {"desde": vig.get("desde"), "hasta": vig.get("hasta")}
    areas_validas = []
    for u in (analisis.get("areas") or {}).get("unicas") or []:
        areas_validas.append(u.get("area_geo") if isinstance(u, dict) else str(u))

    plan = {"turno": base.get("turno") or params.get("turno"), "vigencia": vig,
            "fecha_cierre": base.get("fecha_cierre") or params.get("fecha_cierre"),
            "areas_validas": [a for a in areas_validas if a],
            "indicadores": indicadores, "focos": focos_plan, "recurrente": recurrente,
            "tendencias": tendencias}

    # --- 9. decisiones y QA -------------------------------------------------
    n1n2_sin_tema = [e["id"] for e in validos
                     if e["tipo"] != "HBP" and not por_ev.get(e["id"], {}).get("tema_id")]
    if n1n2_sin_tema:
        faltantes.append("Eventos N1/N2 sin tema (decisión del usuario): " + ", ".join(n1n2_sin_tema))
    decisiones_out = pend_tema + dec_rev_pend
    qa = {
        "eventos_leidos": {"total": len(eventos), "por_tipo": por_tipo, "sin_tipo": sin_tipo},
        "turno_analizado": (analisis.get("turno") or {}).get("texto"),
        "focos": qa_focos,
        "n1n2_fuera": n1n2_fuera,
        "temas_solo_hbp_excluidos": solo_hbp_excl,
        "revalorizaciones_aplicadas": rev_info,
        "revalorizaciones_sin_decidir": rev_pend,
        "eventos_sin_resolver": {
            "sin_tema": [d["id"] for d in pend_tema if d["clase"] == "sin_tema"],
            "tema_ambiguo": [d["id"] for d in pend_tema if d["clase"] == "tema_ambiguo"]},
        "ya_en_assurance_anterior": sorted(set(ya_anteriores)),
        "datos_faltantes": faltantes,
        "supuestos": [
            "El tema se asigna solo con título y descripción (el campo riesgo no se usa).",
            "Sin decisión sobre una revalorización sugerida, el evento mantiene su tipo original.",
            "Los eventos sin tema o con tema ambiguo quedan fuera hasta que el usuario decida (defecto: ignorar).",
        ] + supuestos,
    }
    if ya_anteriores:
        log.append(f"⚠ {len(set(ya_anteriores))} evento(s) de los focos ya se usaron en un Assurance anterior")
    if decisiones_out:
        log.append(f"⚠ {len(decisiones_out)} decisión(es) pendientes para el usuario")
    log.append("✔ Plan armado")
    return {"plan": plan, "qa": qa, "decisiones": decisiones_out, "log": log}
