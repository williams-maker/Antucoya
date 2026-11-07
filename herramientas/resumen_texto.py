"""resumen_texto.py: resume hallazgos para la columna "Evento" de un Assurance.

Solo stdlib (Python 3.12 / Pyodide).

    resumir(titulo, descripcion, palabras_tema, max_chars=120) -> str

Pipeline:
  0. quitar nombres de personas (se conserva el cargo si aparece);
  1. separar la descripción en oraciones;
  2. quitar la cláusula introductoria de contexto si lo que sigue tiene >= 25 chars;
  3. quitar muletillas ("se observa que", "se encontró", "se solicita"...);
  4. puntuar cada oración (tema +, hallazgo +, contexto -);
  5. acortar la mejor cortando en límite de cláusula o de palabra completa,
     sin terminar en preposición/artículo/conector;
  6. si no hay descripción, usar el título sin prefijos ("HBP-", "CBP -", "ASP-HAP:"...).

No se corrige la ortografía de origen (tampoco se introducen errores nuevos).
"""
from __future__ import annotations

import re
import unicodedata

__all__ = ["resumir"]

# --------------------------------------------------------------------------
# Utilidades de texto
# --------------------------------------------------------------------------

def _plano(s: str) -> str:
    """Minúsculas y sin tildes, conservando la longitud (índices alineables)."""
    salida = []
    for ch in s:
        d = "".join(
            c for c in unicodedata.normalize("NFD", ch) if not unicodedata.combining(c)
        ).lower()
        salida.append(d if len(d) == 1 else ch)
    return "".join(salida)

def _espacios(s: str) -> str:
    s = re.sub(r"\s+", " ", s).strip()
    return re.sub(r"\s+([,;:.])", r"\1", s)

# --------------------------------------------------------------------------
# Nombres de personas
# --------------------------------------------------------------------------

_NOM = (
    r"(?!(?:Los|Las|Sin|Del|Con|Por|Que|Una|Durante|Previo|Cuando|Pero|Como|Para)\b)"
    r"[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}"
)
_SEQ = rf"{_NOM}(?:[ \t]+{_NOM}){{0,2}}"

_CARGOS = [
    "jefe de turno", "jefe de obra", "jefe de área", "jefe de area",
    "administrador de contrato", "administrador de obra",
    "supervisor", "supervisora", "rigger", "prevencionista", "capataz",
    "operador", "operadora", "ingeniero", "ingeniera", "jefe", "jefa",
    "encargado", "encargada", "coordinador", "coordinadora", "gerente",
    "residente", "inspector", "inspectora", "maestro", "bodeguero",
    "señalero", "soldador", "electricista", "conductor", "conductora",
    "gruista", "trabajador", "trabajadora", "mecánico", "técnico", "técnica",
    "asesor", "asesora",
]
_ALT_CARGOS = "|".join(re.escape(c) for c in sorted(_CARGOS, key=len, reverse=True))

_NOMBRES = [
    "Juan", "Pedro", "José", "Jose", "María", "Maria", "Luis", "Carlos", "Ana",
    "Jorge", "Miguel", "Andrés", "Andres", "Felipe", "Francisco", "Claudio",
    "Patricio", "Rodrigo", "Cristián", "Cristian", "Sergio", "Pablo", "Diego",
    "Manuel", "Marcelo", "Daniel", "Héctor", "Hector", "Roberto", "Ricardo",
    "Fernando", "Javier", "Mario", "Luisa", "Carla", "Camila", "Paula",
    "Sofía", "Sofia", "Claudia", "Patricia", "Marcela", "Carolina", "Verónica",
    "Veronica", "Gabriela", "Daniela", "Elena", "Isabel", "Alejandro",
    "Eduardo", "Víctor", "Victor", "Raúl", "Raul", "Óscar", "Oscar", "Gonzalo",
    "Hugo", "Ignacio", "Nicolás", "Nicolas", "Tomás", "Tomas", "Matías",
    "Matias", "Sebastián", "Sebastian", "Jaime", "Álvaro", "Alvaro",
    "Mauricio", "César", "Cesar", "Iván", "Ivan",
]
_ALT_NOMBRES = "|".join(_NOMBRES)

_TITULOS_RE = re.compile(rf"\b(?P<t>Sra|Srta|Sr|Dra|Dr|Ing|Don|Doña)\.?[ \t]+{_SEQ}")
_CARGO_RE = re.compile(rf"\b(?P<cargo>(?i:{_ALT_CARGOS}))\b[ \t]+{_SEQ}")
_NOMBRES_RE = re.compile(rf"\b(?:{_ALT_NOMBRES})\b(?:[ \t]+{_NOM}){{0,2}}")

def _sub_titulo(m: re.Match) -> str:
    return {"ing": "ingeniero", "dr": "doctor", "dra": "doctora"}.get(
        m.group("t").lower(), "trabajador"
    )

def _quitar_nombres(t: str) -> str:
    """Quita nombres propios; conserva el cargo; sin cargo, usa 'trabajador'."""
    t = _TITULOS_RE.sub(_sub_titulo, t)
    t = _CARGO_RE.sub(r"\g<cargo>", t)
    return _NOMBRES_RE.sub("trabajador", t)

# --------------------------------------------------------------------------
# Limpieza de oraciones
# --------------------------------------------------------------------------

_MIN_RESTO = 25  # mínimo de caracteres que debe quedar tras quitar contexto

_INTRO = re.compile(
    r"^(?:en\s+(?:la\s+|el\s+|una\s+|un\s+)?(?:inspeccion|recorrido|visita|revision"
    r"|verificacion|terreno|auditoria|ronda|supervision|fiscalizacion|reunion|charla)"
    r"|durante\b|al\s+momento\s+de|previo\s+a|previa\s+a|posterior\s+a|luego\s+de|tras\s"
    r"|con\s+motivo\s+de|como\s+parte\s+de|en\s+el\s+marco\s+de"
    r"|al\s+(?:realizar|efectuar|ejecutar|llegar|ingresar)\b)"
)

_MULETILLA = re.compile(
    r"\bse\s+(?:observ(?:a|an|o|aron)|detect(?:a|an|o|aron)|evidenci(?:a|an|o|aron)"
    r"|constat(?:a|an|o|aron)|verific(?:a|an|o|aron)|comprueba(?:n)?|comprob(?:o|aron)"
    r"|identific(?:a|an|o|aron)|encontr(?:o|aron)|apreci(?:a|an)|advierte(?:n)?)\b"
    r"(?:\s+que\b)?\s*"
)

_ADMIN = re.compile(
    r"^(?:se\s+solicitan?|se\s+informa|cabe\s+(?:destacar|senalar|mencionar|indicar))"
    r"\s+(?:que\s+)?"
)

_CONECTOR_INICIAL = re.compile(r"^(?:sin\s+embargo|ademas|asimismo|tambien|pero|y|e)\b,?\s+")

_AL_RE = re.compile(
    r"^al\s+(?:solicitar|revisar|verificar|consultar|requerir|pedir|inspeccionar"
    r"|chequear|controlar|medir)\s+(?P<obj>.+?),?\s+se\s+"
    r"(?:verifica|observa|detecta|evidencia|constata|comprueba|identifica)(?:n)?"
    r"\s+que\s+(?P<hall>.+)$"
)
_DET_INICIAL = re.compile(r"(?:el|la|los|las|un|una|unos|unas|lo|se)\b")

def _quitar_parentesis(t: str) -> str:
    return re.sub(
        r"\s*\((?=[^()]*(?:foto|anexo|imagen|adjunto|ver\s))[^()]*\)", "", t, flags=re.I
    )

def _quitar_intro(t: str) -> str:
    """'En inspección ..., ' / 'Durante ..., ' / 'Previo a ..., ': se quita si queda >= 25."""
    for _ in range(3):
        if not _INTRO.match(_plano(t)):
            break
        i = t.find(",")
        if i == -1:
            break
        resto = t[i + 1:].strip()
        if len(resto) < _MIN_RESTO:
            break
        t = resto
    return t

def _quitar_conectores_iniciales(t: str) -> str:
    m = _CONECTOR_INICIAL.match(_plano(t))
    if m and len(t) - m.end() >= 10:
        return t[m.end():]
    return t

def _reordenar_al(t: str) -> str:
    """'al solicitar X se verifica que Y' -> 'X Y' (sin artículo inicial; 'al' -> 'del')."""
    m = _AL_RE.match(_plano(t))
    if not m:
        return t
    obj = t[m.start("obj"):m.end("obj")].strip()
    hall = t[m.start("hall"):m.end("hall")].strip()
    if _DET_INICIAL.match(_plano(hall)):
        return hall  # el hallazgo trae su propio sujeto
    obj = re.sub(r"^(?:el|la|los|las|un|una|unos|unas)\s+", "", obj, flags=re.I)
    obj = re.sub(r"\bal\b", "del", obj)
    return f"{obj} {hall}"

def _quitar_muletillas(t: str) -> str:
    for _ in range(4):
        m = _MULETILLA.search(_plano(t))
        if not m:
            break
        prefijo = t[:m.start()].rstrip(" ,;")
        resto = t[m.end():].strip()
        if not resto:
            break
        if prefijo and _INTRO.match(_plano(prefijo)) and len(resto) >= _MIN_RESTO:
            t = resto  # el prefijo era solo contexto
        else:
            t = f"{prefijo} {resto}".strip()
    return t

def _quitar_administrativo(t: str) -> str:
    m = _ADMIN.match(_plano(t))
    if m and len(t) - m.end() >= 3:
        return t[m.end():]
    return t

def _limpiar_oracion(o: str) -> str:
    t = _espacios(o)
    t = re.sub(r"^\s*(?:[-•*·]+|\d+[.)])\s+", "", t)
    t = _quitar_parentesis(t)
    t = t.strip().rstrip(".;:!? ")
    t = _quitar_intro(t)
    t = _quitar_conectores_iniciales(t)
    t = _reordenar_al(t)
    t = _quitar_muletillas(t)
    t = _quitar_administrativo(t)
    t = _quitar_conectores_iniciales(t)
    return _espacios(t).strip(" ,;:-")

def _dividir_oraciones(t: str) -> list[str]:
    partes = re.split(r"(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ¿¡0-9])|\n+", t)
    return [p.strip() for p in partes if p and p.strip()]

# --------------------------------------------------------------------------
# Puntuación
# --------------------------------------------------------------------------

_HALLAZGO = [
    re.compile(p)
    for p in (
        r"\bsin\b", r"falta", r"ausencia", r"no cuenta", r"no corresponde",
        r"correspondia", r"incumpl", r"deficien", r"danad", r"expuest",
        r"sobre", r"fuera de",
    )
]
_CONTEXTO = ("se realiza", "se solicita", "cabe destacar", "se informa", "se gestiono")

def _puntuar(original: str, limpia: str, tema: list[str]) -> int:
    p_l, p_o = _plano(limpia), _plano(original)
    puntos = sum(1 for w in set(tema) if w in p_l)           # + por palabra de tema
    if any(h.search(p_l) for h in _HALLAZGO):                # + si trae hallazgo
        puntos += 2
    puntos -= 2 * sum(1 for c in _CONTEXTO if c in p_o)      # - si es contexto
    if len(limpia) < 12:
        puntos -= 3
    return puntos

# --------------------------------------------------------------------------
# Acortado
# --------------------------------------------------------------------------

# Palabras con las que una frase no puede terminar (comparadas sin tildes).
_FINALES = frozenset("""
de del con sin falta faltan ausencia carencia ya que por para a al en el la los las
un una unos unas lo y e o u ni se su sus como entre sobre hacia desde hasta segun ante
tras durante mediante cuando donde pero aunque si no mas muy tambien ademas porque pues
contra bajo cada embargo cuenta cuentan carece carecen tiene tienen debe deben es son
esta estan fue fueron era eran ha han hay ser estar haber
""".split())

_CORTE = re.compile(
    r"(?:(?<!\d),|,(?!\d)|;)"
    r"|\s+(?:ya\s+que|por\s+lo\s+que|sin\s+embargo|por\s+lo\s+tanto|dado\s+que"
    r"|debido\s+a\s+que|lo\s+que)\b"
)

def _limpiar_final(s: str) -> str:
    """Quita puntuación final, paréntesis abiertos y conectores colgantes."""
    s = s.strip()
    while s:
        anterior = s
        s = s.rstrip(" ,;:.-–—")
        if s.count("(") > s.count(")"):
            s = s[: s.rfind("(")].rstrip()
        partes = s.rsplit(None, 1)
        if partes:
            ult = partes[-1]
            if _plano(ult) in _FINALES and not (len(ult) == 1 and ult.isupper()):
                s = partes[0] if len(partes) == 2 else ""
        if s == anterior:
            break
    return s.strip()

def _acortar(texto: str, max_chars: int) -> str:
    texto = texto.strip()
    if max_chars < 1 or not texto:
        return ""
    if len(texto) <= max_chars:
        return _limpiar_final(texto) or texto
    plano = _plano(texto)
    minimo = min(30, max(10, max_chars // 2))
    corte = 0
    for m in _CORTE.finditer(plano):          # límite de cláusula
        if m.start() > max_chars:
            break
        if len(texto[: m.start()].strip()) >= minimo:
            corte = m.start()
    if corte:
        cand = _limpiar_final(texto[:corte])
        if cand:
            return cand
    trozo = texto[:max_chars]                 # límite de palabra completa
    if texto[max_chars] not in " \t,;:" and " " in trozo:
        trozo = trozo.rsplit(" ", 1)[0]
    return _limpiar_final(trozo) or trozo.strip()

def _minuscula_inicial(t: str) -> str:
    if not t:
        return t
    if re.match(r"[A-ZÁÉÍÓÚÑ]{2,}(?![a-záéíóúñ])", t):  # siglas (EPP, RIT...) se respetan
        return t
    return t[0].lower() + t[1:]

# --------------------------------------------------------------------------
# Título
# --------------------------------------------------------------------------

_TOK = r"(?!EPP\b)(?=(?:[a-z]*[A-Z]){2})[A-Za-z]{2,5}"
_PREFIJO_TITULO = re.compile(rf"^\s*(?:{_TOK}(?:\s*[-/]\s*{_TOK})*\s*[-:–]\s*)+")

def _quitar_prefijo_titulo(titulo: str) -> str:
    t = _PREFIJO_TITULO.sub("", titulo, count=1).strip()
    return t or titulo.strip()

# --------------------------------------------------------------------------
# API pública
# --------------------------------------------------------------------------

def resumir(titulo, descripcion, palabras_tema, max_chars: int = 120) -> str:
    """Una frase concreta (hecho observado) de hasta `max_chars` caracteres."""
    if not max_chars or max_chars < 1:
        return ""
    titulo = unicodedata.normalize("NFC", titulo or "")
    desc = unicodedata.normalize("NFC", descripcion or "")
    tema = [
        p for p in (_plano(x).strip() for x in (palabras_tema or []) if isinstance(x, str)) if p
    ]

    candidatos = []
    if desc.strip():
        desc = _quitar_nombres(desc)
        for orden, o in enumerate(_dividir_oraciones(desc)):
            limpia = _limpiar_oracion(o)
            if re.search(r"\w", limpia):
                candidatos.append((_puntuar(o, limpia, tema), orden, limpia))

    if candidatos:
        mejor = max(candidatos, key=lambda c: (c[0], -c[1]))[2]
    else:
        base = _quitar_prefijo_titulo(titulo)
        mejor = _espacios(_quitar_nombres(base)).strip(" ,;:-.")

    res = _acortar(mejor, max_chars)
    return _minuscula_inicial(res.rstrip(" ."))

# --------------------------------------------------------------------------
# Autopruebas (solo textos sintéticos inventados)
# --------------------------------------------------------------------------

if __name__ == "__main__":
    _CONECTORES = (
        "de", "del", "con", "sin", "falta", "ya que", "que", "por", "para", "en",
        "a", "al", "el", "la", "los", "las", "un", "una", "y", "o", "e", "se",
    )
    _NOMBRES_PROHIBIDOS = ("Juan", "Pérez", "Soto", "Pedro", "González", "Alvarado")

    def _termina_en_conector(s: str) -> bool:
        pl = _plano(s).rstrip(" ,;:.")
        return any(pl == c or pl.endswith(" " + c) for c in _CONECTORES)

    def _ok(r: str, mc: int) -> None:
        assert r, "resultado vacío"
        assert len(r) <= mc, (len(r), mc, r)
        assert not _termina_en_conector(r), r
        assert r[0] == r[0].lower(), r
        assert not r.endswith("."), r
        assert not any(n in r for n in _NOMBRES_PROHIBIDOS), r

    D1 = (
        "En inspección realizada a trabajos de izaje con camión pluma de la obra X, "
        "al solicitar el plan de izaje al rigger se verifica que correspondía a otra "
        "carga y no a la izada."
    )
    T1 = ["izaje", "plan", "carga"]
    D2 = (
        "Se observa que el andamio del sector norte no cuenta con barandas laterales, "
        "ya que el montaje fue entregado sin revisión previa por el área de prevención "
        "de riesgos de la obra."
    )
    T2 = ["andamio", "baranda"]
    D3 = (
        "Se realiza inspección general del área de acopio. El extintor de la bodega "
        "se encuentra sin carga vigente y con la etiqueta vencida."
    )
    T3 = ["extintor", "carga"]

    # 1) ejemplo del enunciado: reordena "al solicitar X se verifica que Y"
    r1 = resumir("HBP-Izaje", D1, T1)
    assert r1 == "plan de izaje del rigger correspondía a otra carga y no a la izada", r1
    _ok(r1, 120)

    # 2) muletilla + corte en límite de cláusula (coma / 'ya que')
    r2 = resumir("", D2, T2)
    assert r2 == "el andamio del sector norte no cuenta con barandas laterales", r2
    _ok(r2, 120)

    # 3) puntuación: se descarta la oración de contexto ('se realiza')
    r3 = resumir("", D3, T3)
    assert r3 == "el extintor de la bodega se encuentra sin carga vigente y con la etiqueta vencida", r3
    _ok(r3, 120)

    # 4-7) descripción vacía: se usa el título sin prefijos
    assert resumir("HBP-Falta de señalización en acceso norte", "", []) == \
        "falta de señalización en acceso norte"
    assert resumir("CBP - Extintor sin carga en bodega", "", []) == \
        "extintor sin carga en bodega"
    assert resumir("ASP-HAP: Orden y limpieza deficiente en taller", "  ", []) == \
        "orden y limpieza deficiente en taller"
    assert resumir("CdP-Cables expuestos en tablero eléctrico", None, None) == \
        "cables expuestos en tablero eléctrico"

    # 8) nombres de personas: se conserva el cargo
    r8 = resumir("", "El supervisor Juan Pérez Soto indicó que el operario trabaja "
                     "sin arnés de seguridad en altura.", ["arnes"])
    assert r8 == "el supervisor indicó que el operario trabaja sin arnés de seguridad en altura", r8
    _ok(r8, 120)

    # 9) nombre sin cargo -> 'trabajador'
    r9 = resumir("", "Pedro González no usa casco en zona de carga.", ["casco"])
    assert r9 == "trabajador no usa casco en zona de carga", r9
    _ok(r9, 120)

    # 10) tratamiento + apellido
    r10 = resumir("", "El Sr. Alvarado no entregó el permiso de trabajo en caliente.", [])
    assert r10 == "el trabajador no entregó el permiso de trabajo en caliente", r10
    _ok(r10, 120)

    # 11) cláusula introductoria 'Durante ...'
    r11 = resumir("", "Durante la verificación de equipos de izaje en patio, el estrobo "
                      "de acero presenta cables cortados y deformación visible.", ["estrobo"])
    assert r11 == "el estrobo de acero presenta cables cortados y deformación visible", r11

    # 12) cláusula introductoria 'Previo a ...'
    r12 = resumir("", "Previo a la partida de la faena en terreno, el operador no cuenta "
                      "con licencia vigente para la máquina.", ["licencia"])
    assert r12 == "el operador no cuenta con licencia vigente para la máquina", r12

    # 13) 'En recorrido por ...' + muletilla 'se observa que'
    r13 = resumir("", "En recorrido por el sector de bodega, se observa que los cilindros "
                      "de gas están sin cadena de sujeción.", ["cilindro"])
    assert r13 == "los cilindros de gas están sin cadena de sujeción", r13
    assert "se observa" not in r13 and "recorrido" not in r13

    # 14) si lo que sigue a la intro tiene < 25 caracteres, NO se quita
    r14 = resumir("", "En inspección, falta señal.", [])
    assert r14 == "en inspección, falta señal", r14

    # 15) 'se solicita ...' (administrativo) pierde frente a una oración con hallazgo
    r15 = resumir("", "Se solicita al contratista enviar el informe. La escalera telescópica "
                      "está dañada y sin zapatas antideslizantes.", ["escalera", "zapata"])
    assert r15 == "la escalera telescópica está dañada y sin zapatas antideslizantes", r15

    # 16) entradas vacías / max_chars inválido
    assert resumir("", "", []) == ""
    assert resumir("x", "y", [], 0) == ""

    # 17) corte con max_chars por defecto: palabra completa, sin conector final
    W = ("La grúa horquilla opera con las horquillas elevadas sobre el nivel de los "
         "peatones en el pasillo central de la bodega norte durante todo el turno de la "
         "mañana sin señalero")
    r17 = resumir("", W, ["horquilla", "peatones"])
    _ok(r17, 120)
    assert r17.endswith("bodega"), r17

    # 18) longitud, conector final, minúscula inicial y sin nombres, con varios max_chars
    for mc in (20, 30, 40, 45, 60, 90, 120):
        _ok(resumir("", D1, T1, mc), mc)
    for mc in (30, 50, 80, 120):
        _ok(resumir("", D2, T2, mc), mc)
    for mc in (30, 50, 70, 120):
        _ok(resumir("", D3, T3, mc), mc)

    # 19) determinismo
    for d, t in ((D1, T1), (D2, T2), (D3, T3)):
        assert len({resumir("t", d, t, 80) for _ in range(20)}) == 1

    print("OK: todos los asserts pasaron")
