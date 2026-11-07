# -*- coding: utf-8 -*-
"""
catalogo_extra.py: temas adicionales del catálogo (solo datos, sin lógica).

Mismas convenciones que catalogo_data.CATALOGO:
  - palabras / excluir: regex en minúscula y sin tildes, para aplicar sobre texto
    ya normalizado (la ñ pasa a n). `excluir` es siempre una regex con contenido real.
  - extras: [regex, cláusula] en minúscula y sin punto final.
  - bloque: "am" (inicio de jornada) o "mediodia".
"""

CATALOGO_EXTRA = [
    # ------------------------------------------------------------------ 1
    {
        "id": "segregacion_acceso",
        "tema": "Segregación y control de acceso",
        "nombre_corto": "segregación y control de acceso deficientes",
        "palabras": [
            r"\bsegregacion\b",
            r"\bsin segregar\b|\bsin segregacion\b",
            r"\bsegregad(o|a|os|as)\b",
            r"barrera(s)? (dura|duras|fisica|fisicas|rigida|rigidas)",
            r"\bnew jersey\b",
            r"\bcinta(s)? (de )?(peligro|seguridad|demarcacion|barricada|advertencia)",
            r"\bbarricada(s)?\b",
            r"\bdelimitacion\b|\bdelimitad(o|a|os|as)\b|\bsin delimitar\b",
            r"\bperimetro\b",
            r"\bingreso(s)? (no autorizado|sin autorizacion)",
            r"\bacceso(s)? (restringido|no autorizado|libre|sin control)",
            r"\blibre transito\b",
            r"\barea(s)? (restringida|sin segregar|sin delimitar|abierta)",
            r"\bpersonas? ajena(s)?\b",
        ],
        "excluir": r"segregacion de residuos|barrera(s)? (sanitaria|aduanera|idiomatica)|cinta (transportadora|metrica|adhesiva)",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que las áreas de trabajo de riesgo estén segregadas con barreras duras "
            "o cinta de peligro, señalizadas y con acceso restringido a personal autorizado; "
            "el supervisor lo controla al habilitar el área y en cada cambio de condición."
        ),
        "extras": [
            [r"cinta", "reemplazando la cinta por barreras duras donde hay equipos en movimiento"],
            [r"ingreso|acceso", "con un control de ingreso en cada acceso al área"],
            [r"libre transito|equipo", "separando los flujos de personas y equipos"],
        ],
        "criterio": "Área de riesgo sin barrera ni cinta, con ingreso de personal no autorizado o con personas y equipos en el mismo espacio.",
        "herramienta": (
            "CdP Segregación + Desafío en Terreno + ASP: recorrer perímetros y accesos de áreas "
            "de riesgo, verificando barreras, cinta y control de ingreso, dos veces por turno."
        ),
        "evidencia": "Registro de segregación de áreas / ASP del frente / registro fotográfico",
        "control_corto": "Áreas segregadas y acceso controlado",
        "brecha_cdp": "segregación y control de acceso a áreas",
        "pendiente": "segregación y control de acceso",
    },
    # ------------------------------------------------------------------ 2
    {
        "id": "senaletica_iluminacion",
        "tema": "Señalética e iluminación de áreas",
        "nombre_corto": "señalética e iluminación deficientes",
        "palabras": [
            r"\bsenaletica\b",
            r"\bsenalizacion\b",
            r"\bsin senalizar\b|\bsin senalizacion\b",
            r"\bsenal(es)? (faltante|ilegible|deteriorada|caida|danada|fuera de posicion|mal ubicada)",
            r"\bsin senal(es)?\b",
            r"\bletrero(s)?\b",
            r"\bdisco(s)? pare\b|\bsenal(es)? pare\b",
            r"\bpictograma(s)?\b",
            r"\biluminacion\b",
            r"\bluminaria(s)?\b",
            r"\bpoca luz\b|\bfalta de luz\b|\bsin luz\b",
            r"\bfoco(s)? (quemado|apagado|danado)",
            r"\bpoca visibilidad\b|\bbaja visibilidad\b",
            r"\bcono(s)? de (transito|senalizacion|seguridad)",
        ],
        "excluir": r"senal (celular|de (celular|telefono|radio|internet|gps|wifi))|iluminacion (navidena|decorativa)",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que la señalética de cada área esté completa, legible y en su posición, "
            "con discos pare y letreros de advertencia visibles, y que la iluminación permita "
            "ver rutas y puntos de trabajo; el supervisor lo revisa al inicio del turno y antes "
            "de operar de noche."
        ),
        "extras": [
            [r"disco|pare", "con discos pare instalados en las intersecciones"],
            [r"ilegible|deteriorad", "reponiendo las señales deterioradas o ilegibles"],
            [r"iluminacion|luminaria|luz", "con luminarias operativas en rutas y frentes nocturnos"],
            [r"fuera de posicion|mal ubicada|caida", "reinstalando las señales caídas o fuera de posición"],
        ],
        "criterio": "Señal faltante, ilegible o fuera de posición en zona de riesgo, o iluminación insuficiente para ver la ruta o la tarea.",
        "herramienta": (
            "ASP + Desafío en Terreno + PASSO: recorrer rutas y frentes verificando señales, discos "
            "pare y luminarias, corrigiendo en el momento, al inicio del turno y al anochecer."
        ),
        "evidencia": "Registro de revisión de señalética / Registro de niveles de iluminación / registro fotográfico",
        "control_corto": "Señalética e iluminación de áreas",
        "brecha_cdp": "señalética e iluminación de áreas",
        "pendiente": "señalética e iluminación de áreas",
    },
    # ------------------------------------------------------------------ 3
    {
        "id": "generadores_equipos_menores",
        "tema": "Generadores y equipos menores",
        "nombre_corto": "generadores y equipos menores sin control",
        "palabras": [
            r"\bgenerador(es)?\b",
            r"\bgrupo(s)? electrogeno(s)?\b",
            r"\btorre(s)? de (iluminacion|luz)\b",
            r"\bcompresor(es)?\b",
            r"\bmotobomba(s)?\b",
            r"\bequipo(s)? menor(es)?\b",
            r"\bsin horometro\b",
            r"\bsin responsable\b|\bresponsable del equipo\b",
            r"\bequipo(s)? sin (documentacion|mantencion|identificacion)",
        ],
        "excluir": r"generador(es)? de (residuos|empleo|ideas)|generacion de residuos",
        "bloque": "am",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que generadores, torres de iluminación y compresores tengan horómetro "
            "operativo, documentación al día, responsable designado y mantención vigente; el "
            "responsable del equipo lo registra y el supervisor lo valida al inicio del turno."
        ),
        "extras": [
            [r"horometro", "registrando la lectura del horómetro al inicio del turno"],
            [r"tierra|electric", "con puesta a tierra y protecciones eléctricas operativas"],
            [r"combustible|derrame|fuga", "con bandeja antiderrame bajo el equipo"],
            [r"compresor", "con mangueras y acoples del compresor asegurados"],
        ],
        "criterio": "Generador, torre o compresor sin horómetro operativo, sin documentación ni responsable asignado, o con fuga visible.",
        "herramienta": (
            "Inspección de Equipos + ASP + PASSO: revisar horómetro, documentación y responsable "
            "de cada generador, torre y compresor del área, al inicio del turno."
        ),
        "evidencia": "Registro de inspección de equipos menores / Documentación y responsable asignado / registro fotográfico",
        "control_corto": "Generadores y equipos menores al día",
        "brecha_cdp": "control de generadores y equipos menores",
        "pendiente": "generadores y equipos menores",
    },
    # ------------------------------------------------------------------ 4
    {
        "id": "inmovilizacion_estacionamiento",
        "tema": "Estacionamiento e inmovilización de equipos",
        "nombre_corto": "estacionamiento e inmovilización de equipos deficientes",
        "palabras": [
            r"\bcunas?\b",
            r"\bsin cunas?\b",
            r"\bfreno(s)? de (parqueo|estacionamiento|mano)",
            r"\bestacionad(o|a|os|as)\b",
            r"\bestacionamiento(s)?\b",
            r"\bestacionar\b|\bestaciono\b|\bestacionando\b",
            r"\bparqueo\b|\bparquead(o|a|os|as)\b|\bparquear\b",
            r"\binmovilizacion\b|\binmovilizad(o|a|os|as)\b|\binmovilizar\b",
            r"\bmal estacionad(o|a|os|as)\b",
            r"zona(s)? no autorizada(s)? (para|de) (estacionar|estacionamiento)",
            r"\bsin freno de (parqueo|estacionamiento|mano)",
        ],
        "excluir": r"estacionamiento (de visitas|de oficinas|administrativo|subterraneo)|cuna (de|del) (bebe|ninos?)",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que todo vehículo o equipo detenido tenga freno de parqueo aplicado, "
            "cuñas en las ruedas y estacionamiento solo en zonas autorizadas y niveladas; el "
            "operador lo aplica al detenerse y el supervisor lo controla al inicio del turno."
        ),
        "extras": [
            [r"cuna", "con cuñas en ambos lados de la rueda en pendiente"],
            [r"zona(s)? no autorizada|mal estacionad", "estacionando solo en las zonas demarcadas"],
            [r"pendiente|talud|rampa", "evitando estacionar en pendientes o rampas"],
        ],
        "criterio": "Vehículo o equipo detenido sin freno de parqueo ni cuñas, o estacionado en zona no autorizada o en pendiente.",
        "herramienta": (
            "EdC Vehículo Liviano + Desafío en Terreno + PASSO: observar el estacionamiento de "
            "vehículos y equipos en patios y frentes, verificando cuñas y freno de parqueo, dos veces por turno."
        ),
        "evidencia": "Registro de recorrido de estacionamientos / Check list pre-uso firmado / registro fotográfico",
        "control_corto": "Cuñas y freno de parqueo aplicados",
        "brecha_cdp": "estacionamiento e inmovilización de equipos",
        "pendiente": "estacionamiento e inmovilización de equipos",
    },
    # ------------------------------------------------------------------ 5
    {
        "id": "instalaciones_higiene",
        "tema": "Instalaciones sanitarias e hidratación",
        "nombre_corto": "instalaciones sanitarias e hidratación deficientes",
        "palabras": [
            r"\bservicio(s)? higienico(s)?\b",
            r"\bbano(s)?\b",
            r"\bbano(s)? quimico(s)?\b",
            r"\bsala(s)? de hidratacion\b",
            r"\bhidratacion\b",
            r"\bagua potable\b",
            r"\bdispensador(es)? de agua\b",
            r"\bcomedor(es)?\b",
            r"\bcolacion(es)?\b",
            r"\bingesta de alimentos\b",
            r"\b(comiendo|ingiriendo|consumiendo|ingerir|consumir) (alimentos|colacion|comida)",
            r"\blavamanos\b",
            r"\bpapel higienico\b",
            r"\bsala(s)? de descanso\b",
        ],
        "excluir": r"residuos sanitarios|barrera sanitaria|aduana sanitaria|bano (de aceite|maria)",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que los servicios higiénicos, la sala de hidratación y el comedor estén "
            "operativos, limpios y con agua potable disponible, y que nadie ingiera alimentos "
            "fuera de las zonas definidas; el supervisor lo revisa a mediodía y tras cada "
            "reposición de insumos."
        ),
        "extras": [
            [r"agua|hidratacion", "con agua potable fresca disponible en cada frente"],
            [r"bano|higienico", "con limpieza de baños registrada en cada turno"],
            [r"alimento|colacion|comiendo|ingiriendo", "señalizando las zonas habilitadas para ingerir alimentos"],
        ],
        "criterio": "Servicio higiénico fuera de uso, sin agua potable disponible o personal ingiriendo alimentos fuera de las zonas definidas.",
        "herramienta": (
            "ASP + Desafío en Terreno + PASSO: revisar servicios higiénicos, sala de hidratación y "
            "comedor, y observar dónde se ingieren alimentos."
        ),
        "evidencia": "Registro de limpieza y revisión de instalaciones / Registro de reposición de agua e insumos / registro fotográfico",
        "control_corto": "Servicios higiénicos e hidratación",
        "brecha_cdp": "instalaciones sanitarias e hidratación",
        "pendiente": "instalaciones sanitarias e hidratación",
    },
    # ------------------------------------------------------------------ 6
    {
        "id": "componentes_mangueras",
        "tema": "Mangueras y componentes hidráulicos",
        "nombre_corto": "mangueras y componentes hidráulicos en mal estado",
        "palabras": [
            r"\bmanguera(s)?\b",
            r"\bflexible(s)? (hidraulic|de alta|roto|danado|reventado|con fuga)",
            r"\bhidraulic(o|a|os|as)\b",
            r"\bfuga(s)?\b",
            r"\bgoteo(s)?\b|\bgoteando\b",
            r"\bcomponente(s)?\b",
            r"\bfalla(s)? (mecanica|hidraulica|electrica|del equipo|de componente)",
            r"\bequipo(s)? (con|presenta|presento) falla",
            r"\bretroexcavadora(s)?\b",
            r"\breventon(es)?\b|\brevento\b|\breventad(o|a|os|as)\b",
            r"\bacople(s)?\b",
            r"\babrazadera(s)?\b",
            r"\bcilindro(s)? hidraulico(s)?\b",
        ],
        "excluir": r"manguera(s)? de (incendio|red humeda)|fuga de (informacion|datos|personal)|falla de (conexion|internet|comunicacion)",
        "bloque": "am",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que mangueras, flexibles y componentes hidráulicos del equipo estén sin "
            "fugas, cortes ni abrazaderas sueltas, y que toda falla se reporte y se repare antes "
            "de operar; el operador lo revisa en el pre-uso y el supervisor lo valida al inicio del turno."
        ),
        "extras": [
            [r"fuga|goteo", "retirando de servicio todo equipo con fuga activa"],
            [r"reventon|revento|reventad", "con protectores en flexibles de alta presión"],
            [r"abrazadera|acople", "con acoples y abrazaderas ajustados"],
            [r"retroexcavadora", "con revisión previa de la retroexcavadora antes de operar"],
        ],
        "criterio": "Manguera o flexible con corte, fuga o acople suelto, o equipo con falla de componente que sigue en operación.",
        "herramienta": (
            "Inspección de Equipos + ASP + PASSO: revisar mangueras, flexibles y acoples de los "
            "equipos en uso, en el acceso a faena, al inicio del turno."
        ),
        "evidencia": "Check list pre-uso firmado / Registro de inspección de equipos / Orden de reparación / registro fotográfico",
        "control_corto": "Mangueras y componentes hidráulicos",
        "brecha_cdp": "estado de mangueras y componentes hidráulicos",
        "pendiente": "mangueras y componentes hidráulicos",
    },
    # ------------------------------------------------------------------ 7
    {
        "id": "pretiles_bordes",
        "tema": "Pretiles y bordes de caminos",
        "nombre_corto": "pretiles y bordes de camino deficientes",
        "palabras": [
            r"\bpretil(es)?\b",
            r"\bberma(s)?\b",
            r"\bborde(s)? de (camino|ruta|via|botadero|stock|rampa|descarga|talud|banco)",
            r"\b(proximo|cerca|cercano|junto|pegado) (a|al|del) borde\b",
            r"\bal borde\b",
            r"\bbotadero(s)?\b",
            r"\bstock ?pile(s)?\b",
            r"\btope(s)? de (descarga|retroceso|seguridad)",
            r"\bsin tope(s)?\b",
            r"\btolva(s)?\b",
            r"\bcamion(es)? tolva\b",
            r"\bdespenamiento\b|\bdespenar\b",
            r"\bcaida (de|del) (camion|equipo|vehiculo) (por|al|desde) (el )?(borde|talud|botadero)",
        ],
        "excluir": r"borde (de losa|sin proteccion|expuesto|de (la )?(excavacion|zanja))|bordes? de (mesa|cinta)",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que caminos y botaderos mantengan pretil o berma continuos, de altura "
            "acorde al equipo, y topes de descarga en buen estado, sin camiones tolva próximos "
            "al borde; el supervisor lo inspecciona antes de iniciar la descarga y al inicio del turno."
        ),
        "extras": [
            [r"berma|pretil", "reponiendo de inmediato los tramos de pretil o berma faltantes"],
            [r"tope", "con topes de descarga instalados a distancia segura del borde"],
            [r"tolva|camion", "indicando al operador la distancia mínima al borde durante la descarga"],
            [r"botadero|stock", "con señalero en el punto de descarga del botadero"],
        ],
        "criterio": "Tramo sin pretil ni berma, tope de descarga ausente o camión tolva detenido o descargando próximo al borde.",
        "herramienta": (
            "CdP Tránsito + Inspección de Equipos + ASP: inspeccionar pretiles, bermas y topes en "
            "botaderos y caminos, y observar la descarga de camiones, antes de descargar."
        ),
        "evidencia": "Registro de inspección de pretiles y bermas / Registro de recorrido de caminos / registro fotográfico",
        "control_corto": "Pretiles, bermas y topes de descarga",
        "brecha_cdp": "pretiles, bermas y bordes de camino",
        "pendiente": "pretiles y bordes de caminos",
    },
]
