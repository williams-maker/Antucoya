# -*- coding: utf-8 -*-
"""
catalogo_data.py: solo datos (sin lógica).

CATALOGO: lista de temas. Convenciones:
  - palabras / excluir: regex en minúscula y sin tildes, para aplicar sobre texto
    ya normalizado. Se usa [nñ] donde podría aparecer ñ, por si la normalización
    la conserva. `excluir` es siempre una regex con contenido real (nunca vacía).
  - extras: [regex, cláusula]. La cláusula va en minúscula, sin punto final, y se
    anexa a la instrucción como ", " + cláusula antes del punto final.
  - bloque: "am" (inicio de jornada) o "mediodia".

REVAL: reglas de revalorización (solo sugerencia, nunca se aplica sola).
"""

CATALOGO = [
    # ------------------------------------------------------------------ 1
    {
        "id": "izaje_plan",
        "tema": "Plan de izaje",
        "nombre_corto": "planes de izaje incompletos o ausentes",
        "palabras": [
            r"plan(es)? de (izaje|levante)",
            r"\bizaje(s)?\b",
            r"\bizar\b",
            r"\bizando\b",
            r"maniobra(s)? de (izaje|levante)",
            r"\bgrua(s)?\b",
            r"camion(es)? (pluma|grua)",
            r"\bpluma\b",
            r"\blevante(s)?\b",
            r"capacidad de (la )?grua",
            r"tabla(s)? de carga",
        ],
        "excluir": r"grua horquilla|montacarga",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que toda maniobra de izaje cuente con plan específico para la carga "
            "y el equipo, con radio de trabajo, peso y accesorios definidos; el rigger lo "
            "porta en terreno y el supervisor lo valida antes del levante."
        ),
        "extras": [
            [r"gener(ico|ica)|incomplet", "evitando planes genéricos o incompletos"],
            [r"\bviento", "considerando el límite de viento del fabricante"],
            [r"izaje de (personas|personal)|canasto", "prohibiendo el izaje de personas fuera de canasto autorizado"],
        ],
        "criterio": "Maniobra sin plan de izaje específico, o con carga, equipo o radio distintos a los del plan.",
        "herramienta": (
            "CdP Izaje + PTS + ASP: validar en terreno el plan de cada maniobra antes del "
            "levante, junto al rigger y al supervisor, en todas las maniobras del turno."
        ),
        "evidencia": "PTS con plan de izaje específico / Check list de equipo y accesorios / registro fotográfico",
        "control_corto": "Plan de izaje específico en terreno",
        "brecha_cdp": "plan de izaje específico",
        "pendiente": "plan de izaje",
    },
    # ------------------------------------------------------------------ 2
    {
        "id": "izaje_segregacion",
        "tema": "Segregación y comunicación en izaje",
        "nombre_corto": "segregación y comunicación deficientes en izaje",
        "palabras": [
            r"area(s)? de (influencia|maniobra|izaje)",
            r"segregacion (de(l)? )?(area|carga|izaje|maniobra|zona)",
            r"\bradio(s)?\b",
            r"comunicacion (por )?radial",
            r"\bviento(s)?\b",
            r"\brafaga(s)?\b",
            r"carga suspendida",
            r"(bajo|debajo de) (la )?carga",
            r"\brigger(s)?\b",
            r"radio de giro",
            r"personas? (en|dentro de) (el )?(radio|area) de (giro|maniobra)",
        ],
        "excluir": r"segregacion de residuos",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que el área de influencia de la carga esté segregada y señalizada, "
            "y que rigger y operador mantengan comunicación por radio o señas definidas; "
            "el supervisor lo confirma antes de izar y controla el viento durante la maniobra."
        ),
        "extras": [
            [r"\bradio(s)?\b|comunicacion", "con canal de radio exclusivo para la maniobra"],
            [r"carga suspendida|bajo (la )?carga", "prohibiendo permanecer bajo la carga suspendida"],
            [r"se[nñ]al", "con señales manuales acordadas entre rigger y operador"],
        ],
        "criterio": "Persona dentro del área de influencia de la carga, sin comunicación operador-rigger o con viento sobre el límite.",
        "herramienta": (
            "CdP Izaje + Desafío en Terreno + ASP: recorrer cada maniobra antes y durante el "
            "levante, confirmando segregación, radio y viento con operador y rigger, dos veces por turno."
        ),
        "evidencia": "ASP de la maniobra / Registro de segregación y señalización / registro fotográfico",
        "control_corto": "Área de izaje segregada y comunicada",
        "brecha_cdp": "segregación y comunicación en izaje",
        "pendiente": "segregación y comunicación en izaje",
    },
    # ------------------------------------------------------------------ 3
    {
        "id": "izaje_certificacion",
        "tema": "Certificación de elementos de izaje",
        "nombre_corto": "elementos de izaje sin certificación o con daño",
        "palabras": [
            r"grillete(s)?",
            r"eslinga(s)?",
            r"\btecle(s)?\b",
            r"\bplaca(s)?\b",
            r"certificad(o|a|os|as)\b",
            r"certificacion(es)?",
            r"linea(s)? de fuego",
            r"estrobo(s)?",
            r"aparejo(s)?",
            r"cadena(s)? de izaje",
            r"\bpestillo(s)? de seguridad",
            r"\bgancho(s)? sin (seguro|traba|pestillo)",
            r"cable(s)? de acero",
            r"\bcancamo(s)?\b",
        ],
        "excluir": r"certificado (medico|de (salud|curso|capacitacion))|licencia medica",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que grilletes, eslingas, tecles y demás elementos de izaje tengan "
            "certificado y placa vigentes y sin daño visible, y que nadie se ubique bajo la "
            "carga; el rigger lo revisa y el supervisor lo valida antes de izar."
        ),
        "extras": [
            [r"grillete", "con grilletes de capacidad acorde y pasador asegurado"],
            [r"eslinga|estrobo", "descartando eslingas con cortes, nudos o desgaste"],
            [r"tecle", "con tecles inspeccionados y placa de capacidad legible"],
            [r"linea(s)? de fuego", "manteniendo al personal fuera de la línea de fuego de la carga"],
        ],
        "criterio": "Elemento de izaje sin certificado ni placa legible, con daño visible, o personas bajo la carga suspendida.",
        "herramienta": (
            "CdP Izaje + Inspección de Equipos + PASSO: revisar placa, certificado y estado de "
            "grilletes, eslingas y tecles en el punto de izaje, antes de cada maniobra."
        ),
        "evidencia": "Registro de inspección de elementos de izaje / Certificados y placas vigentes / registro fotográfico",
        "control_corto": "Certificación de elementos de izaje",
        "brecha_cdp": "certificación de elementos de izaje",
        "pendiente": "certificación de elementos de izaje",
    },
    # ------------------------------------------------------------------ 4
    {
        "id": "preuso_equipos",
        "tema": "Inspección pre-uso de equipos",
        "nombre_corto": "inspección pre-uso de equipos deficiente",
        "palabras": [
            r"check ?list",
            r"pre[- ]?uso",
            r"neumatico(s)?",
            r"horometro",
            r"letrero (de )?supervisor",
            r"\bletrero(s)?\b",
            r"plan de transito",
            r"inspeccion (diaria|previa) (del|de) (equipo|vehiculo|camioneta)",
            r"libro de (vida|novedades)",
            r"\bcamioneta(s)?\b",
            r"\bfreno(s)?\b",
        ],
        "excluir": r"check ?list de (andamio|izaje|plataforma)",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que vehículos y equipos tengan check list pre-uso completo, "
            "neumáticos en buen estado, horómetro operativo, letrero de supervisor visible y "
            "plan de tránsito; el operador lo ejecuta y el supervisor lo valida al inicio del turno."
        ),
        "extras": [
            [r"horometro", "registrando la lectura del horómetro en el check list"],
            [r"neumatico", "con desgaste y presión de neumáticos revisados"],
            [r"alarma|bocina|\bluces\b", "con alarmas, bocina y luces probadas"],
        ],
        "criterio": "Check list pre-uso incompleto o con falla crítica (frenos, neumáticos, alarmas, luces); equipo sin letrero de supervisor.",
        "herramienta": (
            "EdC Vehículo Liviano + Inspección de Equipos + Acompañamiento ASP: revisar el check "
            "list y el estado de vehículos y equipos en el acceso a faena, al inicio del turno."
        ),
        "evidencia": "Check list pre-uso firmado / Registro de inspección de equipos / registro fotográfico",
        "control_corto": "Check list pre-uso y letrero de supervisor",
        "brecha_cdp": "check list pre-uso de equipos",
        "pendiente": "inspección pre-uso de equipos",
    },
    # ------------------------------------------------------------------ 5
    {
        "id": "transito_peaton_equipo",
        "tema": "Tránsito y segregación peatón-equipo",
        "nombre_corto": "segregación peatón-equipo deficiente",
        "palabras": [
            r"transito (vehicular|peatonal|interno)",
            r"\bpeaton(es)?\b",
            r"segregacion (de )?(peaton|peatones|vehiculo|vehiculos|equipo|equipos)",
            r"\bvelocidad(es)?\b",
            r"limite de velocidad",
            r"\bretroceso\b",
            r"\bretrocede\b|\bretrocediendo\b",
            r"aculat(ar|ado|amiento|ando)",
            r"\bsenaler(o|os|a|as)\b",
            r"punto(s)? ciego(s)?",
            r"interaccion (hombre|persona)[- ]?(maquina|equipo)",
            r"cruce(s)? peatonal(es)?",
            r"camino(s)? de acarreo",
            r"\batropell",
        ],
        "excluir": r"accidente de trayecto",
        "bloque": "am",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que las rutas de tránsito estén segregadas entre peatones y equipos, "
            "con velocidad máxima respetada, alarma de retroceso operativa y señalero en "
            "maniobras de poca visibilidad; el supervisor lo controla al inicio del turno y en "
            "cada cambio de frente."
        ),
        "extras": [
            [r"aculat", "con tope de aculatamiento en el punto de descarga"],
            [r"velocidad", "con límite de velocidad señalizado en cada tramo"],
            [r"punto(s)? ciego|retroceso", "con puntos ciegos del equipo informados a los peatones"],
        ],
        "criterio": "Peatones y equipos en la misma ruta sin segregación, o equipo en retroceso sin señalero ni alarma operativa.",
        "herramienta": (
            "CdP Tránsito + Desafío en Terreno + PASSO: observar el cruce de peatones y equipos "
            "en rutas y frentes, corrigiendo en el momento, en dos recorridos por turno."
        ),
        "evidencia": "Registro de recorrido de rutas / ASP del frente / registro fotográfico",
        "control_corto": "Segregación peatón-equipo",
        "brecha_cdp": "segregación peatón-equipo",
        "pendiente": "segregación peatón-equipo",
    },
    # ------------------------------------------------------------------ 6
    {
        "id": "andamios_plataformas",
        "tema": "Liberación de andamios y plataformas",
        "nombre_corto": "andamios y plataformas sin liberación",
        "palabras": [
            r"\bandamio(s)?\b",
            r"\bplataforma(s)? de trabajo",
            r"liberacion de (andamio|plataforma)",
            r"\brodapie(s)?\b",
            r"\bbarand(a|as|illa|illas)\b",
            r"\bpeldano(s)?\b",
            r"tarjeta(s)? (verde|roja|amarilla)",
            r"tarjeta(s)? de (andamio|liberacion)",
            r"\bplataforma(s)? (elevadora|de elevacion|tijera)",
            r"\btablon(es)?\b",
            r"\bmodulo(s)? de andamio",
        ],
        "excluir": r"plataforma (digital|web|informatica)",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que todo andamio o plataforma de trabajo tenga tarjeta de liberación "
            "vigente, baranda, rodapié, peldaños completos y acceso seguro; el inspector "
            "habilitado lo libera y el supervisor lo confirma antes de habilitar el uso."
        ),
        "extras": [
            [r"tarjeta", "con la tarjeta visible en el acceso y la fecha de la última revisión"],
            [r"rodapie|barand", "con barandas y rodapié continuos en todo el perímetro"],
            [r"peldano|escalera|acceso", "con escalera de acceso asegurada"],
            [r"tijera|elevadora", "con inspección previa de la plataforma elevadora"],
        ],
        "criterio": "Andamio o plataforma sin tarjeta de liberación vigente, sin baranda ni rodapié, o con peldaños faltantes.",
        "herramienta": (
            "Inspección de Plataformas + CdP Altura + ASP: verificar tarjeta, barandas, rodapiés "
            "y peldaños de cada andamio en uso."
        ),
        "evidencia": "Tarjetas de liberación vigentes / Registro de Inspección de Plataformas / registro fotográfico",
        "control_corto": "Tarjeta de andamios y plataformas",
        "brecha_cdp": "liberación de andamios y plataformas",
        "pendiente": "liberación de andamios y plataformas",
    },
    # ------------------------------------------------------------------ 7
    {
        "id": "altura_caidas",
        "tema": "Protección contra caídas en altura",
        "nombre_corto": "protección contra caídas deficiente",
        "palabras": [
            r"trabajo(s)? en altura",
            r"\barnes(es)?\b",
            r"linea(s)? de vida",
            r"doble (gancho|cola|rabo)",
            r"proteccion contra caidas?",
            r"\bcaida(s)? (de|a) (distinto |diferente )?(altura|nivel)",
            r"\banclaje(s)?\b",
            r"sin (arnes|anclaje)",
            r"borde(s)? (de losa|sin proteccion|expuesto)",
            r"\bcaida(s)? de objetos",
            r"\bamortiguador(es)? de impacto",
            r"\bdesnivel(es)?\b",
        ],
        "excluir": r"altura (de|del) (acopio|apilamiento|pila|material)|altura geografica|gran altura",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que todo trabajo sobre 1,8 metros use arnés con doble gancho anclado a "
            "punto o línea de vida certificados, sin exposición a bordes sin protección; el "
            "supervisor lo comprueba antes de iniciar la tarea y en cada cambio de nivel."
        ),
        "extras": [
            [r"linea(s)? de vida", "con línea de vida inspeccionada y anclajes certificados"],
            [r"doble (gancho|cola)", "manteniendo siempre un gancho conectado durante el desplazamiento"],
            [r"caida(s)? de objetos", "con barreras y red contra caída de objetos"],
        ],
        "criterio": "Trabajo en altura sin arnés, sin doble gancho o sin punto de anclaje certificado.",
        "herramienta": (
            "CdP Altura + Desafío en Terreno + PASSO: observar anclaje, doble gancho y línea de "
            "vida en trabajos sobre 1,8 m, en cada frente activo, dos veces por turno."
        ),
        "evidencia": "Registro de inspección de arnés y líneas de vida / PTS con trabajo en altura / registro fotográfico",
        "control_corto": "Arnés, doble gancho y línea de vida",
        "brecha_cdp": "protección contra caídas",
        "pendiente": "protección contra caídas",
    },
    # ------------------------------------------------------------------ 8
    {
        "id": "tableros_electricos",
        "tema": "Tableros eléctricos y energización",
        "nombre_corto": "tableros y líneas eléctricas sin control",
        "palabras": [
            r"\btablero(s)? (electrico|electricos|de distribucion|de faena|provisorio)",
            r"\btablero(s)?\b",
            r"linea(s)? (por|a) energizar",
            r"\benergizad(o|a|os|as)\b",
            r"\benergizar\b",
            r"\benergizacion\b",
            r"\bcable(s)? (energizado|vivo)",
            r"\bdiferencial(es)?\b",
            r"\bmedia tension\b|\balta tension\b|\bbaja tension\b",
            r"canalizacion(es)? electrica",
            r"puesta a tierra",
            r"\bshock electrico|\bcontacto electrico|\barco electrico",
            r"\belectricista(s)?\b",
            r"\bsubestacion(es)?\b",
        ],
        "excluir": r"tablero (informativo|de avisos|de comunicaciones)|diario mural",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que tableros eléctricos y líneas por energizar estén cerrados, "
            "rotulados, con puesta a tierra y protecciones diferenciales operativas, y que solo "
            "personal autorizado los intervenga; el electricista lo valida y el supervisor lo "
            "confirma antes de energizar."
        ),
        "extras": [
            [r"puesta a tierra", "con continuidad de la puesta a tierra medida"],
            [r"diferencial", "con prueba del diferencial registrada"],
            [r"energiz", "con aviso al personal del área antes de energizar"],
            [r"rotul", "con todos los circuitos rotulados"],
        ],
        "criterio": "Tablero abierto o sin rotular, conductor expuesto o línea por energizar sin puesta a tierra ni autorización.",
        "herramienta": (
            "CdP Eléctrico + PTS + ASP: revisar el estado de tableros y líneas por energizar y su "
            "autorización de energización, en cada frente, antes de energizar."
        ),
        "evidencia": "Registro de revisión de tableros / Autorización de energización / registro fotográfico",
        "control_corto": "Tableros eléctricos y puesta a tierra",
        "brecha_cdp": "tableros eléctricos y energización",
        "pendiente": "tableros eléctricos y líneas por energizar",
    },
    # ------------------------------------------------------------------ 9
    {
        "id": "extensiones_electricas",
        "tema": "Extensiones eléctricas y pasacables",
        "nombre_corto": "extensiones eléctricas en mal estado",
        "palabras": [
            r"extension(es)? electrica(s)?",
            r"\balargador(es)?\b",
            r"\bpasacable(s)?\b",
            r"\bempalme(s)?\b",
            r"\benchufe(s)?\b",
            r"\bcable(s)? (en el suelo|por el piso|tirado|expuesto|sin proteccion|cruzando|pelado)",
            r"\bconductor(es)? expuesto",
            r"\bcable(s)? electrico(s)?",
            r"\bcable(s)? da[nñ]ado",
            r"\bconexion(es)? improvisada",
        ],
        "excluir": r"extension (de plazo|de contrato|de jornada|de horario)",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que extensiones eléctricas y alargadores estén sin empalmes ni "
            "conductor expuesto, con enchufes industriales, protección diferencial y pasacables "
            "donde crucen tránsito; el usuario los revisa y el supervisor los valida al inicio del turno."
        ),
        "extras": [
            [r"empalme", "retirando de uso toda extensión con empalmes"],
            [r"enchufe", "usando solo enchufes industriales"],
            [r"piso|suelo|cruz", "protegiendo los cables que cruzan zonas de tránsito"],
        ],
        "criterio": "Extensión con empalme, conductor expuesto o enchufe no industrial; cable cruzando tránsito sin pasacables.",
        "herramienta": (
            "CdP Eléctrico + Desafío en Terreno + ASP: recorrer las extensiones y pasacables en "
            "uso, en los frentes con cables que cruzan tránsito."
        ),
        "evidencia": "Registro de inspección de extensiones / ASP del frente / registro fotográfico",
        "control_corto": "Extensiones eléctricas y pasacables",
        "brecha_cdp": "extensiones eléctricas y pasacables",
        "pendiente": "extensiones eléctricas y pasacables",
    },
    # ------------------------------------------------------------------ 10
    {
        "id": "bloqueo_energias",
        "tema": "Bloqueo y energías peligrosas",
        "nombre_corto": "bloqueo de energías peligrosas deficiente",
        "palabras": [
            r"\bbloqueo(s)?\b",
            r"\bcandado(s)?\b",
            r"tarjeta(s)? de bloqueo",
            r"energia(s)? (peligrosa|residual|almacenada|cero)",
            r"energias peligrosas",
            r"\blockout\b|\bloto\b",
            r"aislamiento de energia",
            r"\bdesenergizar|\bdesenergizado",
            r"\bsin bloquear\b|\bbloquear\b",
            r"punto(s)? de bloqueo",
            r"\bpartida inesperada|puesta en marcha inesperada",
            r"\bdespresuriz(ar|acion)",
            r"\batrapamiento\b",
        ],
        "excluir": r"bloqueo (de (acceso|accesos|via|vias|ruta|camino|transito)|vial)",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que toda intervención de equipos con energía peligrosa tenga bloqueo "
            "personal con candado y tarjeta, y comprobación de energía cero antes de iniciar; "
            "cada trabajador aplica su candado y el supervisor valida el aislamiento antes de la tarea."
        ),
        "extras": [
            [r"candado", "con un candado por trabajador"],
            [r"residual|almacenada|presion", "descargando la energía residual o almacenada"],
            [r"tarjeta", "con tarjeta de bloqueo visible en cada punto"],
        ],
        "criterio": "Intervención sin candado ni tarjeta personal, o sin comprobar energía cero.",
        "herramienta": (
            "CdP Bloqueo + PTS + Acompañamiento ASP: acompañar la aplicación de candados y la "
            "comprobación de energía cero en cada intervención, antes de iniciar el trabajo."
        ),
        "evidencia": "Registro de bloqueo y energía cero / PTS con aislamiento definido / registro fotográfico",
        "control_corto": "Bloqueo personal y energía cero",
        "brecha_cdp": "bloqueo y control de energías",
        "pendiente": "bloqueo y energías peligrosas",
    },
    # ------------------------------------------------------------------ 11
    {
        "id": "excavaciones_zanjas",
        "tema": "Excavaciones y zanjas",
        "nombre_corto": "excavaciones sin protección",
        "palabras": [
            r"\bexcavacion(es)?\b",
            r"\bzanja(s)?\b",
            r"\bzanjeo\b",
            r"\bentibacion(es)?\b",
            r"\btalud(es)?\b",
            r"\bsocavon(es)?\b",
            r"\bderrumbe(s)?\b",
            r"\bsobreexcavacion",
            r"\bcalicata(s)?\b",
            r"catastro de interferencias|servicios enterrados|interferencias subterraneas",
            r"borde(s)? de (la )?(excavacion|zanja)",
            r"material excavado",
            r"\bapuntalamiento",
        ],
        "excluir": r"hallazgo arqueologico|arqueolog",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que toda excavación o zanja cuente con autorización, entibación o "
            "talud según el suelo, acceso seguro, borde despejado y barreras perimetrales; el "
            "supervisor lo inspecciona antes de que el personal ingrese y al inicio de cada turno."
        ),
        "extras": [
            [r"borde|acopio|material excavado", "manteniendo el material excavado alejado del borde"],
            [r"servicios enterrados|interferencia", "con catastro de servicios enterrados antes de excavar"],
            [r"entibacion|talud", "con entibación calculada para el tipo de suelo"],
            [r"\bagua\b|inundacion|\bnapa\b", "controlando el ingreso de agua a la excavación"],
        ],
        "criterio": "Excavación sin autorización, sin entibación o talud, o con personal dentro sin acceso seguro.",
        "herramienta": (
            "CdP Excavaciones + PTS + ASP: inspeccionar autorización, entibación o talud y accesos "
            "de cada excavación abierta, antes del ingreso."
        ),
        "evidencia": "Autorización de excavación / Registro de inspección de zanjas / registro fotográfico",
        "control_corto": "Excavaciones autorizadas y protegidas",
        "brecha_cdp": "protección y autorización de excavaciones",
        "pendiente": "excavaciones y zanjas",
    },
    # ------------------------------------------------------------------ 12
    {
        "id": "espacios_confinados",
        "tema": "Espacios confinados y atmósferas",
        "nombre_corto": "ingresos a espacios confinados sin control",
        "palabras": [
            r"espacio(s)? confinado(s)?",
            r"\bconfinad(o|a|os|as)\b",
            r"\batmosfera(s)?\b",
            r"\bgases\b",
            r"\boxigeno\b",
            r"\bmonoxido\b",
            r"\bh2s\b|sulfuro de hidrogeno",
            r"\bexplosimetro\b|detector(es)? de gases",
            r"medicion(es)? de (gases|atmosfera|oxigeno)",
            r"ventilacion forzada|\bextractor(es)?\b",
            r"\bvigia\b",
            r"\bestanque(s)?\b",
            r"\bsilo(s)?\b",
            r"\bcisterna(s)?\b",
            r"camara(s)? (subterranea|de inspeccion)",
            r"rescate en espacio",
        ],
        "excluir": r"cuarentena|covid|aislamiento sanitario",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que todo ingreso a espacio confinado cuente con permiso vigente, "
            "medición de oxígeno y gases antes y durante la tarea, ventilación, vigía en el "
            "exterior y plan de rescate; el supervisor autoriza y el vigía controla el acceso."
        ),
        "extras": [
            [r"ventilacion", "con ventilación forzada operando"],
            [r"h2s|sulfuro", "con medición de H2S y detector personal"],
            [r"rescate", "con equipo de rescate disponible en el acceso"],
            [r"vigia", "con vigía exclusivo en el acceso"],
        ],
        "criterio": "Ingreso sin permiso, sin medición de atmósfera vigente, sin vigía o sin plan de rescate.",
        "herramienta": (
            "CdP Espacios Confinados + PTS + ASP: verificar permiso, medición de atmósfera, vigía "
            "y plan de rescate en cada ingreso, antes de entrar y durante la tarea."
        ),
        "evidencia": "Permiso de ingreso / Registro de medición de atmósfera / Plan de rescate / registro fotográfico",
        "control_corto": "Permiso y monitoreo de atmósfera",
        "brecha_cdp": "control de espacios confinados y atmósferas",
        "pendiente": "espacios confinados y atmósferas",
    },
    # ------------------------------------------------------------------ 13
    {
        "id": "equipos_emergencia",
        "tema": "Equipos de emergencia",
        "nombre_corto": "equipos de emergencia con falencias",
        "palabras": [
            r"\bextintor(es)?\b",
            r"\bducha(s)? de emergencia",
            r"\bduchas?\b",
            r"\blavaojos\b|\blava ojos\b",
            r"\bvia(s)? de evacuacion",
            r"\bruta(s)? de evacuacion",
            r"\bzona(s)? de (seguridad|reunion)",
            r"\bpunto(s)? de reunion",
            r"\bsalida(s)? de emergencia",
            r"\bbotiquin(es)?\b|\bkit de emergencia",
            r"\bcamilla(s)?\b",
            r"detector(es)? de humo",
            r"\bgrifo(s)?\b|\bred humeda\b|\bred seca\b",
            r"se[nñ]al(etica|izacion) de emergencia",
        ],
        "excluir": r"extintor (del|de) (vehiculo|camioneta|equipo movil)",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que extintores, duchas de emergencia y lavaojos estén operativos, "
            "señalizados y sin obstrucciones, y que las vías de evacuación y zonas de reunión "
            "estén despejadas; el supervisor los revisa al inicio del turno y tras cada traslado de materiales."
        ),
        "extras": [
            [r"ducha|lavaojos", "probando el flujo de duchas y lavaojos"],
            [r"extintor", "con extintores cargados, con sello y revisión vigente"],
            [r"evacuacion|salida", "con vías y salidas libres y señalizadas"],
        ],
        "criterio": "Extintor, ducha o vía de evacuación bloqueado, sin carga, sin señalizar o fuera de su lugar.",
        "herramienta": (
            "ASP + Desafío en Terreno + PASSO: revisar estado y ubicación de extintores, duchas y "
            "vías de evacuación en toda el área."
        ),
        "evidencia": "Registro de revisión de equipos de emergencia / Plano de vías de evacuación / registro fotográfico",
        "control_corto": "Extintores y vías de evacuación",
        "brecha_cdp": "equipos y vías de emergencia",
        "pendiente": "equipos de emergencia",
    },
    # ------------------------------------------------------------------ 14
    {
        "id": "trabajos_caliente",
        "tema": "Trabajos en caliente",
        "nombre_corto": "trabajos en caliente sin control",
        "palabras": [
            r"trabajo(s)? (en|con) caliente",
            r"\bsoldadura(s)?\b",
            r"\bsoldar\b|\bsoldador(es)?\b",
            r"\boxicorte\b|\boxiacetilen",
            r"\bcorte (con )?(plasma|oxigeno|disco)",
            r"\besmeril(es)?\b|\bamoladora(s)?\b|\bpulidora(s)?\b",
            r"\bchispa(s)?\b",
            r"permiso (de trabajo )?(en )?caliente",
            r"\bmanta(s)? ignifuga",
            r"\bvigia de fuego",
            r"cilindro(s)? de (gas|oxigeno|acetileno)",
            r"\bpantalla(s)? de soldadura",
        ],
        "excluir": r"agua caliente|temperatura ambiente|golpe de calor|estres termico",
        "bloque": "mediodia",
        "cc_o_emergencia": True,
        "instruccion": (
            "Verificar que todo trabajo en caliente (soldadura, corte o esmeril) cuente con "
            "permiso vigente, área despejada de combustibles, extintor al alcance, pantallas "
            "contra chispas y vigía de fuego; el supervisor lo autoriza antes de iniciar y al retomar la tarea."
        ),
        "extras": [
            [r"cilindro|oxiacetilen", "con cilindros asegurados y válvulas antirretorno"],
            [r"manta|chispa", "con manta ignífuga que contenga las chispas"],
            [r"altura|nivel inferior", "protegiendo el nivel inferior de la caída de chispas"],
        ],
        "criterio": "Trabajo en caliente sin permiso, con combustibles en el área o sin extintor ni vigía de fuego.",
        "herramienta": (
            "CdP Trabajos en Caliente + PTS + ASP: validar permiso, extintor y vigía de fuego en "
            "cada punto de soldadura o corte, antes de iniciar y al retomar la tarea."
        ),
        "evidencia": "Permiso de trabajo en caliente / Registro de revisión del área / registro fotográfico",
        "control_corto": "Permiso de trabajo en caliente",
        "brecha_cdp": "permiso y control de trabajos en caliente",
        "pendiente": "trabajos en caliente",
    },
    # ------------------------------------------------------------------ 15
    {
        "id": "herramientas_manuales",
        "tema": "Herramientas manuales y portátiles",
        "nombre_corto": "herramientas en mal estado o sin resguardo",
        "palabras": [
            r"herramienta(s)? (manual|manuales|portatil|portatiles|electrica|electricas|neumatica|neumaticas)",
            r"\bherramienta(s)?\b",
            r"\bresguardo(s)?\b",
            r"\bdisco(s)? de (corte|desbaste)",
            r"\bsin resguardo",
            r"\btaladro(s)?\b",
            r"\bsierra(s)?\b",
            r"\bpistola(s)? (neumatica|de impacto)",
            r"\bllave(s)? (de impacto|inglesa)",
            r"\bmartillo(s)?\b",
            r"\bcincel(es)?\b",
            r"\bmango(s)?\b",
            r"\bcutter\b|\bcuchillo(s)?\b",
        ],
        "excluir": r"herramienta(s)? (de gestion|informatica|digital)",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que herramientas manuales y portátiles tengan resguardos instalados, "
            "cables y mangos sin daño, discos o accesorios adecuados a la herramienta y revisión "
            "de estado vigente; el usuario las inspecciona y el supervisor las valida al inicio del turno."
        ),
        "extras": [
            [r"\bdisco", "usando discos acordes a la velocidad de la herramienta"],
            [r"resguardo", "sin retirar resguardos de fábrica"],
            [r"cable|electric", "con cables y enchufes sin daño"],
            [r"neumatic", "con mangueras y acoples asegurados"],
        ],
        "criterio": "Herramienta sin resguardo, con cable o mango dañado, o con disco inadecuado para su velocidad.",
        "herramienta": (
            "ASP + Inspección de Equipos + PASSO: revisar resguardos, cables y discos de las "
            "herramientas en uso, en cada frente."
        ),
        "evidencia": "Registro de inspección de herramientas / ASP del frente / registro fotográfico",
        "control_corto": "Resguardos y estado de herramientas",
        "brecha_cdp": "estado y resguardos de herramientas",
        "pendiente": "herramientas manuales y portátiles",
    },
    # ------------------------------------------------------------------ 16
    {
        "id": "housekeeping",
        "tema": "Acopio de material y housekeeping",
        "nombre_corto": "acopio de material y housekeeping deficiente",
        "palabras": [
            r"\bacopio(s)?\b",
            r"\bacopiad(o|a|os|as)\b",
            r"\bhousekeeping\b",
            r"orden y aseo",
            r"\bdesorden\b",
            r"\bdesordenad(o|a|os|as)\b",
            r"\bmaterial(es)? (en el suelo|esparcido|suelto|botado|tirado)",
            r"\bresiduo(s)?\b",
            r"\bbasura\b",
            r"\bescombro(s)?\b",
            r"\bobstruy(e|en|endo)\b|\bobstruccion\b|\bobstruido\b",
            r"\bobstaculo(s)?\b",
            r"\btropiezo(s)?\b|\btropezar\b|\bresbalon\b|\bresbalar\b",
            r"\bmadera(s)? con clavos|\bclavo(s)? expuesto",
            r"\bpila(s)? de material",
            r"\blimpieza\b",
        ],
        "excluir": r"pila de lixiviacion|acopio de mineral|stock ?pile",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que el material y los residuos estén acopiados en áreas demarcadas, "
            "sin obstruir vías de tránsito, accesos ni salidas de emergencia, y que el área "
            "quede ordenada y limpia; cada cuadrilla lo cumple y el supervisor lo revisa durante "
            "el turno y al cierre."
        ),
        "extras": [
            [r"clavo|madera", "retirando o doblando los clavos de la madera"],
            [r"manguera|cable", "ordenando mangueras y cables fuera de las vías"],
            [r"residuo|basura", "con contenedores de residuos rotulados"],
            [r"resbal|tropiez", "con pisos libres de riesgo de tropiezo"],
        ],
        "criterio": "Material que obstruye vías, accesos o salidas de emergencia, o pila inestable en zona de tránsito.",
        "herramienta": (
            "ASP + Desafío en Terreno + PASSO: recorrer frentes y accesos verificando acopios, "
            "vías y salidas despejadas."
        ),
        "evidencia": "Registro de recorrido de orden y aseo / ASP del frente / registro fotográfico",
        "control_corto": "Orden y acopio en áreas de tránsito",
        "brecha_cdp": "orden, aseo y acopio de material",
        "pendiente": "acopio de material y housekeeping",
    },
    # ------------------------------------------------------------------ 17
    {
        "id": "sustancias_derrames",
        "tema": "Sustancias peligrosas y derrames",
        "nombre_corto": "manejo de sustancias peligrosas deficiente",
        "palabras": [
            r"sustancia(s)? peligrosa(s)?",
            r"\bhds\b",
            r"hoja(s)? de datos de seguridad",
            r"\bderrame(s)?\b",
            r"\bderramad(o|a|os|as)\b|\bderramo\b|\bderramar\b",
            r"\brotulad(o|a|os|as)\b|\brotulo(s)?\b|\brotulacion\b",
            r"\bcontencion\b",
            r"\bantiderrame\b",
            r"\bcombustible(s)?\b",
            r"\bdiesel\b|\bpetroleo\b",
            r"\baceite(s)?\b|\blubricante(s)?\b|\bgrasa(s)?\b",
            r"\bquimico(s)?\b",
            r"\bacido(s)?\b|\bsolvente(s)?\b|\bthinner\b|\bsoda caustica\b",
            r"\btrasvasij(e|ar|ado)",
            r"\bbidon(es)?\b|\btambor(es)?\b",
            r"\bhidrocarburo(s)?\b",
            r"\bfuga(s)? de (aceite|combustible|liquido|producto)",
        ],
        "excluir": r"\bdrogas?\b|\balcohol(emia)?\b|\balcotest\b",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que las sustancias peligrosas estén rotuladas, con hoja de datos de "
            "seguridad (HDS) disponible, almacenadas en contención secundaria y con kit "
            "antiderrame cercano; el usuario lo mantiene y el supervisor lo revisa antes de usar "
            "o trasvasijar el producto."
        ),
        "extras": [
            [r"derrame", "con plan de respuesta ante derrames conocido por el equipo"],
            [r"combustible|diesel|petroleo", "con transferencia de combustible en zona contenida"],
            [r"bidon|tambor", "con envases rotulados y cerrados"],
        ],
        "criterio": "Sustancia sin rotular ni HDS, sin contención, o derrame sin kit antiderrame disponible.",
        "herramienta": (
            "CdP Sustancias Peligrosas + ASP + PASSO: revisar rotulado, HDS y contención en los "
            "puntos de almacenamiento y uso, antes de trasvasijar o usar el producto."
        ),
        "evidencia": "HDS disponibles en terreno / Registro de revisión de rotulado y contención / registro fotográfico",
        "control_corto": "Rotulado, HDS y contención",
        "brecha_cdp": "rotulado y contención de sustancias",
        "pendiente": "sustancias peligrosas y derrames",
    },
    # ------------------------------------------------------------------ 18
    {
        "id": "manejo_manual",
        "tema": "Manejo manual de cargas",
        "nombre_corto": "manejo manual de cargas deficiente",
        "palabras": [
            r"manejo manual",
            r"\bcarga(s)? manual(es)?",
            r"levantamiento manual",
            r"\bpeso (excesivo|maximo)",
            r"\bsobreesfuerzo(s)?\b|\bsobre esfuerzo",
            r"\bpostura(s)?\b",
            r"\bergonomic(o|a|os|as)\b|\bergonomia\b",
            r"\btraslado manual|\btrasladar (a mano|manualmente)",
            r"\bcargar (a mano|manualmente)",
            r"\bayuda(s)? mecanica(s)?",
            r"\blumbar(es)?\b|\bespalda\b",
            r"\bcarretilla(s)?\b|\bcarro(s)? de mano",
            r"\blevant(ar|ando|o) (un |una |el |la )?(saco|bulto|tambor|plancha|tubo|pieza)",
        ],
        "excluir": r"manejo defensivo",
        "bloque": "mediodia",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que el manejo manual de cargas respete el peso máximo permitido, use "
            "ayudas mecánicas o apoyo de otro trabajador cuando corresponda y mantenga una "
            "postura segura; el supervisor lo evalúa antes de iniciar traslados y al cambiar de tarea."
        ),
        "extras": [
            [r"saco|bulto", "dividiendo los sacos o bultos pesados"],
            [r"postura|espalda|lumbar", "con pausas y rotación de tareas"],
            [r"plancha|tubo", "con apoyo de dos personas en cargas largas"],
        ],
        "criterio": "Carga sobre el límite de peso o postura insegura sin ayuda mecánica ni segundo trabajador.",
        "herramienta": (
            "ASP + CdP Manejo Manual de Cargas + PASSO: observar traslados y posturas de trabajo "
            "en los frentes, corrigiendo en el momento, en un recorrido por turno."
        ),
        "evidencia": "ASP de manejo manual / Registro de observaciones de postura / registro fotográfico",
        "control_corto": "Manejo manual y ayudas mecánicas",
        "brecha_cdp": "manejo manual de cargas",
        "pendiente": "manejo manual de cargas",
    },
    # ------------------------------------------------------------------ 19
    {
        "id": "epp",
        "tema": "Uso y estado de EPP",
        "nombre_corto": "uso inadecuado de EPP",
        "palabras": [
            r"\bepp\b",
            r"\bcasco(s)?\b",
            r"\blentes?\b|\bantiparra(s)?\b",
            r"\bguante(s)?\b",
            r"\b(zapato|zapatos|calzado|bota|botas) de seguridad",
            r"\bprotector(es)? (auditivo|facial|respiratorio)|\bproteccion (auditiva|facial|respiratoria|visual)",
            r"\brespirador(es)?\b|\bmascarilla(s)?\b",
            r"\bchaleco(s)?\b",
            r"\bsin (casco|lentes|guantes|chaleco|barbiquejo)",
            r"\bbarbiquejo\b",
            r"equipo(s)? de proteccion personal",
            r"\bropa de trabajo\b",
        ],
        "excluir": r"entrega de epp|stock de epp|bodega de epp",
        "bloque": "am",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que cada trabajador use el EPP básico y específico de la tarea (casco, "
            "lentes, guantes, calzado y arnés si aplica), en buen estado y vigente; el "
            "supervisor lo revisa al inicio del turno y reemplaza lo deteriorado antes de continuar."
        ),
        "extras": [
            [r"respirador|mascarilla", "con filtros del respirador vigentes"],
            [r"lente|antiparra", "con protección visual adecuada al riesgo"],
            [r"guante", "con guantes según el riesgo de la tarea"],
            [r"auditiv", "con protección auditiva en zonas de ruido"],
        ],
        "criterio": "Persona sin EPP obligatorio o con EPP dañado o vencido en la tarea en curso.",
        "herramienta": (
            "ASP + PASSO + Desafío en Terreno: revisar EPP completo y su estado en cada frente, "
            "corrigiendo en el momento, en dos recorridos por turno."
        ),
        "evidencia": "Registro de inspección de EPP / ASP del frente / registro fotográfico",
        "control_corto": "EPP completo y en buen estado",
        "brecha_cdp": "uso y estado de EPP",
        "pendiente": "uso y estado de EPP",
    },
    # ------------------------------------------------------------------ 20
    {
        "id": "documentacion_terreno",
        "tema": "Documentación en terreno",
        "nombre_corto": "documentación en terreno sin actualizar",
        "palabras": [
            r"\bpts\b",
            r"\bart\b",
            r"\bcdp\b",
            r"\bmiper\b",
            r"\bsin firma\b|\bsin firmar\b|\bfirmas? (faltante|incompleta)",
            r"\bsin actualizar\b|\bdesactualizad(o|a|os|as)\b",
            r"actualizacion de (pts|art|miper|procedimiento)",
            r"\bprocedimiento(s)?\b",
            r"analisis de riesgo",
            r"\bcharla(s)? (diaria|de (5|cinco) minutos)",
            r"\bdocumento(s)? (incompleto|desactualizado|sin firma)",
            r"\bpermiso(s)? (vencido|faltante)|\bsin permiso\b",
            r"matriz de riesgo|identificacion de peligros",
            r"gestion del cambio",
        ],
        "excluir": r"documentacion (medica|contable)|licencia medica",
        "bloque": "am",
        "cc_o_emergencia": False,
        "instruccion": (
            "Verificar que el PTS, ART, CdP y MIPER de la tarea estén actualizados al trabajo "
            "real, firmados por todo el equipo y con los controles críticos identificados; el "
            "supervisor los revisa en terreno antes de iniciar y ante cualquier cambio de condiciones."
        ),
        "extras": [
            [r"firma", "con firma de todos los participantes antes de iniciar"],
            [r"cambio", "reemplazando el documento ante cambios de alcance o condiciones"],
            [r"miper|matriz", "con la MIPER vinculada a los controles críticos"],
        ],
        "criterio": "PTS, ART o CdP sin firma, desactualizado o que no representa la tarea que se ejecuta.",
        "herramienta": (
            "PTS + Gestión del Cambio + ASP: contrastar PTS, ART y CdP con la tarea real en cada "
            "frente y actualizarlos ante cambios, al inicio del turno."
        ),
        "evidencia": "PTS, ART y CdP actualizados y firmados / Registro de revisión en terreno / registro fotográfico",
        "control_corto": "PTS, ART y CdP actualizados en terreno",
        "brecha_cdp": "actualización y firma de PTS, ART y CdP",
        "pendiente": "documentación en terreno",
    },
]

# ---------------------------------------------------------------------------
# REVAL: reglas de revalorización (solo SUGERENCIA; el usuario decide).
# Todas las regex van en minúscula y sin tildes, sobre texto normalizado.
# ---------------------------------------------------------------------------
REVAL = {
    # Un evento solo sube. Transiciones válidas por tipo de origen.
    "transiciones": {
        "HBP": ["CBP", "HAP", "CA"],
        "CBP": ["CA"],
        "HAP": ["CA"],
    },

    # Persona mencionada en la descripción.
    "persona": (
        r"\b(trabajador(es)?|persona(s)?|rigger(s)?|operador(es)?|operadora|supervisor(es)?|"
        r"ayudante(s)?|maestro(s)?|soldador(es)?|electricista(s)?|mecanico(s)?|capataz|"
        r"peon(es)?|senalero(s)?|conductor(es)?|personal|cuadrilla|jefe de turno|"
        r"prevencionista|montajista(s)?|andamiero(s)?|instrumentista(s)?|gruero(s)?)\b"
    ),

    # Indica que NO había persona presente (anula la exposición).
    "sin_persona": (
        r"\b(sin (personas?|personal|trabajadores?|ocupantes)|nadie|ninguna persona|"
        r"area despejada|equipo detenido|sin presencia de personas)\b"
    ),

    # Tarea de riesgo efectivamente en curso.
    "tarea_en_curso": (
        r"\b(se encontraba(n)?|estaba(n)?|realizaba(n)?|ejecutaba(n)?|efectuaba(n)?|"
        r"trabajaba(n)?|transitaba(n)?|operaba(n)?|izaba(n)?|maniobraba(n)?|descendia(n)?|"
        r"ascendia(n)?|excavaba(n)?|soldaba(n)?|en ejecucion|en pleno (trabajo|izaje|maniobra)|"
        r"durante (la|el) (tarea|maniobra|trabajo|izaje|operacion|excavacion|montaje)|"
        r"mientras (se )?(ejecuta|ejecutaba|realiza|realizaba|izaba|operaba|trabajaba))\b"
    ),

    # Contexto que NO es tarea en curso (inspección, recorrido, auditoría).
    "sin_tarea": (
        r"\b(durante (una |un |la |el )?(inspeccion|recorrido|auditoria|revision documental)|"
        r"al inspeccionar|al revisar|en inspeccion|en recorrido)\b"
    ),

    # Negación / potencial: bajan la confianza de la sugerencia.
    "negacion_potencial": (
        r"\b(podria(n)?|pudo|pudiera(n)?|eventual(mente)?|potencial(mente)?|riesgo de|"
        r"posible(mente)?|probable|habria|hubiera|hubiese|de haber|en caso de|"
        r"susceptible|pudiendo|existe la posibilidad|no hubo (personas|lesion|contacto))\b"
    ),

    # Exposiciones que justifican subir a CA. Cada una exige "persona" en el texto.
    "exposiciones": {
        "linea_fuego_equipo": {
            "etiqueta": "línea de fuego de equipo o vehículo en movimiento",
            "regex": [
                r"linea(s)? de fuego",
                r"(equipo|vehiculo|camion|camioneta|excavadora|cargador|retroexcavadora|maquina|grua) "
                r"(en movimiento|en retroceso|retrocediendo|avanzando|girando|en operacion)",
                r"(a|en) (menos de )?\d+([.,]\d+)? ?(m|mt|mts|metro|metros) (del?|al) "
                r"(equipo|vehiculo|camion|camioneta|excavadora|cargador|maquina)",
                r"\batropell",
                r"\baculat",
                r"(peaton|persona|trabajador|personal)\w* (dentro|en) (del )?(radio|area) de (giro|operacion|maniobra)",
                r"(transitaba|caminaba|cruzaba|se ubicaba)\w* (frente|delante|detras|junto) (a|de|al) "
                r"(equipo|vehiculo|camion|camioneta|excavadora|cargador|maquina)",
            ],
            "frase": "persona expuesta a la línea de fuego de un equipo o vehículo en movimiento",
        },
        "carga_suspendida": {
            "etiqueta": "carga suspendida",
            "regex": [
                r"(bajo|debajo de|debajo del|bajo el|bajo la) (la )?(carga|pluma|izaje)",
                r"carga suspendida",
                r"(persona|trabajador|personal|rigger|ayudante)\w* "
                r"(bajo|debajo|dentro del (radio|area)|en el area de influencia)",
                r"(se encontraba|estaba|permanecia)\w* (bajo|debajo de) (la )?(carga|pluma)",
                r"area de influencia de la carga",
                r"carga (en el aire|izada|levantada) (sobre|con) (personas|personal|trabajadores)",
            ],
            "frase": "persona expuesta a carga suspendida",
        },
        "energia_energizada": {
            "etiqueta": "energía peligrosa energizada",
            "regex": [
                r"(cable|conductor|linea|tablero|equipo|barra|celda|circuito)\w* "
                r"(energizad|con tension|vivo|con energia)",
                r"(contacto|toque|descarga|arco) electrico",
                r"sin (bloqueo|candado|energia cero|desenergizar)",
                r"(intervino|intervenia|manipulo|manipulaba|trabajo|trabajaba)\w* "
                r"(en|sobre) (equipo|tablero|linea|cable)\w* energizad",
                r"energia (residual|almacenada|peligrosa) (liberada|presente|sin controlar)",
            ],
            "frase": "persona expuesta a energía peligrosa energizada",
        },
        "altura_sin_proteccion": {
            "etiqueta": "altura sin protección contra caídas",
            "regex": [
                r"sin (arnes|linea de vida|anclaje|proteccion contra caidas?)",
                r"(arnes|doble gancho|cola) (sin conectar|desconectado|sin anclar)",
                r"(borde|abertura|losa|plataforma|andamio|techo|estructura)\w* "
                r"sin (baranda|proteccion|rodapie|tapa|cierre)",
                r"(a|sobre|en) (mas de )?(1[.,]8|2|3|4|5|6|8|10|12) ?(m|mt|mts|metro|metros) de altura",
                r"trabajo(s)? en altura sin",
                r"(en|desde) (altura|techo|estructura|andamio|plataforma) sin (arnes|proteccion|anclaje)",
            ],
            "frase": "persona trabajando en altura sin protección contra caídas",
        },
        "zanja_sin_proteccion": {
            "etiqueta": "zanja o excavación sin protección o autorización",
            "regex": [
                r"(dentro|en el interior|al interior) de (la |una )?(zanja|excavacion)",
                r"(zanja|excavacion)\w* sin (entibacion|talud|proteccion|autorizacion|permiso|barrera|acceso)",
                r"sin (entibacion|talud|autorizacion de excavacion)",
                r"(ingreso|ingreso a|trabajaba en|trabajaban en|descendio a) (la |una )?(zanja|excavacion)",
                r"(talud|pared)\w* (inestable|sin proteccion|desprendimiento)",
            ],
            "frase": "persona dentro de una zanja o excavación sin protección o autorización",
        },
        "atmosfera_confinado": {
            "etiqueta": "atmósfera peligrosa o espacio confinado sin control",
            "regex": [
                r"(ingreso|ingreso a|ingresaron a|ingreso al|trabajaba en|trabajaban en) "
                r"(un |el |la )?(espacio confinado|estanque|silo|cisterna|camara)",
                r"espacio confinado sin (permiso|medicion|vigia|ventilacion|rescate)",
                r"sin (medicion|monitoreo) de (atmosfera|gases|oxigeno)",
                r"(atmosfera|ambiente) (peligrosa|deficiente|toxica|explosiva|inflamable)",
                r"(gas|h2s|monoxido|sulfuro)\w* (presente|detectado|sobre el limite|en el ambiente)",
                r"sin vigia en (el )?(acceso|exterior)",
            ],
            "frase": "persona en espacio confinado o atmósfera peligrosa sin control",
        },
    },

    # Subida a CBP: persona que interactuó de cerca con la situación de riesgo.
    "interaccion_cercana": (
        r"\b(a (pocos|escasos) (metros|centimetros)|a (menos de )?\d+([.,]\d+)? ?(m|mt|mts|metro|metros|cm)\b|"
        r"muy cerca|cerca de|proximidad|junto a|al lado de|a un costado|se aparto|se retiro a tiempo|"
        r"esquivo|evito (el )?(golpe|contacto|impacto)|roce|rozo|casi (fue )?(golpead|atrapad|impactad)|"
        r"contacto (con|visual con)|interactuo|interactuaba|se acerco)\b"
    ),

    # Subida a HAP: control crítico no implementado o erosionado durante la tarea.
    "control_critico_erosionado": (
        r"\b(sin (plan de izaje|permiso|autorizacion|bloqueo|candado|arnes|anclaje|segregacion|"
        r"certificado|tarjeta|vigia|medicion|senalero|barrera|entibacion|check ?list)|"
        r"no (contaba|contaban|tenia|tenian|se aplico|se implemento|estaba implementado|estaba vigente)|"
        r"omitio|omitieron|omision|incumpl\w+|vencid[oa]s?|sin vigencia|control (anulado|erosionado|ausente|inexistente)|"
        r"puenteo|inhabilit\w+)\b"
    ),

    # Casos que NO se revalorizan aunque el título suene grave.
    "no_revalorizar": (
        r"\b(inspeccion sin tarea|durante (una |un |la |el )?(inspeccion|recorrido|auditoria)|"
        r"hallazgo documental|observacion documental|podria haber (pasado|ocurrido)|"
        r"no hubo (personas|personal|exposicion)|sin personas expuestas|equipo detenido|"
        r"area despejada|condicion (potencial|latente))\b"
    ),

    # Criterios de confianza.
    "confianza": {
        # "alta": persona + exposición + tarea en curso y sin negación/potencial.
        "alta": ["persona", "exposicion", "tarea_en_curso", "sin_negacion_potencial"],
        # "baja": falta alguno de los anteriores o hay negación/potencial.
        "baja": ["persona", "exposicion"],
    },

    # Texto base de motivo por destino (para componer la sugerencia).
    "motivos": {
        "CA": "La descripción muestra a una persona efectivamente expuesta a {exposicion}.",
        "CBP": "La descripción indica que una persona interactuó de cerca con la situación de riesgo, sin potencial grave evidente.",
        "HAP": "La descripción indica un control crítico no implementado o erosionado mientras se ejecutaba una tarea de riesgo crítico.",
    },
}


# temas adicionales (cobertura de eventos de obra/mina)
try:
    from catalogo_extra import CATALOGO_EXTRA
    CATALOGO = CATALOGO + CATALOGO_EXTRA
except ImportError:
    pass

try:
    from catalogo_ajustes import AJUSTES
    for _t in CATALOGO:
        _a = AJUSTES.get(_t["id"])
        if not _a:
            continue
        if _a.get("tema"):
            _t["tema"] = _a["tema"]
        if _a.get("bonus"):
            _t["bonus"] = list(_a["bonus"])
        if _a.get("no_foco_solo_hbp"):
            _t["no_foco_solo_hbp"] = True
        if _a.get("excluir_mas"):
            _e = _t.get("excluir")
            _t["excluir"] = (_e + "|" if isinstance(_e, str) and _e else "") + _a["excluir_mas"]
except ImportError:
    pass

try:
    from catalogo_ajustes import REVAL_EXPOSICIONES
    for _k, _rs in REVAL_EXPOSICIONES.items():
        _x = REVAL["exposiciones"].get(_k)
        if isinstance(_x, dict):
            _x["regex"] = list(_x.get("regex") or []) + list(_rs)
except ImportError:
    pass

try:
    from catalogo_ajustes import HERRAMIENTA_CABEZA
    for _t in CATALOGO:
        _h = HERRAMIENTA_CABEZA.get(_t["id"])
        if _h and ":" in _t["herramienta"]:
            _t["herramienta"] = _h + ":" + _t["herramienta"].split(":", 1)[1]
except ImportError:
    pass
