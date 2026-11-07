"""Ajustes finos del catálogo (palabras de peso y exclusiones) por tema. Sin datos de Antucoya.
Campos: bonus = [[regex, puntos]] suma puntos si calza en título+descripción;
        excluir_mas = regex que se agrega a 'excluir'; no_foco_solo_hbp = el tema nunca forma foco solo con HBP (queda en Recurrente)."""
AJUSTES = {
    "izaje_segregacion": {"bonus": [
        [r"segregaci\w+.{0,40}izaj|izaj.{0,40}segregaci\w+|maxisaco", 4],
        [r"eslinga\w*.{0,60}(zanja|borde)|(zanja|borde).{0,60}eslinga", 3]]},
    "segregacion_acceso": {"excluir_mas": r"izaj|levante|maxisaco|carga suspendida"},
    "andamios_plataformas": {"bonus": [
        [r"(acopio|material\w*|pieza\w*).{0,40}(plataforma|andamio)|(plataforma|andamio).{0,40}(acopio|material\w*|pieza\w*)", 4]]},
    "documentacion_terreno": {"no_foco_solo_hbp": True},
    "izaje_certificacion": {"tema": "Certificación y línea de fuego", "bonus": [
        [r"certificaci\w+.{0,60}(grillete|eslinga|tecle|accesorio|elemento)|(grillete|eslinga|tecle).{0,60}(certific|placa)", 7],
        [r"rigger.{0,120}linea de fuego|linea de fuego.{0,120}rigger|aculat\\w+", 6]]},
    "transito_peaton_equipo": {"excluir_mas": r"rigger|camion pluma|izaj"},
    "tableros_electricos": {"bonus": [[r"tablero\w* electric\w+", 3]]},
    "housekeeping": {"bonus": [[r"housekeeping|orden y aseo", 3]]},
    "equipos_emergencia": {"bonus": [[r"extinci\w+|extintor\w*|manometro", 3]]},
    "extensiones_electricas": {"bonus": [[r"extension(es)? electric\w+|pasacable\w*", 3]]},
}

# nombres de herramientas SGE como las usa la gerencia (cabeza de la columna Herramienta, antes de los dos puntos)
HERRAMIENTA_CABEZA = {
    "izaje_plan": "GCC + CdP Maniobras de Izaje + ASP",
    "izaje_segregacion": "GCC + CdP Maniobras de Izaje + Desafío en Terreno",
    "izaje_certificacion": "GCC + CdP Maniobras de Izaje + ASP",
    "preuso_equipos": "EdC Vehículo Liviano + Inspección de Equipos + ASP",
    "andamios_plataformas": "Inspección de Plataformas + CdP Trabajos en Altura + ASP",
    "altura_caidas": "ASP + CdP Trabajos en Altura + Inspección de Equipos",
    "tableros_electricos": "Desafío en Terreno + CdP Energía + ASP",
    "extensiones_electricas": "Desafío en Terreno + CdP Energía + ASP",
    "equipos_emergencia": "Inspección de Equipos + CdP Emergencia + PASSO",
}

# patrones extra para la detección de exposición (revalorización)
REVAL_EXPOSICIONES = {
    "linea_fuego_equipo": [
        r"a (escasos|pocos|solo)? ?\d+([.,]\d+)? ?(m|mt|mts|metro|metros) (del?|al) (equipo|vehiculo|camion|camioneta|excavadora|cargador|retroexcavadora|maquina|grua)",
    ],
}
