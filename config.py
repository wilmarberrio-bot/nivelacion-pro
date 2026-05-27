"""
config.py — Configuración central de Nivelación Pro Web
Sin dependencias de openpyxl. Solo constantes operativas y variables de entorno.
"""
import os
from datetime import datetime, timedelta

try:
    import pytz
    TZ_BOGOTA = pytz.timezone("America/Bogota")
    def now_bogota():
        return datetime.now(TZ_BOGOTA)
    def tomorrow_bogota():
        return datetime.now(TZ_BOGOTA) + timedelta(days=1)
except ImportError:
    def now_bogota():
        return datetime.now()
    def tomorrow_bogota():
        return datetime.now() + timedelta(days=1)

# ─────────────────────────────────────────────
# METABASE — Variables de entorno (NUNCA hardcodear)
# ─────────────────────────────────────────────
METABASE_URL      = os.environ.get("METABASE_URL", "")
METABASE_USER     = os.environ.get("METABASE_USER", "")
METABASE_PASSWORD = os.environ.get("METABASE_PASSWORD", "")
METABASE_CARD_ID  = int(os.environ.get("METABASE_CARD_ID", "0"))
METABASE_API_KEY  = os.environ.get("METABASE_API_KEY", "")

COL_ORDER_ID   = os.environ.get("COL_ORDER_ID",   "id_orden")
COL_TECH       = os.environ.get("COL_TECH",        "tecnico")
COL_STATUS     = os.environ.get("COL_STATUS",      "estado")
COL_FRANJA     = os.environ.get("COL_FRANJA",      "franja")
COL_TIPO       = os.environ.get("COL_TIPO",        "tipo_trabajo")
COL_ZONE       = os.environ.get("COL_ZONE",        "zona")
COL_SUBZONE    = os.environ.get("COL_SUBZONE",     "subzona")
COL_ADDRESS    = os.environ.get("COL_ADDRESS",     "direccion")
COL_GMAPS      = os.environ.get("COL_GMAPS",       "google_maps")
COL_LAT        = os.environ.get("COL_LAT",         "latitud")
COL_LON        = os.environ.get("COL_LON",         "longitud")
COL_UPDATED_AT = os.environ.get("COL_UPDATED_AT",  "updated_at")

# ─────────────────────────────────────────────
# FRANJAS HORARIAS OPERATIVAS
# ─────────────────────────────────────────────
FRANJAS = [
    "08:00-09:30",
    "10:00-11:30",
    "13:00-14:30",
    "14:30-16:00",
]

# ─────────────────────────────────────────────
# ESTADOS — Clasificación operativa
# ─────────────────────────────────────────────
MOVABLE_STATUSES = [
    "por programar",
    "programado",
    "programada",
]

BLOCKED_STATUSES = [
    "en camino",
    "en sitio",
    "mac enviada",
    "mac principal enviada",
    "dispositivos subidos",
    "dispositivos cargados",
    "por auditar",
    "finalizado",
    "finalizada",
    "cerrado",
    "cerrada",
    "cancelado",
    "cancelada",
]

NEAR_FINISH_STATUSES = [
    "mac enviada",
    "mac principal enviada",
    "dispositivos subidos",
    "dispositivos cargados",
]

FINALIZED_STATUSES = [
    "finalizado",
    "finalizada",
    "por auditar",
    "cerrado",
    "cerrada",
]

STATUS_PROGRESS = {
    "por programar":         0,
    "programado":            0,
    "programada":            0,
    "en camino":             1,
    "en sitio":              2,
    "iniciado":              3,
    "iniciada":              3,
    "mac enviada":           4,
    "mac principal enviada": 4,
    "dispositivos subidos":  5,
    "dispositivos cargados": 5,
    "por auditar":           6,
    "finalizado":            7,
    "finalizada":            7,
}

# ─────────────────────────────────────────────
# UMBRALES OPERATIVOS
# ─────────────────────────────────────────────
MAX_IDEAL_LOAD          = 5
MIN_IDEAL_LOAD          = 4
MAX_ABSOLUTE_LOAD       = 6
MAX_ORDERS_PER_SLOT     = 2
MAX_DUPLICATED_SLOTS    = 1
MIN_IMBALANCE_TO_MOVE   = 2
ORDER_DURATION_HOURS    = 1.0
MAX_ORDER_DURATION_HOURS = 1.5
MAX_ALLOWED_DISTANCE_KM = 8.0
MAX_SUBZONES_SOFT       = 3

# Alerta "En sitio" sin iniciar: máximo 30 min
# Técnico llegó pero no ha marcado "Iniciado" — posible bloqueo de acceso o cliente ausente.
ONSITE_ALERT_MINUTES    = int(os.environ.get("ONSITE_ALERT_MINUTES", "30"))

# Alerta "Iniciado/a" en ejecución prolongada: máximo 90 min
# Orden en ejecución activa que supera 90 min = problema técnico o sin reporte.
INICIADO_ALERT_MINUTES  = int(os.environ.get("INICIADO_ALERT_MINUTES", "90"))

# Alerta: franja iniciada hace N+ minutos y el técnico no ha marcado ninguna orden
ACTIVE_SLOT_NO_PROGRESS_MINUTES = int(os.environ.get("ACTIVE_SLOT_NO_PROGRESS_MINUTES", "30"))

# Alerta preventiva: minutos antes del FIN de franja con 0 marcaciones
SLOT_RISK_MINUTES_BEFORE_END = int(os.environ.get("SLOT_RISK_MINUTES_BEFORE_END", "30"))

# Técnico con N+ órdenes en una franja = sobrecarga
OVERLOAD_PER_SLOT       = 2

# ─────────────────────────────────────────────
# SCORING
# ─────────────────────────────────────────────
FRAGMENTATION_PENALTY            = 900
INTERZONE_DISTANCE_PENALTY       = 1500
ZONE_ONLY_NO_COORDS_PENALTY      = 50000
EFFICIENT_TECH_PROTECTION_SCORE  = 0.85
MIN_ROUTE_SAVINGS_KM             = 1.0
MIN_ROUTE_SAVINGS_PCT            = 0.30
MIN_ROUTE_SCORE_BENEFIT          = 350
NEARBY_BUILDING_RADIUS_KM        = 0.25
MAX_SWAP_DISTANCE_INCREASE_KM    = 2.0
MIN_SAVED_KM_FOR_SWAP            = 0.5
MAX_INTERZONE_ASSIGNMENTS_PER_TECH = 1

# ─────────────────────────────────────────────
# ZONAS ADYACENTES (Área Metropolitana de Medellín)
# ─────────────────────────────────────────────
ZONE_ADJACENCY = {
    "MEDELLIN":    ["BELLO", "ENVIGADO", "ITAGUI", "SABANETA"],
    "BELLO":       ["MEDELLIN"],
    "ENVIGADO":    ["MEDELLIN", "SABANETA", "ITAGUI"],
    "ITAGUI":      ["MEDELLIN", "ENVIGADO", "SABANETA", "LA ESTRELLA"],
    "SABANETA":    ["ENVIGADO", "ITAGUI", "LA ESTRELLA", "CALDAS", "MEDELLIN"],
    "LA ESTRELLA": ["ITAGUI", "SABANETA", "CALDAS"],
    "CALDAS":      ["LA ESTRELLA", "SABANETA"],
    "RIONEGRO":    [],
}

# ─────────────────────────────────────────────
# CACHE EN MEMORIA (TTL en segundos)
# ─────────────────────────────────────────────
DATA_CACHE_TTL = int(os.environ.get("DATA_CACHE_TTL", "300"))
