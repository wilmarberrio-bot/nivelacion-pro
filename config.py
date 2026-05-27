"""
config.py - Configuracion central de Nivelacion Pro Web
Sin dependencias de openpyxl. Solo constantes operativas y variables de entorno.
"""
import os
from datetime import datetime

try:
    import pytz
    TZ_BOGOTA = pytz.timezone("America/Bogota")
    def now_bogota():
        return datetime.now(TZ_BOGOTA)
except ImportError:
    def now_bogota():
        return datetime.now()

# =============================================================
# METABASE - Variables de entorno (NUNCA hardcodear)
# =============================================================
METABASE_URL      = os.environ.get("METABASE_URL", "")
METABASE_USER     = os.environ.get("METABASE_USER", "")
METABASE_PASSWORD = os.environ.get("METABASE_PASSWORD", "")
METABASE_CARD_ID  = int(os.environ.get("METABASE_CARD_ID", "0"))
METABASE_API_KEY  = os.environ.get("METABASE_API_KEY", "")

# Columnas esperadas en el resultado de Metabase
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

# =============================================================
# FRANJAS HORARIAS OPERATIVAS
# =============================================================
FRANJAS = [
    "08:00-09:30",
    "10:00-11:30",
    "13:00-14:30",
    "14:30-16:00",
]

# =============================================================
# ESTADOS - Clasificacion operativa
# =============================================================
MOVABLE_STATUSES = [
    "por programar",
    "programado",
    "programada",
]

BLOCKED_STATUSES = [
    "en camino",
    "en sitio",
    "iniciado",
    "iniciada",
    "finalizado",
    "completado",
    "dispositivos subidos",
    "cancelado",
    "cancelado cliente",
    "cancelado operativo",
    "no ejecutado",
    "reprogramado",
    "reagendado",
]

FINALIZED_STATUSES = [
    "finalizado",
    "completado",
    "dispositivos subidos",
]

CANCELLED_STATUSES = [
    "cancelado",
    "cancelado cliente",
    "cancelado operativo",
    "no ejecutado",
]

RESCHEDULED_STATUSES = [
    "reprogramado",
    "reagendado",
]

IN_PROGRESS_STATUSES = [
    "en camino",
    "en sitio",
    "iniciado",
    "iniciada",
]

# Progress codes (para alertas)
# 0=sin iniciar, 1=en camino, 2=en sitio, 3=iniciado,
# 4=trabajando, 5=terminando, 6=finalizado
PROGRESS_FINALIZED = 6

# =============================================================
# UMBRALES OPERATIVOS
# =============================================================
MIN_IDEAL_LOAD          = 3     # Carga ideal minima por tecnico
MAX_IDEAL_LOAD          = 5     # Carga ideal maxima por tecnico
MAX_ABSOLUTE_LOAD       = 6     # Maximo absoluto (Lun-Vie)
MAX_ORDERS_PER_SLOT     = 2     # Maximo de ordenes en una misma franja
MAX_DUPLICATED_SLOTS    = 1     # Maximo de franjas duplicadas permitidas
MIN_IMBALANCE_TO_MOVE   = 2     # Diferencia minima de carga para proponer movimiento
ORDER_DURATION_HOURS    = 1.0   # Duracion estimada por orden (horas)
MAX_ORDER_DURATION_HOURS = 1.5
MAX_ALLOWED_DISTANCE_KM = 8.0
MAX_SUBZONES_SOFT       = 3

# Alertas de tiempo (minutos)
ONSITE_ALERT_MINUTES              = int(os.environ.get("ONSITE_ALERT_MINUTES",   "30"))
INICIADO_ALERT_MINUTES            = int(os.environ.get("INICIADO_ALERT_MINUTES", "90"))
ACTIVE_SLOT_NO_PROGRESS_MINUTES   = int(os.environ.get("ACTIVE_SLOT_NO_PROGRESS_MINUTES", "45"))
SLOT_RISK_MINUTES_BEFORE_END      = int(os.environ.get("SLOT_RISK_MINUTES_BEFORE_END",    "30"))

# Tecnico con N+ ordenes en una franja = sobrecarga
OVERLOAD_PER_SLOT       = 2

# =============================================================
# SCORING
# =============================================================
FRAGMENTATION_PENALTY              = 900
INTERZONE_DISTANCE_PENALTY         = 1500
ZONE_ONLY_NO_COORDS_PENALTY        = 50000
EFFICIENT_TECH_PROTECTION_SCORE    = 0.85
MIN_ROUTE_SAVINGS_KM               = 1.0
MIN_ROUTE_SAVINGS_PCT              = 0.30
MIN_ROUTE_SCORE_BENEFIT            = 350
NEARBY_BUILDING_RADIUS_KM          = 0.25
MAX_SWAP_DISTANCE_INCREASE_KM      = 2.0
MIN_SAVED_KM_FOR_SWAP              = 0.5
MAX_INTERZONE_ASSIGNMENTS_PER_TECH = 1

# =============================================================
# ZONAS ADYACENTES (Area Metropolitana de Medellin)
# =============================================================
ZONE_ADJACENCY = {
    "MEDELLIN":   ["BELLO", "ENVIGADO", "ITAGUI", "SABANETA"],
    "BELLO":      ["MEDELLIN"],
    "ENVIGADO":   ["MEDELLIN", "ITAGUI", "SABANETA"],
    "ITAGUI":     ["MEDELLIN", "ENVIGADO", "SABANETA", "LA ESTRELLA"],
    "SABANETA":   ["ITAGUI", "ENVIGADO", "LA ESTRELLA"],
    "LA ESTRELLA":["ITAGUI", "SABANETA", "CALDAS"],
    "CALDAS":     ["LA ESTRELLA", "SABANETA"],
    "RIONEGRO":   [],
}

# =============================================================
# CACHE EN MEMORIA (TTL en segundos)
# =============================================================
DATA_CACHE_TTL = int(os.environ.get("DATA_CACHE_TTL", "300"))  # 5 minutos

# =============================================================
# GOOGLE SHEETS - Export diario
# URL del Web App del Apps Script (configurar en Render como variable de entorno)
# =============================================================
SHEETS_WEBAPP_URL = os.environ.get("SHEETS_WEBAPP_URL", "")
