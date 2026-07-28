"""
config.py - Configuracion central de Nivelacion Pro Web
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

FRANJAS = [
    "08:00-09:30",
    "10:00-11:30",
    "13:00-14:30",
    "14:30-16:00",
    "16:00-17:30",   # T2 exclusiva
]

MOVABLE_STATUSES = [
    "por programar",
    "programado",
    "programada",
    "inbound",
]

BLOCKED_STATUSES = [
    "en camino",
    "en sitio",
    "iniciado",
    "iniciada",
    "finalizado",
    "completado",
    "dispositivos subidos",
    "dispositivos cargados",
    "por auditar",
    "mac principal enviada",
    "sin cargar dispositivos",
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
    "dispositivos cargados",   # igual a dispositivos subidos
    "por auditar",             # orden ejecutada, en revisión de calidad
    "mac principal enviada",   # paso final FTTH — instalación completa
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

NEAR_FINISH_STATUSES = [
    "iniciado",
    "iniciada",
    "trabajando",
    "terminando",
    "sin cargar dispositivos",  # terminó visita, falta cargar equipo
]

STATUS_PROGRESS = {
    "por programar":        0,
    "programado":           0,
    "programada":           0,
    "en camino":            1,
    "en sitio":             2,
    "iniciado":             3,
    "iniciada":             3,
    "trabajando":           4,
    "terminando":           5,
    "finalizado":           6,
    "completado":           6,
    "dispositivos subidos": 6,
    "cancelado":            6,
    "cancelado cliente":    6,
    "cancelado operativo":  6,
    "no ejecutado":         6,
    "reprogramado":         6,
    "reagendado":           6,
    "dispositivos cargados":6,
    "por auditar":          6,
    "mac principal enviada":6,
    "sin cargar dispositivos":5,
}

PROGRESS_FINALIZED = 6

# ── Carga operativa ──────────────────────────────────────────────────
MIN_IDEAL_LOAD               = 4   # mínimo operativo real (antes 3)
MAX_IDEAL_LOAD               = 5
MAX_ABSOLUTE_LOAD            = 6
MAX_ORDERS_PER_SLOT          = 2
MAX_DUPLICATED_SLOTS         = 1
MIN_IMBALANCE_TO_MOVE        = 2
ORDER_DURATION_HOURS         = 1.0
MAX_ORDER_DURATION_HOURS     = 1.5
MAX_ALLOWED_DISTANCE_KM      = 2.5  # cap geográfico primario (antes 8.0)
MAX_DIST_EXCEPTION_KM        = 4.0  # excepción justificada (geo aislado)
MAX_SUBZONES_SOFT            = 3

# ── Turnos T1 / T2 ──────────────────────────────────────────────────
T1_START_HOUR          = 7.5    # 07:30
T1_END_HOUR            = 15.5   # 15:30
T2_START_HOUR          = 10.0   # 10:00
T2_END_HOUR            = 18.0   # 18:00
# T1 y T2 almuerzan igual: 11:30 a 12:30
LUNCH_START_T1         = 11.5   # T1 almuerza 11:30
LUNCH_END_T1           = 12.5   # T1 sale de almuerzo 12:30
LUNCH_START_T2         = 11.5   # T2 almuerza 11:30
LUNCH_END_T2           = 12.5   # T2 sale de almuerzo 12:30
LUNCH_START            = 11.5   # alias genérico
LUNCH_END              = 12.5
T2_MAX_ORDERS_10H_SLOT = 1      # T2: máx 1 orden en 10:00-11:30 (almuerza 11:30)
ALCANZADO_BUFFER_HOURS = 0.5    # margen antes de marcar ALCANZADO

T1_FRANJAS = [
    "08:00-09:30",
    "10:00-11:30",
    "13:00-14:30",
    "14:30-16:00",
]
T2_FRANJAS = [
    "10:00-11:30",
    "13:00-14:30",
    "14:30-16:00",
    "16:00-17:30",
]

# Ruta al CSV maestro de turnos dentro del repo
TURNOS_CSV_PATH = os.path.join(os.path.dirname(__file__), "data", "turnos.csv")

ONSITE_ALERT_MINUTES             = int(os.environ.get("ONSITE_ALERT_MINUTES",            "30"))
INICIADO_ALERT_MINUTES           = int(os.environ.get("INICIADO_ALERT_MINUTES",          "90"))
ACTIVE_SLOT_NO_PROGRESS_MINUTES  = int(os.environ.get("ACTIVE_SLOT_NO_PROGRESS_MINUTES", "45"))
SLOT_RISK_MINUTES_BEFORE_END     = int(os.environ.get("SLOT_RISK_MINUTES_BEFORE_END",    "30"))

OVERLOAD_PER_SLOT = 2

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

# Ponderaciones de score para sugerencias
GEO_BONUS_0_5KM    = 1500   # < 0.5 km — misma cuadra prácticamente
GEO_BONUS_1KM      = 900    # < 1 km
GEO_BONUS_2KM      = 400    # < 2 km
GEO_PENALTY_OVER   = 5000   # > MAX_ALLOWED_DISTANCE_KM
FRANJA_DUP_PENALTY = 800    # ya tiene orden en esa franja (era bloqueo duro)

ZONE_ADJACENCY = {
    "MEDELLIN":    ["BELLO", "ENVIGADO", "ITAGUI", "SABANETA"],
    "BELLO":       ["MEDELLIN"],
    "ENVIGADO":    ["MEDELLIN", "ITAGUI", "SABANETA"],
    "ITAGUI":      ["MEDELLIN", "ENVIGADO", "SABANETA", "LA ESTRELLA"],
    "SABANETA":    ["ITAGUI", "ENVIGADO", "LA ESTRELLA"],
    "LA ESTRELLA": ["ITAGUI", "SABANETA", "CALDAS"],
    "CALDAS":      ["LA ESTRELLA", "SABANETA"],
    "RIONEGRO":    [],
}

DATA_CACHE_TTL = int(os.environ.get("DATA_CACHE_TTL", "300"))

SHEETS_WEBAPP_URL = os.environ.get("SHEETS_WEBAPP_URL", "")
