"""
config.py - Configuracion central de Nivelacion Pro Web
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

# Metabase
METABASE_URL      = os.environ.get("METABASE_URL",      "https://metabase.somosinternet.com")
METABASE_USER     = os.environ.get("METABASE_USER",     "")
METABASE_PASSWORD = os.environ.get("METABASE_PASSWORD", "")
METABASE_CARD_ID  = int(os.environ.get("METABASE_CARD_ID", "26359"))
METABASE_API_KEY  = os.environ.get("METABASE_API_KEY",  "")

FRANJAS = ["08:00-09:30","10:00-11:30","13:00-14:30","14:30-16:00"]

MOVABLE_STATUSES = ["por programar","programado","programada"]
BLOCKED_STATUSES = ["en camino","en sitio","mac enviada","mac principal enviada",
    "dispositivos subidos","dispositivos cargados","por auditar",
    "finalizado","finalizada","cerrado","cerrada","cancelado","cancelada"]
NEAR_FINISH_STATUSES = ["mac enviada","mac principal enviada","dispositivos subidos","dispositivos cargados"]
FINALIZED_STATUSES   = ["finalizado","finalizada","por auditar","cerrado","cerrada"]
STATUS_PROGRESS = {
    "por programar":0,"programado":0,"programada":0,
    "en camino":1,"en sitio":2,"iniciado":3,"iniciada":3,
    "mac enviada":4,"mac principal enviada":4,
    "dispositivos subidos":5,"dispositivos cargados":5,
    "por auditar":6,"finalizado":7,"finalizada":7,
}

# Umbrales operativos
# Objetivo: cada tecnico debe finalizar entre MIN_IDEAL_LOAD y MAX_IDEAL_LOAD ordenes por dia
MAX_IDEAL_LOAD           = 5     # Techo: tecnico con 5+ total NO recibe mas sin decision del coordinador
MIN_IDEAL_LOAD           = 4     # Piso: tecnico con menos de 4 es candidato prioritario para recibir
MAX_ABSOLUTE_LOAD        = 6     # Hard limit: nunca superar 6
MAX_ORDERS_PER_SLOT      = 2
MAX_DUPLICATED_SLOTS     = 1
MIN_IMBALANCE_TO_MOVE    = 2     # Diferencia minima de totales para sugerir movimiento
ORDER_DURATION_HOURS     = 1.0
MAX_ORDER_DURATION_HOURS = 1.5
MAX_ALLOWED_DISTANCE_KM  = 8.0
MAX_SUBZONES_SOFT        = 3
ONSITE_ALERT_MINUTES     = int(os.environ.get("ONSITE_ALERT_MINUTES", "90"))
OVERLOAD_PER_SLOT        = 2

FRAGMENTATION_PENALTY             = 900
INTERZONE_DISTANCE_PENALTY        = 1500
ZONE_ONLY_NO_COORDS_PENALTY       = 50000
EFFICIENT_TECH_PROTECTION_SCORE   = 0.85
MIN_ROUTE_SAVINGS_KM              = 1.0
MIN_ROUTE_SAVINGS_PCT             = 0.30
NEARBY_BUILDING_RADIUS_KM         = 0.25
MAX_INTERZONE_ASSIGNMENTS_PER_TECH = 1

ZONE_ADJACENCY = {
    "MEDELLIN":    ["BELLO","ENVIGADO","ITAGUI","SABANETA"],
    "BELLO":       ["MEDELLIN"],
    "ENVIGADO":    ["MEDELLIN","SABANETA","ITAGUI"],
    "ITAGUI":      ["MEDELLIN","ENVIGADO","SABANETA","LA ESTRELLA"],
    "SABANETA":    ["ENVIGADO","ITAGUI","LA ESTRELLA","CALDAS","MEDELLIN"],
    "LA ESTRELLA": ["ITAGUI","SABANETA","CALDAS"],
    "CALDAS":      ["LA ESTRELLA","SABANETA"],
    "RIONEGRO":    [],
}

DATA_CACHE_TTL = int(os.environ.get("DATA_CACHE_TTL", "300"))
