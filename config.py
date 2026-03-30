from datetime import datetime

try:
    import pytz
    TZ_BOGOTA = pytz.timezone("America/Bogota")

    def now_bogota():
        return datetime.now(TZ_BOGOTA)
except ImportError:
    def now_bogota():
        return datetime.now()

# =========================
# CONFIGURACION BASE (LV)
# =========================
# =========================
# CONFIGURACION BASE (LV)
# =========================
MAX_IDEAL_LOAD = 5         # Carga ideal maxima por tecnico (LV)
MAX_ABSOLUTE_LOAD = 6      # ✅ Confirmado: Lunes-Viernes
MAX_ORDERS_PER_SLOT = 2    # Permitir solape como ultimo recurso
MAX_DUPLICATED_SLOTS = 1   # Permitir maximo 1 franja duplicada (hard)
MIN_IMBALANCE_TO_MOVE = 2
MIN_ROUTE_SAVINGS_KM = 1.0
MIN_ROUTE_SCORE_BENEFIT = 350
MIN_ROUTE_SAVINGS_PCT = 0.30
EFFICIENT_TECH_PROTECTION_SCORE = 0.85
MAX_SUBZONES_SOFT = 3
FRAGMENTATION_PENALTY = 900
ORDER_DURATION_HOURS = 1.0
MAX_ORDER_DURATION_HOURS = 1.5
MAX_ALLOWED_DISTANCE_KM = 8.0
MAX_INTERZONE_ASSIGNMENTS_PER_TECH = 1
ZONE_ONLY_NO_COORDS_PENALTY = 50000
INTERZONE_DISTANCE_PENALTY = 1500

# Optimización por edificio / swaps
NEARBY_BUILDING_RADIUS_KM = 0.25            # 250m cuenta como misma unidad si hay coords
MAX_SWAP_DISTANCE_INCREASE_KM = 2.0         # el técnico que recibe el swap no debe empeorar más de 2km
MIN_SAVED_KM_FOR_SWAP = 0.5                 # ahorro mínimo para sugerir swaps en misma zona


ZONE_ADJACENCY = {
    'MEDELLIN': ['BELLO', 'ENVIGADO', 'ITAGUI', 'SABANETA'],
    'BELLO': ['MEDELLIN'],
    'ENVIGADO': ['MEDELLIN', 'SABANETA', 'ITAGUI'],
    'ITAGUI': ['MEDELLIN', 'ENVIGADO', 'SABANETA', 'LA ESTRELLA'],
    'SABANETA': ['ENVIGADO', 'ITAGUI', 'LA ESTRELLA', 'CALDAS', 'MEDELLIN'],
    'LA ESTRELLA': ['ITAGUI', 'SABANETA', 'CALDAS'],
    'CALDAS': ['LA ESTRELLA', 'SABANETA'],
    'RIONEGRO': [],
}

MOVABLE_STATUSES = ['programado', 'programada', 'por programar']
FINALIZED_STATUSES = [
    'finalizado', 'finalizada', 'por auditar', 'cancelado', 'cancelada',
    'cerrado', 'cerrada', 'completado', 'completada'
]

STATUS_PROGRESS = {
    'programado': 0,
    'programada': 0,
    'inbound': 1,
    'en sitio': 2,
    'iniciado': 3,
    'iniciada': 3,
    'mac principal enviada': 4,
    'dispositivos cargados': 5,
}

NEAR_FINISH_STATUSES = ['dispositivos cargados', 'mac principal enviada']


# =========================
# Estilos Excel
# =========================
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
ALERT_FILL = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
SUCCESS_FILL = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
WARN_FILL = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

