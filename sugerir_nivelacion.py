import math
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import glob
from datetime import datetime

try:
    import pytz
    TZ_BOGOTA = pytz.timezone('America/Bogota')
    def now_bogota():
        return datetime.now(TZ_BOGOTA)
except ImportError:
    def now_bogota():
        return datetime.now()


# =========================
# CONFIGURACION BASE (LV)
# =========================
MAX_IDEAL_LOAD = 5         # Carga ideal maxima por tecnico (LV)
MAX_ABSOLUTE_LOAD = 6      # ✅ Confirmado: Lunes-Viernes
MAX_ORDERS_PER_SLOT = 2    # Permitir solape como ultimo recurso
MAX_DUPLICATED_SLOTS = 1   # Permitir maximo 1 franja duplicada (hard)
MIN_IMBALANCE_TO_MOVE = 1
ORDER_DURATION_HOURS = 1.0
MAX_ORDER_DURATION_HOURS = 1.5
MAX_ALLOWED_DISTANCE_KM = 8.0

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


# =========================
# Utils
# =========================
def norm_text(x, default=""):
    if x is None:
        return default
    return str(x).strip()

def norm_zone(x):
    s = norm_text(x, "SIN_ZONA")
    if not s or s.lower() in ['none', 'nan', '']:
        return "SIN_ZONA"
    return s.strip().upper()

def norm_subzone(x):
    s = norm_text(x, "SIN_SUBZONA")
    if not s or s.lower() in ['none', 'nan', '']:
        return "SIN_SUBZONA"
    return s.strip().upper()

def normalize_address_key(addr: str) -> str:
    """
    Normaliza address para usarlo como "clave edificio".
    No es perfecto, pero reduce variabilidad.
    """
    if not addr:
        return ""
    a = addr.strip().upper()
    for ch in [",", ".", "#", "-", "_"]:
        a = a.replace(ch, " ")
    a = " ".join(a.split())
    return a[:80]

def build_address_key(address: str, subzona: str) -> str:
    base = normalize_address_key(address) if address else ""
    if base:
        return base
    return (subzona or "SIN_SUBZONA").upper()[:60]

def extract_coords_from_text(text):
    """
    Extrae lat/lon de:
      - google maps: .../@lat,lon,17z
      - query: ...?q=lat,lon
      - texto: "lat, lon"
    """
    if not text:
        return None, None
    t = str(text)

    m = re.search(r'@(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    m = re.search(r'[?&]q=(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    m = re.search(r'(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    return None, None

def is_same_unit(o1: dict, o2: dict) -> bool:
    """
    Determina si dos órdenes son de la misma unidad/edificio:
    - addr_key igual, o
    - coords disponibles y distancia <= NEARBY_BUILDING_RADIUS_KM
    """
    ak1 = o1.get('addr_key') or ""
    ak2 = o2.get('addr_key') or ""
    if ak1 and ak1 == ak2:
        return True
    if o1.get('lat', 0) and o2.get('lat', 0) and o1.get('lon', 0) and o2.get('lon', 0):
        try:
            d = haversine(o1['lat'], o1['lon'], o2['lat'], o2['lon'])
            return d <= NEARBY_BUILDING_RADIUS_KM
        except Exception:
            return False
    return False

def get_latest_preruta_file():
    files = glob.glob("*.xlsx")
    candidates = [
        f for f in files
        if ("pre_ruta" in f.lower() or "nivelacion" in f.lower())
        and not f.startswith("~$") and not f.startswith("sugerencias_")
    ]
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]

def get_col_index(headers, possible_names):
    if isinstance(possible_names, str):
        possible_names = [possible_names]
    headers_str = [str(h).strip().lower() if h else "" for h in headers]
    for name in possible_names:
        if name.lower() in headers_str:
            return headers_str.index(name.lower())
    return -1

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = (math.sin(dLat / 2) ** 2 +
         math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
         math.sin(dLon / 2) ** 2)
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def get_centroid(locations):
    if not locations:
        return (0, 0)
    avg_lat = sum(loc[0] for loc in locations) / len(locations)
    avg_lon = sum(loc[1] for loc in locations) / len(locations)
    return (avg_lat, avg_lon)

def parse_franja_hours(franja_str):
    if not franja_str or franja_str == 'Sin Franja':
        return None, None
    try:
        clean = str(franja_str).replace('\u2013', '-').replace('\u2014', '-').replace('\ufffd', '-')
        parts = clean.split('-')
        if len(parts) < 2:
            return None, None

        def parse_time(t):
            t = t.strip()
            for seg in t.split():
                if ':' in seg:
                    h, m = seg.strip().split(':')[:2]
                    return int(h) + int(m) / 60.0
            if ':' in t:
                h, m = t.split(':')[:2]
                return int(h) + int(m) / 60.0
            return None

        start = parse_time(parts[0])
        end = parse_time(parts[1])
        if start is None or end is None:
            return None, None
        return start, end
    except:
        return None, None

def is_status(status_lower, status_list):
    for s in status_list:
        if s in status_lower:
            return True
    return False

def get_status_progress(status):
    sl = str(status).lower()
    for key, val in STATUS_PROGRESS.items():
        if key in sl:
            return val
    return 0

def estimate_remaining_hours(status, onsite_hour=None, current_hour=None):
    progress = get_status_progress(status)
    base_duration = ORDER_DURATION_HOURS
    if progress == 3: base_duration = 0.6
    if progress == 4: base_duration = 0.3
    if progress >= 5: base_duration = 0.1

    if onsite_hour is not None and current_hour is not None:
        elapsed = current_hour - onsite_hour
        if elapsed > 0:
            remaining = max(0.1, base_duration - elapsed)
            return remaining
    return base_duration

def count_duplicated_slots(franja_counts):
    return sum(1 for count in franja_counts.values() if count >= 2)

def format_hour(h_decimal):
    if h_decimal is None: return "N/A"
    h = int(h_decimal)
    m = int((h_decimal - h) * 60)
    return f"{h:02d}:{m:02d}"

def coords_to_sector(lat, lon, subzona):
    if lat == 0 and lon == 0:
        return "Sin ubicacion"
    try:
        return f"{subzona} ({round(float(lat), 4)}, {round(float(lon), 4)})"
    except (ValueError, TypeError):
        return f"{subzona} (Err Coords)"

def style_header_row(ws, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = THIN_BORDER

def auto_fit_columns(ws):
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 100),
                                min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 55)


def can_tech_handle_franja(tech_franja_counts, tech_all_orders, order_franja, current_hour,
                           allow_same_unit_override=False, same_unit=False):
    """
    allow_same_unit_override + same_unit:
      - si ya excede franja duplicada permitida, pero es MISMA UNIDAD, se permite como sugerencia condicional
    """
    franja_start, franja_end = parse_franja_hours(order_franja)

    current_in_slot = tech_franja_counts.get(order_franja, 0)
    if current_in_slot >= MAX_ORDERS_PER_SLOT:
        return False, "Ya tiene 2 ordenes en esta franja"

    if current_in_slot >= 1:
        existing_dups = count_duplicated_slots(tech_franja_counts)
        if existing_dups >= MAX_DUPLICATED_SLOTS:
            if allow_same_unit_override and same_unit:
                return True, "OK (EXCEPCION: MISMA UNIDAD)"
            return False, "Ya tiene su franja duplicada permitida"

    tarde_count = 0
    if franja_start is not None and franja_start >= 14.5:
        for f, c in tech_franja_counts.items():
            fs, _ = parse_franja_hours(f)
            if fs is not None and fs >= 14.5:
                tarde_count += c
        if tarde_count >= 2:
            return False, "Ya tiene 2 ordenes en franjas de tarde (>=14:30)"

    if franja_start is not None:
        active_order = next((o for o in tech_all_orders if 1 <= get_status_progress(o['status']) < 6), None)
        rem_hours = 0
        if active_order:
            rem_hours = estimate_remaining_hours(active_order['status'], active_order.get('onsite_hour'), current_hour)

        prog_before = 0
        for o in tech_all_orders:
            if is_status(o['status'].lower(), MOVABLE_STATUSES):
                fs_h, _ = parse_franja_hours(o['franja'])
                if fs_h is not None and fs_h < franja_start:
                    prog_before += 1

        estimated_ready_hour = current_hour + rem_hours + (prog_before * ORDER_DURATION_HOURS)

        if franja_end is not None:
            is_already_late = current_hour > (franja_end - 0.5)
            if not is_already_late:
                if estimated_ready_hour > (franja_end + 0.5):
                    return False, f"No alcanza: listo ~{estimated_ready_hour:.1f}h, franja termina {franja_end:.1f}h"

    return True, "OK"


def estimate_arrival_for_franja(tech_all_orders, order_franja, current_hour):
    franja_start, franja_end = parse_franja_hours(order_franja)
    if franja_start is None:
        return None, None, None, None

    active_order = next((o for o in tech_all_orders if 1 <= get_status_progress(o['status']) < 6), None)
    rem_n = 0.0
    if active_order:
        rem_n = estimate_remaining_hours(active_order['status'], active_order.get('onsite_hour'), current_hour)
    rem_m = rem_n * (MAX_ORDER_DURATION_HOURS / ORDER_DURATION_HOURS) if ORDER_DURATION_HOURS > 0 else rem_n

    prog_before = 0
    for o in tech_all_orders:
        if is_status(o['status'].lower(), MOVABLE_STATUSES):
            fs_h, _ = parse_franja_hours(o['franja'])
            if fs_h is not None and fs_h < franja_start:
                prog_before += 1

    ready_normal = current_hour + rem_n + (prog_before * ORDER_DURATION_HOURS)
    ready_max = current_hour + rem_m + (prog_before * MAX_ORDER_DURATION_HOURS)

    arrival_normal = max(ready_normal, franja_start)
    arrival_max = max(ready_max, franja_start)

    return arrival_normal, arrival_max, franja_start, franja_end


def generate_suggestions(input_file, forced_hour=None):
    _now = now_bogota()
    if forced_hour is not None:
        current_hour = forced_hour
    else:
        current_hour = _now.hour + _now.minute / 60.0

    print(f"Leyendo archivo: {input_file}")
    print(f"Hora actual (COL): {current_hour:.2f} ({_now.strftime('%H:%M')} America/Bogota)")

    try:
        wb = openpyxl.load_workbook(input_file, read_only=True, data_only=True)
        sheet = wb.active
    except Exception as e:
        return f"Error leyendo el archivo: {str(e)}", None

    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]

    idx_status = get_col_index(headers, ['status_txt', 'Estado', 'Estado de orden de trabajo'])
    idx_zone = get_col_index(headers, ['Cities__name', 'Ciudad', 'Zona'])
    idx_zone_fallback = get_col_index(headers, ['Zone Name'])
    idx_zone_op = get_col_index(headers, ['zona_op'])
    idx_subzone = get_col_index(headers, ['Subzone', 'Subzona', 'Sites'])
    idx_tech = get_col_index(headers, ['tecnico', 'technenvician', 'Tecnico asignado', 'Tecnico', 'Tech'])
    idx_order = get_col_index(headers, ['appointment_id', 'ID', 'Numero de orden', 'Order ID'])
    idx_franja = get_col_index(headers, ['franja_label', 'Franja', 'Cita', 'Ventana'])
    idx_lat = get_col_index(headers, ['Latitude', 'Latitud', 'lat'])
    idx_lon = get_col_index(headers, ['Longitude', 'Longitud', 'lon', 'lng'])

    # ✅ NUEVO: link de Google Maps
    idx_maps = get_col_index(headers, ['Google Maps Link', 'google maps link', 'maps link', 'map link', 'GoogleMapsLink'])

    idx_address = get_col_index(headers, ['Addresses__address', 'Direccion', 'direccion'])
    idx_onsite = get_col_index(headers, ['onsite_at_cot'])

    missing = []
    if idx_tech == -1: missing.append("Tecnico")
    if idx_order == -1: missing.append("ID Orden")
    if missing:
        return f"Error: No se encontraron las columnas: {', '.join(missing)}", None

    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        tech = norm_text(row[idx_tech], "SIN_ASIGNAR") if (idx_tech != -1) else "SIN_ASIGNAR"
        if not tech or tech.lower() in ['none', 'nan', '']:
            tech = "SIN_ASIGNAR"

        # Zona: prioridad Cities__name > Zone Name > zona_op
        zona = "SIN_ZONA"
        if idx_zone != -1 and row[idx_zone] and str(row[idx_zone]).strip().lower() not in ['none', '']:
            zona = str(row[idx_zone]).strip()
        elif idx_zone_fallback != -1 and row[idx_zone_fallback] and str(row[idx_zone_fallback]).strip().lower() not in ['none', '']:
            zona = str(row[idx_zone_fallback]).strip()
        elif idx_zone_op != -1 and row[idx_zone_op] and str(row[idx_zone_op]).strip().lower() not in ['none', '']:
            zona = "Zona " + str(row[idx_zone_op]).strip()

        subzona = "SIN_SUBZONA"
        if idx_subzone != -1 and row[idx_subzone] and str(row[idx_subzone]).strip().lower() not in ['none', '']:
            subzona = str(row[idx_subzone]).strip()

        zona = norm_zone(zona)
        subzona = norm_subzone(subzona)

        order_id = row[idx_order] if idx_order != -1 else "N/A"
        status = norm_text(row[idx_status], "Sin Estado") if (idx_status != -1) else "Sin Estado"
        franja = norm_text(row[idx_franja], "Sin Franja") if (idx_franja != -1) else "Sin Franja"
        franja = franja.replace('\ufffd', '-').replace('\u2013', '-').replace('\u2014', '-')

        lat, lon = 0.0, 0.0
        if idx_lat != -1 and row[idx_lat]:
            try: lat = float(row[idx_lat])
            except: pass
        if idx_lon != -1 and row[idx_lon]:
            try: lon = float(row[idx_lon])
            except: pass

        address = norm_text(row[idx_address], "") if (idx_address != -1) else ""
        maps_link = norm_text(row[idx_maps], "") if (idx_maps != -1) else ""

        # ✅ FALLBACK: si no hay coords, intentar extraerlas del link de Google Maps y luego del address
        if (lat == 0.0 or lon == 0.0):
            lat2, lon2 = extract_coords_from_text(maps_link)
            if lat2 is not None and lon2 is not None:
                lat, lon = lat2, lon2
            else:
                lat3, lon3 = extract_coords_from_text(address)
                if lat3 is not None and lon3 is not None:
                    lat, lon = lat3, lon3

        addr_key = build_address_key(address, subzona)

        onsite_dt = row[idx_onsite] if idx_onsite != -1 else None
        onsite_hour = None
        if onsite_dt and isinstance(onsite_dt, datetime):
            onsite_hour = onsite_dt.hour + onsite_dt.minute / 60.0

        data.append({
            'tech': tech, 'zona': zona, 'subzona': subzona,
            'order_id': order_id, 'status': status, 'franja': franja,
            'lat': lat, 'lon': lon, 'address': address, 'maps_link': maps_link, 'addr_key': addr_key,
            'sector': coords_to_sector(lat, lon, subzona),
            'onsite_hour': onsite_hour
        })

    wb.close()

    if not data:
        return "Error: No se encontraron datos en el archivo.", None

    print(f"Total ordenes: {len(data)}")

    zones = sorted(set(d['zona'] for d in data if d['zona'] != "SIN_ZONA"))

    unique_franjas = sorted(list(set(d['franja'] for d in data if d['franja'] != 'Sin Franja')))
    is_saturday_shift = (len(unique_franjas) <= 2) or (_now.weekday() == 5)

    # --- Detectar si ya pasó la primera franja (para apoyo inter-zona) ---
    first_f_end = None
    for fr in unique_franjas:
        fs, fe = parse_franja_hours(fr)
        if fe is not None:
            first_f_end = fe
            break
    if first_f_end is None:
        # fallback: si no se puede parsear, usamos mediodía como proxy
        after_first_franja = current_hour >= 12.0
    else:
        # apoyo proactivo se habilita cuando termina la primera franja (+15 min)
        after_first_franja = current_hour >= (first_f_end + 0.25)

    global MAX_IDEAL_LOAD, MAX_ABSOLUTE_LOAD, ORDER_DURATION_HOURS, MAX_ORDER_DURATION_HOURS
    if is_saturday_shift:
        print(f"MODO SABADO DETECTADO ({len(unique_franjas)} franjas). Ajustando limites: 3-4 ordenes.")
        MAX_IDEAL_LOAD = 3
        MAX_ABSOLUTE_LOAD = 4
        ORDER_DURATION_HOURS = 0.75
        MAX_ORDER_DURATION_HOURS = 1.1
    else:
        MAX_IDEAL_LOAD = 5
        MAX_ABSOLUTE_LOAD = 6   # ✅ LV
        ORDER_DURATION_HOURS = 1.0
        MAX_ORDER_DURATION_HOURS = 1.5

    suggestions = []
    alerts = []
    zone_summaries = []
    subzone_summaries = []

    # =========================
    # Cargas y mapas globales
    # =========================
    tech_total = {}
    tech_finalized = {}
    tech_pending = {}
    tech_movable = {}
    tech_all_orders = {}
    tech_locations = {}
    tech_franja_counts = {}
    tech_subzones = {}
    tech_has_near_finish = {}
    tech_main_zone = {}

    # Para optimización por edificio
    active_buildings = {}   # addr_key -> set(techs) con estado >=2 en esa unidad
    building_orders_programmed = {}  # addr_key -> list de programadas

    for d in data:
        t = d['tech']
        status_lower = d['status'].lower()

        # Programadas por unidad
        if d.get('addr_key') and is_status(status_lower, MOVABLE_STATUSES):
            building_orders_programmed.setdefault(d['addr_key'], []).append(d)

        if t == "SIN_ASIGNAR":
            if is_status(status_lower, MOVABLE_STATUSES):
                tech_movable.setdefault(t, []).append(d)
            continue

        tech_all_orders.setdefault(t, []).append(d)
        tech_total[t] = tech_total.get(t, 0) + 1

        if t not in tech_main_zone and d['zona'] != "SIN_ZONA":
            tech_main_zone[t] = d['zona']

        if is_status(status_lower, FINALIZED_STATUSES):
            tech_finalized[t] = tech_finalized.get(t, 0) + 1
        else:
            tech_pending[t] = tech_pending.get(t, 0) + 1

            tech_franja_counts.setdefault(t, {})
            tech_franja_counts[t][d['franja']] = tech_franja_counts[t].get(d['franja'], 0) + 1

            if d['lat'] != 0 and d['lon'] != 0:
                tech_locations.setdefault(t, []).append((d['lat'], d['lon']))

            tech_subzones.setdefault(t, set()).add(d['subzona'])

            if is_status(status_lower, NEAR_FINISH_STATUSES):
                tech_has_near_finish[t] = True

        if is_status(status_lower, MOVABLE_STATUSES):
            tech_movable.setdefault(t, []).append(d)

        prog = get_status_progress(d['status'])
        if prog >= 2 and prog < 6 and d.get('addr_key'):
            active_buildings.setdefault(d['addr_key'], set()).add(t)

    donors_interzone_count = {}
    techs_moved_from_zone = set()

    # =========================
    # Analisis por zona (alertas + nivelación principal)
    # =========================
    for z in zones:
        zone_data = [d for d in data if d['zona'] == z]

        all_techs_in_zone = sorted(set(d['tech'] for d in zone_data if d['tech'] != "SIN_ASIGNAR"))
        active_techs_in_zone = [t for t in all_techs_in_zone if tech_pending.get(t, 0) > 0 or tech_finalized.get(t, 0) > 0]

        # ===== Alertas =====
        for t in active_techs_in_zone:
            total = tech_total.get(t, 0)

            if total > MAX_ABSOLUTE_LOAD:
                alerts.append({
                    'tipo': 'SOBRECARGA',
                    'zona': z, 'tecnico': t,
                    'detalle': f"Tiene {total} ordenes totales (max: {MAX_ABSOLUTE_LOAD})"
                })

            all_tech_orders = tech_all_orders.get(t, [])
            active_conflicts = [o for o in all_tech_orders if 2 <= get_status_progress(o['status']) < 6]
            if len(active_conflicts) > 1:
                ids = [str(o['order_id']) for o in active_conflicts]
                alerts.append({
                    'tipo': 'MULTI-ESTADO ACTIVO',
                    'zona': z, 'tecnico': t,
                    'detalle': f"El técnico tiene {len(active_conflicts)} órdenes activas simultáneas: {', '.join(ids)}."
                })

            # Alertas riesgo franja
            active_list = [o for o in all_tech_orders if 1 <= get_status_progress(o['status']) < 6]
            movable_list = sorted(
                [o for o in all_tech_orders if is_status(o['status'].lower(), MOVABLE_STATUSES)],
                key=lambda x: parse_franja_hours(x['franja'])[0] if parse_franja_hours(x['franja'])[0] is not None else 0
            )

            proj_normal = current_hour
            proj_max = current_hour

            if active_list:
                o_act = active_list[0]
                rem_n = estimate_remaining_hours(o_act['status'], o_act.get('onsite_hour'), current_hour)
                rem_m = rem_n * (MAX_ORDER_DURATION_HOURS / ORDER_DURATION_HOURS) if ORDER_DURATION_HOURS > 0 else rem_n
                proj_normal += rem_n
                proj_max += rem_m

            for dmv in movable_list:
                f_start, f_end = parse_franja_hours(dmv['franja'])
                if f_start is None:
                    continue

                arrival_normal = max(proj_normal, f_start)
                arrival_max = max(proj_max, f_start)

                proj_normal = arrival_normal + ORDER_DURATION_HOURS
                proj_max = arrival_max + MAX_ORDER_DURATION_HOURS

                if f_end is not None and arrival_max > f_end:
                    alerts.append({
                        'tipo': 'FRANJA EN RIESGO',
                        'zona': z, 'tecnico': t,
                        'detalle': f"Orden {dmv['order_id']} (Franja {dmv['franja']}) - Arribo ~{format_hour(arrival_max)} (Límite {format_hour(f_end)})."
                    })

            # Duplicadas existentes
            if t in tech_franja_counts:
                duplicated = count_duplicated_slots(tech_franja_counts[t])
                if duplicated > MAX_DUPLICATED_SLOTS:
                    franjas_dup = [f"{f}: {c}" for f, c in tech_franja_counts[t].items() if c >= 2]
                    alerts.append({
                        'tipo': 'FRANJAS DUPLICADAS',
                        'zona': z, 'tecnico': t,
                        'detalle': f"{duplicated} franjas duplicadas: {', '.join(franjas_dup)}"
                    })

        # ===== Resumen zona =====
        active_techs_in_zone = [t for t, m_zone in tech_main_zone.items() if m_zone == z]
        all_techs_in_zone = active_techs_in_zone

        unassigned_orders = [
            d for d in zone_data if d['tech'] == "SIN_ASIGNAR"
            and is_status(d['status'].lower(), MOVABLE_STATUSES)
        ]

        initial_pending_total = sum(tech_pending.get(t, 0) for t in active_techs_in_zone)
        initial_finalized_total = sum(tech_finalized.get(t, 0) for t in active_techs_in_zone)
        zone_pends = [tech_pending.get(t, 0) for t in active_techs_in_zone]

        current_zone_summary = {
            'zona': z,
            'techs': len(active_techs_in_zone),
            'pendientes_inicial': initial_pending_total,
            'pendientes_final': initial_pending_total,
            'sin_asignar_inicial': len(unassigned_orders),
            'sin_asignar_final': len(unassigned_orders),
            'total_finalizadas': initial_finalized_total,
            'avg_inicial': round(float(initial_pending_total) / len(active_techs_in_zone), 1) if active_techs_in_zone else 0.0,
            'min_inicial': min(zone_pends) if zone_pends else 0,
            'max_inicial': max(zone_pends) if zone_pends else 0,
        }
        zone_summaries.append(current_zone_summary)

        # ===== Resumen por técnico (subzonas/estados) =====
        active_zone_pending_data = [d for d in zone_data if not is_status(d['status'].lower(), FINALIZED_STATUSES)]

        tech_subzone_map = {}
        for dd in active_zone_pending_data:
            t = dd['tech']
            sz = dd['subzona']
            st = dd['status']
            tech_subzone_map.setdefault(t, {}).setdefault(sz, {})
            tech_subzone_map[t][sz][st] = tech_subzone_map[t][sz].get(st, 0) + 1

        for t in sorted(tech_subzone_map.keys()):
            sz_map = tech_subzone_map[t]
            total_pending = sum(sum(v for v in sts.values()) for sts in sz_map.values())
            subzona_lines = []
            for sz in sorted(sz_map.keys()):
                sts = sz_map[sz]
                estados_str = "  /  ".join(f"{k}({v})" for k, v in sts.items())
                subzona_lines.append(f"{sz}: {estados_str}")
            subzone_summaries.append({
                'zona': z,
                'tecnico': t,
                'subzonas_detalle': "\n".join(subzona_lines),
                'num_subzonas': len(sz_map),
                'pendientes': total_pending,
                'finalizadas': tech_finalized.get(t, 0) if t != "SIN_ASIGNAR" else 0,
                'carga_total': tech_total.get(t, 0) if t != "SIN_ASIGNAR" else total_pending,
            })

        # =========================
        # PASO 5: NIVELACION por carga y riesgo (tu lógica, casi igual)
        # =========================
        donors = []

        unassigned_in_zone = [d for d in tech_movable.get("SIN_ASIGNAR", []) if d['zona'] == z]
        if unassigned_in_zone:
            donors.append("SIN_ASIGNAR")

        avg_zone_pending = (sum(tech_pending.get(t, 0) for t in active_techs_in_zone) / len(active_techs_in_zone)) if active_techs_in_zone else 0

        for t in sorted(active_techs_in_zone):
            if t == "SIN_ASIGNAR":
                continue

            if tech_total.get(t, 0) > MAX_ABSOLUTE_LOAD and t in tech_movable and tech_movable[t]:
                if t not in donors: donors.append(t)
                continue

            if (tech_pending.get(t, 0) > MAX_IDEAL_LOAD or tech_pending.get(t, 0) > (avg_zone_pending + 1.1)) \
               and t in tech_movable and tech_movable[t]:
                if t not in donors: donors.append(t)
                continue

        for t in active_techs_in_zone:
            if t == "SIN_ASIGNAR":
                continue
            for dmv in tech_movable.get(t, []):
                f_start, f_end = parse_franja_hours(dmv['franja'])
                if f_start is None:
                    continue
                active_order = next((o for o in tech_all_orders.get(t, []) if 1 <= get_status_progress(o['status']) < 6), None)
                rem_h = estimate_remaining_hours(active_order['status'], active_order.get('onsite_hour'), current_hour) if active_order else 0
                prog_before = sum(
                    1 for o in tech_all_orders.get(t, [])
                    if is_status(o['status'].lower(), MOVABLE_STATUSES)
                    and parse_franja_hours(o['franja'])[0] is not None
                    and parse_franja_hours(o['franja'])[0] < f_start
                )
                est_ready = current_hour + rem_h + (prog_before * ORDER_DURATION_HOURS)

                if f_end is not None and est_ready > (f_end + 0.1):
                    if t not in donors: donors.append(t)
                    break

        def has_zone_capacity(zone_name):
            techs_in_z = [t for t, mz in tech_main_zone.items() if mz == zone_name]
            active_now = sum(1 for t in techs_in_z if t not in techs_moved_from_zone)
            return active_now > 1

        for donor in donors:
            if donor == "SIN_ASIGNAR":
                donor_orders = [d for d in tech_movable.get("SIN_ASIGNAR", []) if d['zona'] == z]
            else:
                donor_orders = list(tech_movable.get(donor, []))

            if not donor_orders:
                continue

            donor_orders.sort(key=lambda x: parse_franja_hours(x['franja'])[0] or 99)

            if donor == "SIN_ASIGNAR":
                moves_limit = len(donor_orders)
            else:
                moves_limit = max(1, tech_total.get(donor, 0) - MAX_IDEAL_LOAD)

            moved_count = 0
            for order in list(donor_orders):
                if moved_count >= moves_limit:
                    break

                best_receiver = None
                best_score = float('inf')
                best_detail = ""

                for pass_num in [1, 2]:
                    if best_receiver:
                        break

                    if pass_num == 1:
                        recipients = [r for r in active_techs_in_zone if r != donor]
                    else:
                        if donors_interzone_count.get(donor, 0) >= 1:
                            break
                        allowed_neighbor_zones = ZONE_ADJACENCY.get(z.upper(), [])
                        recipients = [
                            t for t in tech_total
                            if t not in all_techs_in_zone
                            and t != donor and t != "SIN_ASIGNAR"
                            and tech_main_zone.get(t, "").upper() in allowed_neighbor_zones
                        ]

                    for r in recipients:
                        if tech_total.get(r, 0) >= MAX_ABSOLUTE_LOAD:
                            continue

                        if donor != "SIN_ASIGNAR":
                            donor_pends = tech_pending.get(donor, 0)
                            recv_pends = tech_pending.get(r, 0)
                            if (donor_pends - recv_pends) < 1:
                                continue

                        can_handle, _ = can_tech_handle_franja(
                            tech_franja_counts.get(r, {}),
                            tech_all_orders.get(r, []),
                            order['franja'],
                            current_hour
                        )
                        if not can_handle:
                            continue

                        is_it_interzone = (tech_main_zone.get(r) != z)
                        if is_it_interzone:
                            if donor in techs_moved_from_zone:
                                continue
                            if not has_zone_capacity(z):
                                continue

                        # SCORE (carga + distancia + subzona + tardanza)
                        score = 0
                        score += tech_pending.get(r, 0) * 800
                        score += tech_total.get(r, 0) * 400

                        dist = 0.0
                        if order['lat'] != 0 and r in tech_locations:
                            centroid = get_centroid(tech_locations[r])
                            dist = haversine(order['lat'], order['lon'], centroid[0], centroid[1])
                            if dist > MAX_ALLOWED_DISTANCE_KM:
                                continue
                            score += dist * 250

                        bonus_sub = 0
                        if order['subzona'] in tech_subzones.get(r, set()):
                            bonus_sub = -2200
                            score += bonus_sub

                        arr_n, arr_m, f_start, f_end = estimate_arrival_for_franja(
                            tech_all_orders.get(r, []), order['franja'], current_hour
                        )
                        late_pen = 0
                        if f_end is not None and arr_m is not None and arr_m > f_end:
                            late_hours = (arr_m - f_end)
                            late_pen = late_hours * 12000
                            score += late_pen
                        elif f_start is not None and arr_n is not None and arr_n > f_start:
                            score += (arr_n - f_start) * 1200

                        if score < best_score:
                            best_score = score
                            best_receiver = r
                            best_detail = f"score={score:.0f} | load={tech_pending.get(r,0)} | dist={dist:.2f} | bonus_sub={bonus_sub} | late_pen={late_pen:.0f}"

                    if pass_num == 2 and best_receiver:
                        donors_interzone_count[donor] = donors_interzone_count.get(donor, 0) + 1
                        techs_moved_from_zone.add(donor)

                if best_receiver:
                    dist_km = 0.0
                    if best_receiver in tech_locations and order['lat'] != 0:
                        c = get_centroid(tech_locations[best_receiver])
                        dist_km = haversine(order['lat'], order['lon'], c[0], c[1])

                    is_interzone = tech_main_zone.get(best_receiver) != z
                    suggestions.append({
                        'zona': f"{z} (AYUDA EXTERNA)" if is_interzone else z,
                        'subzona': order['subzona'],
                        'origen': donor,
                        'destino': best_receiver,
                        'order_id': order['order_id'],
                        'franja': order['franja'],
                        'address': order.get('address', ''),
                        'distancia_estimada': f"{dist_km:.2f} km",
                        'alerta': f"Inter-Zona ({tech_main_zone.get(best_receiver)})" if is_interzone else ("Nivelacion Carga" if donor != "SIN_ASIGNAR" else "Sin Asignar"),
                        'pendientes_origen': tech_pending.get(donor, 0),
                        'pendientes_destino': tech_pending.get(best_receiver, 0),
                        'justificacion': best_detail
                    })

                    if donor != "SIN_ASIGNAR":
                        tech_total[donor] -= 1
                        tech_pending[donor] = max(0, tech_pending.get(donor, 0) - 1)
                        current_zone_summary['pendientes_final'] -= 1
                    else:
                        current_zone_summary['sin_asignar_final'] -= 1
                        current_zone_summary['pendientes_final'] += 1

                    tech_total[best_receiver] = tech_total.get(best_receiver, 0) + 1
                    tech_pending[best_receiver] = tech_pending.get(best_receiver, 0) + 1

                    tech_franja_counts.setdefault(best_receiver, {})
                    tech_franja_counts[best_receiver][order['franja']] = tech_franja_counts[best_receiver].get(order['franja'], 0) + 1

                    tech_all_orders.setdefault(best_receiver, []).append(order)

                    donor_orders.remove(order)
                    tech_movable[donor].remove(order)
                    moved_count += 1

        # desbalance final
        if active_techs_in_zone:
            final_pends = [tech_pending.get(t, 0) for t in active_techs_in_zone]
            current_zone_summary['desbalance_final'] = max(final_pends) - min(final_pends)
        else:
            current_zone_summary['desbalance_final'] = 0


    # ==========================================================
    # PASO 5.9: APOYO PROACTIVO INTER-ZONA (post 1ra franja)
    # - Si una zona vecina está saturada y esta zona está liviana,
    #   sugerimos mover capacidad (técnico libre) a la zona vecina.
    # - Regla: el técnico a mover debe quedar SIN pendientes, o
    #   redistribuimos sus programadas a otros técnicos de su zona.
    # ==========================================================
    if after_first_franja:
        # --- construir mapa de técnicos por zona y avg de pendientes ---
        zone_techs = {}
        for t, zname in tech_main_zone.items():
            if not zname or t == "SIN_ASIGNAR":
                continue
            zone_techs.setdefault(zname, []).append(t)

        zone_avg_pending = {}
        for zname, techs in zone_techs.items():
            if not techs:
                zone_avg_pending[zname] = 0.0
            else:
                zone_avg_pending[zname] = sum(tech_pending.get(t, 0) for t in techs) / len(techs)

        # --- addr_key por técnico (para reasignar sin coords por misma unidad) ---
        tech_addr_keys = {}
        for tt, orders in tech_all_orders.items():
            keys = set()
            for o in orders:
                ak = o.get('addr_key')
                if ak:
                    keys.add(ak)
            tech_addr_keys[tt] = keys

        def tech_has_active_order(tname):
            # cualquier estado activo (inbound/en sitio/iniciado...) cuenta como NO movible de zona
            return any(1 <= get_status_progress(o['status']) < 6 for o in tech_all_orders.get(tname, []))

        def redistribute_to_free_tech(tname, zname):
            # Redistribuye órdenes programadas del técnico dentro de su zona para dejarlo sin pendientes.
            movable_list = [o for o in list(tech_movable.get(tname, [])) if o.get('zona') == zname]
            if not movable_list:
                return True

            recipients = [r for r in zone_techs.get(zname, []) if r != tname]
            if not recipients:
                return False

            def fr_start(od):
                fs, _ = parse_franja_hours(od.get('franja'))
                return fs if fs is not None else 99
            movable_list.sort(key=fr_start)

            for order in list(movable_list):
                best_r = None
                best_score = float('inf')
                best_reason = ""

                for r in recipients:
                    if tech_total.get(r, 0) >= MAX_ABSOLUTE_LOAD:
                        continue

                    can_ok, reason = can_tech_handle_franja(
                        tech_franja_counts.get(r, {}),
                        tech_all_orders.get(r, []),
                        order.get('franja', ''),
                        current_hour,
                        allow_same_unit_override=False,
                        same_unit=False
                    )
                    if not can_ok:
                        continue

                    dist = 0.0
                    if order.get('lat', 0) and order.get('lon', 0) and r in tech_locations:
                        c = get_centroid(tech_locations[r])
                        dist = haversine(order['lat'], order['lon'], c[0], c[1])
                        if dist > MAX_ALLOWED_DISTANCE_KM:
                            continue
                    else:
                        ak = order.get('addr_key') or ""
                        if not ak or ak not in tech_addr_keys.get(r, set()):
                            continue

                    score = tech_pending.get(r, 0) * 900 + tech_total.get(r, 0) * 400 + dist * 250
                    if order.get('subzona') in tech_subzones.get(r, set()):
                        score -= 1500

                    if score < best_score:
                        best_score = score
                        best_r = r
                        best_reason = reason

                if not best_r:
                    return False

                suggestions.append({
                    'zona': zname,
                    'subzona': order.get('subzona', ''),
                    'origen': tname,
                    'destino': best_r,
                    'order_id': order.get('order_id', ''),
                    'franja': order.get('franja', ''),
                    'address': order.get('address', ''),
                    'distancia_estimada': "N/A" if not (order.get('lat',0) and order.get('lon',0)) else "Ver distancia",
                    'alerta': "LIBERAR TECNICO (PRE-APOYO)",
                    'pendientes_origen': tech_pending.get(tname, 0),
                    'pendientes_destino': tech_pending.get(best_r, 0),
                    'justificacion': f"Redistribución para liberar técnico y apoyar otra zona. {best_reason}."
                })

                tech_total[tname] = max(0, tech_total.get(tname, 0) - 1)
                tech_pending[tname] = max(0, tech_pending.get(tname, 0) - 1)
                if tname in tech_franja_counts:
                    cnt = tech_franja_counts[tname].get(order.get('franja',''), 0)
                    if cnt <= 1:
                        tech_franja_counts[tname].pop(order.get('franja',''), None)
                    else:
                        tech_franja_counts[tname][order.get('franja','')] = cnt - 1

                tech_total[best_r] = tech_total.get(best_r, 0) + 1
                tech_pending[best_r] = tech_pending.get(best_r, 0) + 1
                tech_franja_counts.setdefault(best_r, {})
                tech_franja_counts[best_r][order.get('franja','')] = tech_franja_counts[best_r].get(order.get('franja',''), 0) + 1

                try:
                    tech_all_orders.get(tname, []).remove(order)
                except:
                    pass
                tech_all_orders.setdefault(best_r, []).append(order)

                if tname in tech_movable and order in tech_movable[tname]:
                    tech_movable[tname].remove(order)
                tech_movable.setdefault(best_r, []).append(order)

                order['tech'] = best_r

            return tech_pending.get(tname, 0) == 0

        # --- ejecutar apoyo proactivo ---
        for z_low in zones:
            low_avg = zone_avg_pending.get(z_low, 0.0)
            if low_avg > (MAX_IDEAL_LOAD - 2):
                continue

            neighbors = ZONE_ADJACENCY.get(z_low.upper(), [])
            if not neighbors:
                continue

            z_high = None
            best_diff = 0.0
            for nz in neighbors:
                high_avg = zone_avg_pending.get(nz, None)
                if high_avg is None:
                    continue
                diff = high_avg - low_avg
                if high_avg >= MAX_IDEAL_LOAD and diff >= 2.0 and diff > best_diff:
                    best_diff = diff
                    z_high = nz

            if not z_high:
                continue

            cand = [t for t in zone_techs.get(z_low, []) if t not in techs_moved_from_zone]
            if len(cand) <= 1:
                continue

            free_cand = [t for t in cand if tech_pending.get(t,0) == 0 and not tech_has_active_order(t)]
            donor = None
            donor_freed_by = "YA LIBRE"
            if free_cand:
                donor = free_cand[0]
            else:
                cand2 = [t for t in cand if not tech_has_active_order(t)]
                if cand2:
                    cand2.sort(key=lambda t: tech_pending.get(t,0))
                    for t_try in cand2[:3]:
                        if redistribute_to_free_tech(t_try, z_low):
                            donor = t_try
                            donor_freed_by = "REDISTRIBUCION"
                            break

            if not donor:
                continue

            target = None
            source_label = None
            pool_un = [d for d in tech_movable.get("SIN_ASIGNAR", []) if d.get('zona') == z_high]
            if pool_un:
                pool_un.sort(key=lambda x: (parse_franja_hours(x.get('franja'))[0] or 99))
                target = pool_un[0]
                source_label = "SIN_ASIGNAR"
            else:
                high_techs = zone_techs.get(z_high, [])
                high_techs.sort(key=lambda t: tech_pending.get(t,0), reverse=True)
                for ht in high_techs:
                    if ht == "SIN_ASIGNAR":
                        continue
                    if tech_pending.get(ht,0) < MAX_IDEAL_LOAD:
                        continue
                    cand_orders = [o for o in tech_movable.get(ht, []) if o.get('zona') == z_high]
                    if cand_orders:
                        cand_orders.sort(key=lambda x: (parse_franja_hours(x.get('franja'))[0] or 99))
                        target = cand_orders[0]
                        source_label = ht
                        break

            if target:
                can_ok, reason = can_tech_handle_franja(
                    tech_franja_counts.get(donor, {}),
                    tech_all_orders.get(donor, []),
                    target.get('franja',''),
                    current_hour,
                    allow_same_unit_override=False,
                    same_unit=False
                )
                if can_ok:
                    suggestions.append({
                        'zona': f"{z_high} (APOYO PROACTIVO)",
                        'subzona': target.get('subzona',''),
                        'origen': donor,
                        'destino': f"{source_label} (Zona {z_high})",
                        'order_id': target.get('order_id',''),
                        'franja': target.get('franja',''),
                        'address': target.get('address',''),
                        'distancia_estimada': "N/A" if not (target.get('lat',0) and target.get('lon',0)) else "Ver distancia",
                        'alerta': "APOYO PROACTIVO INTERZONA",
                        'pendientes_origen': tech_pending.get(donor, 0),
                        'pendientes_destino': round(zone_avg_pending.get(z_high,0.0), 1),
                        'justificacion': (
                            f"Post 1ra franja: {z_low} liviana (avg {low_avg:.1f}) y {z_high} saturada (avg {zone_avg_pending.get(z_high,0):.1f}). "
                            f"Técnico {donor} {donor_freed_by}. {reason}."
                        )
                    })
                    techs_moved_from_zone.add(donor)

                    if source_label == "SIN_ASIGNAR":
                        try:
                            tech_movable.get("SIN_ASIGNAR", []).remove(target)
                        except:
                            pass
                    else:
                        try:
                            tech_movable.get(source_label, []).remove(target)
                        except:
                            pass
                else:
                    suggestions.append({
                        'zona': f"{z_high} (APOYO PROACTIVO)",
                        'subzona': "",
                        'origen': donor,
                        'destino': f"CAPACIDAD STANDBY (Zona {z_high})",
                        'order_id': "N/A",
                        'franja': "Post 1ra franja",
                        'address': "",
                        'distancia_estimada': "",
                        'alerta': "APOYO PROACTIVO INTERZONA",
                        'pendientes_origen': tech_pending.get(donor, 0),
                        'pendientes_destino': round(zone_avg_pending.get(z_high,0.0), 1),
                        'justificacion': (
                            f"Post 1ra franja: {z_low} liviana (avg {low_avg:.1f}) y {z_high} saturada (avg {zone_avg_pending.get(z_high,0):.1f}). "
                            f"Técnico {donor} {donor_freed_by}. No se asigna orden específica: {reason}."
                        )
                    })
                    techs_moved_from_zone.add(donor)
            else:
                suggestions.append({
                    'zona': f"{z_high} (APOYO PROACTIVO)",
                    'subzona': "",
                    'origen': donor,
                    'destino': f"CAPACIDAD STANDBY (Zona {z_high})",
                    'order_id': "N/A",
                    'franja': "Post 1ra franja",
                    'address': "",
                    'distancia_estimada': "",
                    'alerta': "APOYO PROACTIVO INTERZONA",
                    'pendientes_origen': tech_pending.get(donor, 0),
                    'pendientes_destino': round(zone_avg_pending.get(z_high,0.0), 1),
                    'justificacion': (
                        f"Post 1ra franja: {z_low} liviana (avg {low_avg:.1f}) y {z_high} saturada (avg {zone_avg_pending.get(z_high,0):.1f}). "
                        f"Técnico {donor} {donor_freed_by}. Sin orden concreta; se propone apoyo si llega nueva."
                    )
                })
                techs_moved_from_zone.add(donor)
    # ==========================================================
    # PASO 6: OPTIMIZACIÓN DE RUTA / UNIDAD + PROXIMIDAD + SWAPS
    # ==========================================================
    all_techs_global = sorted([t for t in tech_total.keys() if t != "SIN_ASIGNAR"])
    PROXIMITY_GAIN_MIN_KM = 1.5

    # 6A: Reasignación por proximidad (intra-zona) - igual que antes
    for t1 in all_techs_global:
        if not tech_movable.get(t1):
            continue
        for t2 in all_techs_global:
            if t2 <= t1:
                continue
            if not tech_movable.get(t2):
                continue

            if tech_main_zone.get(t1) != tech_main_zone.get(t2):
                continue

            if abs(tech_pending.get(t1, 0) - tech_pending.get(t2, 0)) > 1:
                continue

            c1 = get_centroid(tech_locations.get(t1, [(0, 0)]))
            c2 = get_centroid(tech_locations.get(t2, [(0, 0)]))

            for o1 in list(tech_movable[t1]):
                # si no hay coords, solo sugerimos si es misma unidad (addr_key) con un activo (lo cubre 6.0 fuera)
                if o1['lat'] == 0:
                    continue

                d_o1_t1 = haversine(o1['lat'], o1['lon'], c1[0], c1[1])
                d_o1_t2 = haversine(o1['lat'], o1['lon'], c2[0], c2[1])
                gain = d_o1_t1 - d_o1_t2

                if gain >= PROXIMITY_GAIN_MIN_KM:
                    can_handle, _ = can_tech_handle_franja(
                        tech_franja_counts.get(t2, {}),
                        tech_all_orders.get(t2, []),
                        o1['franja'], current_hour
                    )
                    if not can_handle:
                        continue

                    suggestions.append({
                        'zona': tech_main_zone.get(t1, ''),
                        'subzona': o1['subzona'],
                        'origen': t1,
                        'destino': t2,
                        'order_id': o1['order_id'],
                        'franja': o1['franja'],
                        'address': o1.get('address', ''),
                        'distancia_estimada': f"{d_o1_t2:.2f} km (antes {d_o1_t1:.2f} km)",
                        'alerta': 'REASIGNACION PROXIMITY',
                        'pendientes_origen': tech_pending.get(t1, 0),
                        'pendientes_destino': tech_pending.get(t2, 0),
                        'justificacion': f"{t2} está {gain:.1f} km más cerca. Reasignar reduce desplazamiento sin afectar carga."
                    })

                    tech_movable[t1].remove(o1)
                    break

    # 6B: INTERCAMBIOS (SWAPS) - AHORA PERMITE FRANJAS DISTINTAS, respetando duplicada (salvo misma unidad)
    for t1 in all_techs_global:
        if not tech_movable.get(t1):
            continue
        for t2 in all_techs_global:
            if t2 <= t1 or not tech_movable.get(t2):
                continue

            z1 = tech_main_zone.get(t1, "SIN_ZONA").upper()
            z2 = tech_main_zone.get(t2, "SIN_ZONA").upper()

            if z1 != z2:
                allowed_neighbors = ZONE_ADJACENCY.get(z1, [])
                if z2 not in allowed_neighbors:
                    continue

            c1 = get_centroid(tech_locations.get(t1, [(0, 0)]))
            c2 = get_centroid(tech_locations.get(t2, [(0, 0)]))

            orders_t1 = list(tech_movable.get(t1, []))
            orders_t2 = list(tech_movable.get(t2, []))

            for o1 in orders_t1:
                for o2 in orders_t2:
                    # Sin coords y no misma unidad => no podemos justificar distancias
                    same_unit_flag = is_same_unit(o1, o2)
                    if (o1.get('lat', 0) == 0 or o2.get('lat', 0) == 0 or o1.get('lon', 0) == 0 or o2.get('lon', 0) == 0) and not same_unit_flag:
                        continue

                    # Validar que cada uno pueda recibir la orden del otro
                    can_t1, reason_t1 = can_tech_handle_franja(
                        tech_franja_counts.get(t1, {}),
                        tech_all_orders.get(t1, []),
                        o2['franja'],
                        current_hour,
                        allow_same_unit_override=True,
                        same_unit=same_unit_flag
                    )
                    if not can_t1:
                        continue

                    can_t2, reason_t2 = can_tech_handle_franja(
                        tech_franja_counts.get(t2, {}),
                        tech_all_orders.get(t2, []),
                        o1['franja'],
                        current_hour,
                        allow_same_unit_override=True,
                        same_unit=same_unit_flag
                    )
                    if not can_t2:
                        continue

                    # Distancias actuales vs swap (si hay coords)
                    if o1.get('lat', 0) and o2.get('lat', 0) and o1.get('lon', 0) and o2.get('lon', 0):
                        dist1_actual = haversine(o1['lat'], o1['lon'], c1[0], c1[1])
                        dist2_actual = haversine(o2['lat'], o2['lon'], c2[0], c2[1])

                        dist1_swap = haversine(o1['lat'], o1['lon'], c2[0], c2[1])
                        dist2_swap = haversine(o2['lat'], o2['lon'], c1[0], c1[1])

                        total_saved = (dist1_actual + dist2_actual) - (dist1_swap + dist2_swap)
                    else:
                        # sin coords pero misma unidad => no calculamos ahorro real
                        dist1_actual = dist2_actual = dist1_swap = dist2_swap = 0.0
                        total_saved = 0.0

                    is_same_zone = (z1 == z2)
                    threshold = MIN_SAVED_KM_FOR_SWAP if is_same_zone else 2.0

                    subzone_bonus = 0.0
                    if is_same_zone:
                        if (o1['subzona'] in tech_subzones.get(t2, set())) and (o2['subzona'] in tech_subzones.get(t1, set())):
                            subzone_bonus = 1.0

                    # Condición: ahorro suficiente O es misma unidad (para que torre decida)
                    allow_by_unit = same_unit_flag
                    if (total_saved + subzone_bonus) > threshold or allow_by_unit:
                        alert = "INTERCAMBIO"
                        if "EXCEPCION" in (reason_t1 + " " + reason_t2):
                            alert = "INTERCAMBIO (CONDICIONAL: MISMA UNIDAD)"
                        if allow_by_unit and (total_saved + subzone_bonus) <= threshold:
                            alert = "INTERCAMBIO (CONDICIONAL: MISMA UNIDAD)"

                        if o1.get('lat', 0) and o2.get('lat', 0):
                            dist_msg = f"Ahorro {total_saved:.1f} km" + (" + Subzona" if subzone_bonus > 0 else "")
                        else:
                            dist_msg = "MISMA UNIDAD (sin coords)"

                        suggestions.append({
                            'zona': z1 if is_same_zone else f"{z1} <-> {z2}",
                            'subzona': f"{o1['subzona']} <-> {o2['subzona']}",
                            'origen': t1,
                            'destino': t2,
                            'order_id': f"{o1['order_id']} -> {t2}  /  {o2['order_id']} -> {t1}",
                            'franja': f"{o1['franja']} / {o2['franja']}",
                            'address': f"{o1['order_id']} ({o1['franja']})  <->  {o2['order_id']} ({o2['franja']})",
                            'distancia_estimada': dist_msg,
                            'alerta': alert,
                            'pendientes_origen': tech_pending.get(t1, 0),
                            'pendientes_destino': tech_pending.get(t2, 0),
                            'justificacion': f"{dist_msg}. {reason_t1}; {reason_t2}. Torre de control valida si aplica."
                        })

                        # Remover para no repetir
                        if o1 in tech_movable.get(t1, []):
                            tech_movable[t1].remove(o1)
                        if o2 in tech_movable.get(t2, []):
                            tech_movable[t2].remove(o2)
                        break
                else:
                    continue
                break

    # =========================
    # EXPORT EXCEL (igual)
    # =========================
    wb_out = openpyxl.Workbook()

    ws_zona = wb_out.active
    ws_zona.title = "Resumen por Zona"
    zh = ['Zona', 'Tecnicos', 'Sin Asignar INI', 'Sin Asignar FIN', 'Pendientes INI', 'Pendientes FIN',
          'Finalizadas', 'Avg Ini', 'Desbalance Ini', 'Desbalance Fin']
    ws_zona.append(zh)
    style_header_row(ws_zona, len(zh))

    for s in zone_summaries:
        row_num = ws_zona.max_row + 1
        desb_ini = s['max_inicial'] - s['min_inicial']
        desb_fin = s.get('desbalance_final', 0)

        ws_zona.append([
            s['zona'], s['techs'], s['sin_asignar_inicial'], s['sin_asignar_final'],
            s['pendientes_inicial'], s['pendientes_final'],
            s['total_finalizadas'], s['avg_inicial'], desb_ini, desb_fin
        ])

        if desb_ini >= 3: ws_zona.cell(row=row_num, column=9).fill = ALERT_FILL
        elif desb_ini >= 2: ws_zona.cell(row=row_num, column=9).fill = WARN_FILL

        if desb_fin >= 3: ws_zona.cell(row=row_num, column=10).fill = ALERT_FILL
        elif desb_fin >= 2: ws_zona.cell(row=row_num, column=10).fill = WARN_FILL

    auto_fit_columns(ws_zona)

    ws_sub = wb_out.create_sheet("Distribucion por Tecnico")
    sh = ['Zona', 'Tecnico', 'Subzonas y Estados', 'Pend. Totales', 'Finalizadas', 'Carga Total', '# Subzonas']
    ws_sub.append(sh)
    style_header_row(ws_sub, len(sh))
    ws_sub.row_dimensions[1].height = 18

    subzone_summaries.sort(key=lambda x: (x['zona'], x['tecnico']))
    for s in subzone_summaries:
        row_num = ws_sub.max_row + 1
        ws_sub.append([
            s['zona'], s['tecnico'], s['subzonas_detalle'],
            s['pendientes'], s['finalizadas'], s['carga_total'], s['num_subzonas']
        ])
        ws_sub.cell(row=row_num, column=3).alignment = Alignment(wrap_text=True, vertical='top')
        ws_sub.row_dimensions[row_num].height = max(18, s['num_subzonas'] * 16)

        if s['carga_total'] > MAX_ABSOLUTE_LOAD:
            ws_sub.cell(row=row_num, column=6).fill = ALERT_FILL
        elif s['carga_total'] >= MAX_IDEAL_LOAD:
            ws_sub.cell(row=row_num, column=6).fill = WARN_FILL
        elif s['pendientes'] <= 2:
            ws_sub.cell(row=row_num, column=4).fill = SUCCESS_FILL

    ws_sub.column_dimensions['C'].width = 60
    ws_sub.column_dimensions['A'].width = 16
    ws_sub.column_dimensions['B'].width = 28
    for col_letter in ['D', 'E', 'F', 'G']:
        ws_sub.column_dimensions[col_letter].width = 14

    ws_sug = wb_out.create_sheet("Sugerencias de Movimientos")
    sugh = [
        'Zona', 'Subzona', 'Direccion', 'De Tecnico (Origen)',
        'Pendientes Origen', 'A Tecnico (Destino)', 'Pendientes Destino',
        'ID Orden', 'Franja', 'Distancia Aprox.', 'Alerta', 'Justificación'
    ]
    ws_sug.append(sugh)
    style_header_row(ws_sug, len(sugh))

    for sug in suggestions:
        row_num = ws_sug.max_row + 1
        addr = sug.get('address', '') or sug.get('subzona', '')
        ws_sug.append([
            sug.get('zona', ''), sug.get('subzona', ''), addr,
            sug.get('origen', ''), sug.get('pendientes_origen', ''),
            sug.get('destino', ''), sug.get('pendientes_destino', ''),
            sug.get('order_id', ''), sug.get('franja', ''),
            sug.get('distancia_estimada', ''), sug.get('alerta', ''),
            sug.get('justificacion', '')
        ])
        if sug.get('alerta'):
            ws_sug.cell(row=row_num, column=11).fill = WARN_FILL

    auto_fit_columns(ws_sug)

    ws_alert = wb_out.create_sheet("Alertas")
    ah = ['Tipo', 'Zona', 'Tecnico', 'Detalle']
    ws_alert.append(ah)
    style_header_row(ws_alert, len(ah))

    priority_map = {'SOBRECARGA': 0, 'FRANJA EN RIESGO': 1, 'FRANJAS DUPLICADAS': 2,
                    'EXCESO TARDE': 3, 'FRANJA AJUSTADA': 4, 'MULTI-ESTADO ACTIVO': 5}
    alerts.sort(key=lambda a: priority_map.get(a['tipo'], 99))

    if alerts:
        for a in alerts:
            row_num = ws_alert.max_row + 1
            ws_alert.append([a['tipo'], a['zona'], a['tecnico'], a['detalle']])
            if a['tipo'] in ['SOBRECARGA', 'FRANJA EN RIESGO']:
                ws_alert.cell(row=row_num, column=1).fill = ALERT_FILL
            elif a['tipo'] in ['FRANJAS DUPLICADAS', 'EXCESO TARDE', 'MULTI-ESTADO ACTIVO']:
                ws_alert.cell(row=row_num, column=1).fill = WARN_FILL
    else:
        ws_alert.append(['OK', 'N/A', 'N/A', 'No se detectaron alertas'])

    auto_fit_columns(ws_alert)

    try:
        date_tag = now_bogota().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"sugerencias_nivelacion_{date_tag}.xlsx"
        output_path = os.path.join(os.getcwd(), filename)
        wb_out.save(output_path)

        msg_parts = [
            "Nivelacion completada.",
            f"  Sugerencias generadas: {len(suggestions)}",
            f"  Alertas detectadas: {len(alerts)}",
            f"  Zonas analizadas: {len(zones)}",
            f"  Total ordenes: {len(data)}",
        ]
        if not suggestions:
            msg_parts.append("\n  NOTA: No se generaron movimientos.")
            msg_parts.append("  La carga esta balanceada o las restricciones impiden mover.")

        return "\n".join(msg_parts), output_path
    except Exception as e:
        return f"Error guardando reporte: {str(e)}", None


if __name__ == "__main__":
    current_hour = datetime.now().hour + datetime.now().minute / 60.0
    f = get_latest_preruta_file()
    if f:
        msg, path = generate_suggestions(f)
        print(msg)
        if path:
            print(f"Archivo guardado en: {path}")
    else:
        print("No se encontro archivo Excel valido.")

