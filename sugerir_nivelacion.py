import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import glob
from datetime import datetime

# --- CONFIGURACION ---
# Defaults will be overridden inside functions for multi-user safety
# (Removed static global filenames)
CURRENT_HOUR = datetime.now().hour + datetime.now().minute / 60.0  # Hora actual decimal

# Constraints
MAX_IDEAL_LOAD = 5       # Carga ideal maxima por tecnico
MAX_ABSOLUTE_LOAD = 6    # Carga maxima absoluta (ultima opcion)
MAX_ORDERS_PER_SLOT = 2  # Permitir solape como ultimo recurso
MAX_DUPLICATED_SLOTS = 1 # Permitir maximo 1 solape
MIN_IMBALANCE_TO_MOVE = 1  # Reducido de 2 a 1 para permitir balanceo mas fino (v5)
ORDER_DURATION_HOURS = 1.25  # Duracion estimada por orden
MAX_ALLOWED_DISTANCE_KM = 8.0 # Aumentado de 5 a 8 para dar mas margen

# Estados - case insensitive
MOVABLE_STATUSES = ['programado', 'programada']
FINALIZED_STATUSES = ['finalizado', 'finalizada', 'por auditar', 'cancelado', 'cancelada',
                      'cerrado', 'cerrada', 'completado', 'completada']

# Progresion de estados activos: valor = que tan cerca de finalizar (mayor = mas cerca)
STATUS_PROGRESS = {
    'programado': 0,
    'programada': 0,
    'inbound': 1,         # En camino al sitio
    'en sitio': 2,        # Apenas llego
    'iniciado': 3,        # Dentro del apartamento/casa trabajando
    'iniciada': 3,
    'mac principal enviada': 4,  # Ya monto equipo principal
    'dispositivos cargados': 5,  # Ya va a finalizar
}

# Estados que indican que el tecnico esta a punto de terminar la orden actual
NEAR_FINISH_STATUSES = ['dispositivos cargados', 'mac principal enviada']

# Estilos Excel
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
ALERT_FILL = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
SUCCESS_FILL = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
WARN_FILL = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)


def get_latest_preruta_file():
    files = glob.glob("*.xlsx")
    candidates = [f for f in files
                  if ("pre_ruta" in f.lower() or "nivelacion" in f.lower())
                  and not f.startswith("~$") and not f.startswith("sugerencias_")]
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
    """Extrae la hora inicio y fin de una franja. Ej '08:00-09:30' -> (8.0, 9.5)"""
    if not franja_str or franja_str == 'Sin Franja':
        return None, None
    try:
        clean = franja_str.replace('\u2013', '-').replace('\u2014', '-').replace('\ufffd', '-')
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
        return start, end
    except:
        return None, None


def is_status(status_lower, status_list):
    for s in status_list:
        if s in status_lower:
            return True
    return False


def get_status_progress(status):
    """Retorna que tan avanzado esta un estado (0=no iniciado, 5=a punto de finalizar)."""
    sl = status.lower()
    for key, val in STATUS_PROGRESS.items():
        if key in sl:
            return val
    return 0


def count_duplicated_slots(franja_counts):
    return sum(1 for count in franja_counts.values() if count >= 2)


def can_tech_handle_franja(tech_franja_counts, tech_all_orders, order_franja, current_hour):
    """
    Verifica si un tecnico puede atender una orden en la franja dada,
    considerando:
    - No mas de 2 ordenes por franja
    - Solo 1 franja duplicada
    - Tiempo real: si tiene una orden activa que no ha terminado, la siguiente se retrasa
    - Si la franja es >=14:30, no apilar mas de 1 orden de tarde
    """
    franja_start, franja_end = parse_franja_hours(order_franja)

    # Restriccion: max ordenes por slot
    current_in_slot = tech_franja_counts.get(order_franja, 0)
    if current_in_slot >= MAX_ORDERS_PER_SLOT:
        return False, "Ya tiene 2 ordenes en esta franja"

    # Restriccion: solo 1 franja duplicada
    if current_in_slot >= 1:
        existing_dups = count_duplicated_slots(tech_franja_counts)
        if existing_dups >= MAX_DUPLICATED_SLOTS:
            return False, "Ya tiene su franja duplicada permitida"

    # Restriccion: tarde (>= 14:30)
    tarde_count = 0
    if franja_start is not None and franja_start >= 14.5:
        for f, c in tech_franja_counts.items():
            fs, _ = parse_franja_hours(f)
            if fs is not None and fs >= 14.5:
                tarde_count += c
        if tarde_count >= 2:
            return False, "Ya tiene 2 ordenes en franjas de tarde (>=14:30)"

    # Restriccion temporal: puede el tecnico atender esta franja a tiempo?
    if franja_start is not None:
        # Verificar si ya tiene UNA orden programada en esta franja exacta o que se solape
        for o in tech_all_orders:
            o_status = o['status'].lower()
            if is_status(o_status, FINALIZED_STATUSES):
                continue
            
            o_start, o_end = parse_franja_hours(o['franja'])
            if o_start is None: continue

            # Si la franja de la orden objetivo coincide con una que ya tiene
            if o_start == franja_start:
                return False, f"Ya tiene orden en franja {order_franja}"

        # Estimar carga actual para ver si llega a tiempo
        active_now = sum(1 for o in tech_all_orders if get_status_progress(o['status']) >= 1)
        prog_before = sum(1 for o in tech_all_orders 
                         if is_status(o['status'].lower(), MOVABLE_STATUSES)
                         and parse_franja_hours(o['franja'])[0] is not None
                         and parse_franja_hours(o['franja'])[0] < franja_start)

        estimated_ready_hour = current_hour + (active_now + prog_before) * ORDER_DURATION_HOURS

        if franja_end is not None and estimated_ready_hour > franja_end:
            return False, f"No alcanza: estaria listo ~{estimated_ready_hour:.1f}h, franja termina {franja_end:.1f}h"

    return True, "OK"

    return True, "OK"


def coords_to_sector(lat, lon, subzona):
    if lat == 0 and lon == 0:
        return "Sin ubicacion"
    return f"{subzona} ({round(lat, 4)}, {round(lon, 4)})"


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


def generate_suggestions(input_file):
    print(f"Leyendo archivo: {input_file}")
    print(f"Hora actual: {CURRENT_HOUR:.2f} ({datetime.now().strftime('%H:%M')})")

    try:
        wb = openpyxl.load_workbook(input_file, read_only=True, data_only=True)
        sheet = wb.active
    except Exception as e:
        return f"Error leyendo el archivo: {str(e)}", None

    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]

    idx_status = get_col_index(headers, ['status_txt', 'Estado', 'Estado de orden de trabajo'])
    # Cities__name = Zona principal
    idx_zone = get_col_index(headers, ['Cities__name', 'Ciudad', 'Zona'])
    idx_zone_fallback = get_col_index(headers, ['Zone Name'])
    idx_zone_op = get_col_index(headers, ['zona_op'])
    # Subzone = Subzona real (prioridad alta), Sites = Backup
    idx_subzone = get_col_index(headers, ['Subzone', 'Subzona', 'Sites'])
    idx_tech = get_col_index(headers, ['tecnico', 'technenvician', 'Tecnico asignado', 'Tecnico', 'Tech'])
    idx_order = get_col_index(headers, ['appointment_id', 'ID', 'Numero de orden', 'Order ID'])
    idx_franja = get_col_index(headers, ['franja_label', 'Franja', 'Cita', 'Ventana'])
    idx_lat = get_col_index(headers, ['Latitude', 'Latitud', 'lat'])
    idx_lon = get_col_index(headers, ['Longitude', 'Longitud', 'lon', 'lng'])
    idx_address = get_col_index(headers, ['Addresses__address', 'Direccion', 'direccion'])

    missing = []
    if idx_tech == -1: missing.append("Tecnico")
    if idx_order == -1: missing.append("ID Orden")
    if missing:
        return f"Error: No se encontraron las columnas: {', '.join(missing)}", None

    # --- CARGAR DATOS ---
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        tech = str(row[idx_tech]).strip() if (idx_tech != -1 and row[idx_tech]) else "SIN_ASIGNAR"
        if not tech or tech.lower() in ['none', 'nan', '']:
            tech = "SIN_ASIGNAR"

        zona = "SIN_ZONA"
        # Prioridad: Cities__name > Zone Name > zona_op
        if idx_zone != -1 and row[idx_zone] and str(row[idx_zone]).strip().lower() not in ['none', '']:
            zona = str(row[idx_zone]).strip()
        elif idx_zone_fallback != -1 and row[idx_zone_fallback] and str(row[idx_zone_fallback]).strip().lower() not in ['none', '']:
            zona = str(row[idx_zone_fallback]).strip()
        elif idx_zone_op != -1 and row[idx_zone_op] and str(row[idx_zone_op]).strip().lower() not in ['none', '']:
            zona = "Zona " + str(row[idx_zone_op]).strip()

        subzona = "SIN_SUBZONA"
        if idx_subzone != -1 and row[idx_subzone] and str(row[idx_subzone]).strip().lower() not in ['none', '']:
            subzona = str(row[idx_subzone]).strip()

        order_id = row[idx_order] if idx_order != -1 else "N/A"
        status = str(row[idx_status]).strip() if (idx_status != -1 and row[idx_status]) else "Sin Estado"
        franja = str(row[idx_franja]).strip() if (idx_franja != -1 and row[idx_franja]) else "Sin Franja"
        franja = franja.replace('\ufffd', '-').replace('\u2013', '-').replace('\u2014', '-')

        lat, lon = 0.0, 0.0
        if idx_lat != -1 and row[idx_lat]:
            try: lat = float(row[idx_lat])
            except: pass
        if idx_lon != -1 and row[idx_lon]:
            try: lon = float(row[idx_lon])
            except: pass

        address = ""
        if idx_address != -1 and row[idx_address]:
            address = str(row[idx_address]).strip()

        data.append({
            'tech': tech, 'zona': zona, 'subzona': subzona,
            'order_id': order_id, 'status': status, 'franja': franja,
            'lat': lat, 'lon': lon, 'address': address,
            'sector': coords_to_sector(lat, lon, subzona)
        })

    wb.close()

    if not data:
        return "Error: No se encontraron datos en el archivo.", None

    print(f"Total ordenes: {len(data)}")

    # --- ANALISIS POR ZONA ---
    zones = sorted(set(d['zona'] for d in data if d['zona'] != "SIN_ZONA"))

    suggestions = []
    alerts = []
    zone_summaries = []
    subzone_summaries = []

    for z in zones:
        zone_data = [d for d in data if d['zona'] == z]

        # ===========================================
        # PASO 1: Calcular carga por tecnico
        # ===========================================
        tech_total = {}
        tech_finalized = {}
        tech_pending = {}
        tech_movable = {}
        tech_all_orders = {}   # TODAS las ordenes del tecnico (para validaciones de tiempo)
        tech_locations = {}
        tech_franja_counts = {}
        tech_status_detail = {}
        tech_subzones = {}
        tech_has_near_finish = {}  # Tecnico tiene orden a punto de terminar

        for d in zone_data:
            t = d['tech']
            status_lower = d['status'].lower()

            if t == "SIN_ASIGNAR":
                if is_status(status_lower, MOVABLE_STATUSES):
                    if t not in tech_movable: tech_movable[t] = []
                    tech_movable[t].append(d)
                continue

            # Registrar TODAS las ordenes del tecnico
            if t not in tech_all_orders: tech_all_orders[t] = []
            tech_all_orders[t].append(d)

            tech_total[t] = tech_total.get(t, 0) + 1

            if is_status(status_lower, FINALIZED_STATUSES):
                tech_finalized[t] = tech_finalized.get(t, 0) + 1
            else:
                tech_pending[t] = tech_pending.get(t, 0) + 1

                if t not in tech_franja_counts: tech_franja_counts[t] = {}
                tech_franja_counts[t][d['franja']] = tech_franja_counts[t].get(d['franja'], 0) + 1

                if d['lat'] != 0 and d['lon'] != 0:
                    if t not in tech_locations: tech_locations[t] = []
                    tech_locations[t].append((d['lat'], d['lon']))

                if t not in tech_subzones: tech_subzones[t] = set()
                tech_subzones[t].add(d['subzona'])

                if t not in tech_status_detail: tech_status_detail[t] = {}
                tech_status_detail[t][d['status']] = tech_status_detail[t].get(d['status'], 0) + 1

                # Detectar si tiene ordenes a punto de terminar
                if is_status(status_lower, NEAR_FINISH_STATUSES):
                    tech_has_near_finish[t] = True

            if is_status(status_lower, MOVABLE_STATUSES):
                if t not in tech_movable: tech_movable[t] = []
                tech_movable[t].append(d)

        # ===========================================
        # PASO 2: Alertas de franja proxima a vencer
        # ===========================================
        for t in tech_total:
            total = tech_total[t]
            pending = tech_pending.get(t, 0)

            if total > MAX_ABSOLUTE_LOAD:
                alerts.append({
                    'tipo': 'SOBRECARGA',
                    'zona': z, 'tecnico': t,
                    'detalle': f"Tiene {total} ordenes totales (max: {MAX_ABSOLUTE_LOAD})"
                })

            # Alertas de franjas proximas a vencer
            for d in tech_all_orders.get(t, []):
                if not is_status(d['status'].lower(), MOVABLE_STATUSES):
                    continue  # Solo alertar programadas, las activas ya estan atendiendose

                f_start, f_end = parse_franja_hours(d['franja'])
                if f_start is not None:
                    # Cuantas ordenes tiene el tecnico antes de esta?
                    active_now = sum(1 for o in tech_all_orders.get(t, [])
                                    if get_status_progress(o['status']) >= 1)
                    prog_before = sum(1 for o in tech_all_orders.get(t, [])
                                     if is_status(o['status'].lower(), MOVABLE_STATUSES)
                                     and parse_franja_hours(o['franja'])[0] is not None
                                     and parse_franja_hours(o['franja'])[0] < f_start)

                    est_ready = CURRENT_HOUR + (active_now + prog_before) * ORDER_DURATION_HOURS

                    if f_end is not None and est_ready > f_end:
                        alerts.append({
                            'tipo': 'FRANJA EN RIESGO',
                            'zona': z, 'tecnico': t,
                            'detalle': f"Orden {d['order_id']} franja {d['franja']} - "
                                      f"Tecnico estaria listo ~{est_ready:.1f}h pero franja termina {f_end:.1f}h. "
                                      f"Considerar mover a otro tecnico que este por finalizar."
                        })
                    elif f_start is not None and est_ready > f_start and f_end and est_ready <= f_end:
                        alerts.append({
                            'tipo': 'FRANJA AJUSTADA',
                            'zona': z, 'tecnico': t,
                            'detalle': f"Orden {d['order_id']} franja {d['franja']} - "
                                      f"Tecnico llegaria ~{est_ready:.1f}h (franja inicia {f_start:.1f}h). Justo a tiempo."
                        })

            # Alertas de franjas duplicadas existentes
            if t in tech_franja_counts:
                duplicated = count_duplicated_slots(tech_franja_counts[t])
                if duplicated > MAX_DUPLICATED_SLOTS:
                    franjas_dup = [f"{f}: {c}" for f, c in tech_franja_counts[t].items() if c >= 2]
                    alerts.append({
                        'tipo': 'FRANJAS DUPLICADAS',
                        'zona': z, 'tecnico': t,
                        'detalle': f"{duplicated} franjas duplicadas: {', '.join(franjas_dup)}"
                    })

        # ===========================================
        # PASO 3: Resumen por Zona
        # ===========================================
        if tech_pending:
            avg_pending = sum(tech_pending.values()) / len(tech_pending)
            zone_summaries.append({
                'zona': z,
                'techs': len(tech_pending),
                'total_pendientes': sum(tech_pending.values()),
                'total_finalizadas': sum(tech_finalized.get(t, 0) for t in tech_total),
                'avg': round(avg_pending, 1),
                'min': min(tech_pending.values()),
                'max': max(tech_pending.values()),
            })

        # ===========================================
        # PASO 4: Resumen por Subzona (tabla dinamica)
        # ===========================================
        active_zone_data = [d for d in zone_data
                           if d['tech'] != "SIN_ASIGNAR"
                           and not is_status(d['status'].lower(), FINALIZED_STATUSES)]

        subzones_in_zone = sorted(set(d['subzona'] for d in active_zone_data))

        for sz in subzones_in_zone:
            sz_data = [d for d in active_zone_data if d['subzona'] == sz]
            sz_tech_detail = {}

            for d in sz_data:
                t = d['tech']
                if t not in sz_tech_detail:
                    sz_tech_detail[t] = {'total': 0, 'estados': {}}
                sz_tech_detail[t]['total'] += 1
                sz_tech_detail[t]['estados'][d['status']] = \
                    sz_tech_detail[t]['estados'].get(d['status'], 0) + 1

            for t, info in sorted(sz_tech_detail.items()):
                breakdown = ", ".join([f"{k}: {v}" for k, v in info['estados'].items()])
                subzone_summaries.append({
                    'zona': z, 'subzona': sz, 'tecnico': t,
                    'pendientes': info['total'],
                    'detalle_estados': breakdown,
                    'finalizadas': tech_finalized.get(t, 0),
                    'carga_total': tech_total.get(t, 0),
                })

        # ===========================================
        # PASO 5: NIVELACION INTELIGENTE
        # ===========================================
        all_techs_in_zone = set(tech_total.keys())

        # DONANTES: quienes tienen ordenes que sobran
        donors = []

        # 1) Ordenes SIN_ASIGNAR
        if "SIN_ASIGNAR" in tech_movable and tech_movable["SIN_ASIGNAR"]:
            donors.append("SIN_ASIGNAR")

        # 2) Tecnicos con carga total > MAX_IDEAL_LOAD y ordenes programadas para mover
        for t in sorted(tech_total.keys(), key=lambda x: tech_total[x], reverse=True):
            if tech_total[t] > MAX_IDEAL_LOAD and t in tech_movable and tech_movable[t]:
                donors.append(t)

        # 3) Si hay desbalance significativo (>= MIN_IMBALANCE_TO_MOVE) en pendientes
        if tech_pending:
            pending_vals = list(tech_pending.values())
            avg_p = sum(pending_vals) / len(pending_vals)
            max_p = max(pending_vals)
            min_p = min(pending_vals)

            if max_p - min_p >= MIN_IMBALANCE_TO_MOVE:
                for t in sorted(tech_pending.keys(), key=lambda x: tech_pending[x], reverse=True):
                    if t not in donors and tech_pending[t] >= avg_p + 1:
                        if t in tech_movable and tech_movable[t]:
                            donors.append(t)

        for donor in donors:
            donor_orders = list(tech_movable.get(donor, []))
            if not donor_orders:
                continue

            if donor == "SIN_ASIGNAR":
                moves_needed = len(donor_orders)
            else:
                excess_total = tech_total.get(donor, 0) - MAX_IDEAL_LOAD
                excess_pending = tech_pending.get(donor, 0) - max(3, int(avg_p))
                moves_needed = max(1, int(max(excess_total, excess_pending)))

            for _ in range(min(moves_needed, len(donor_orders))):
                # Buscar receptores
                eligible = []
                for t in all_techs_in_zone:
                    if t == donor:
                        continue
                    t_total = tech_total.get(t, 0)
                    t_pending = tech_pending.get(t, 0)

                    # REGLA CLAVE: no mover si la diferencia resultante seria < MIN_IMBALANCE_TO_MOVE
                    if donor != "SIN_ASIGNAR":
                        donor_pending_after = tech_pending.get(donor, 0) - 1
                        receiver_pending_after = t_pending + 1
                        # REGLA v5: Solo bloquear si el receptor queda con MAS que el donante.
                        # Si quedan IGUALES (ej 3->2 y 1->2), es un movimiento vlido para nivelar.
                        if receiver_pending_after > donor_pending_after and donor_pending_after <= MAX_IDEAL_LOAD:
                            continue

                    # Preferir receptores con carga < MAX_IDEAL_LOAD
                    if t_total < MAX_IDEAL_LOAD:
                        priority = 'ideal'
                    elif t_total < MAX_ABSOLUTE_LOAD:
                        priority = 'last_resort'
                    else:
                        continue  # Ya esta lleno

                    eligible.append((t, priority))

                if not eligible:
                    break

                # Buscar mejor match
                best_order = None
                best_receiver = None
                best_score = float('inf')

                for order in donor_orders:
                    order_franja = order['franja']
                    order_has_loc = (order['lat'] != 0 and order['lon'] != 0)
                    order_hora_start, order_hora_end = parse_franja_hours(order_franja)

                    for r, r_type in eligible:
                        # VALIDAR franjas + tiempo
                        can_handle, reason = can_tech_handle_franja(
                            tech_franja_counts.get(r, {}),
                            tech_all_orders.get(r, []),
                            order_franja,
                            CURRENT_HOUR
                        )
                        if not can_handle:
                            continue

                        # SCORING
                        score = 0

                        # Priorizar ideal vs last_resort
                        if r_type == 'last_resort':
                            score += 3000

                        # PENALIZACION FUERTE por solape de franja (v5.1)
                        receiver_in_slot = tech_franja_counts.get(r, {}).get(order_franja, 0)
                        if receiver_in_slot >= 1:
                            score += 5000  # Solo mover si es extremadamente necesario

                        # Bonus si el receptor tiene orden a punto de terminar (va a liberar capacidad)
                        if tech_has_near_finish.get(r, False):
                            score -= 400

                        # Bonus por misma subzona (v4)
                        r_subzones = tech_subzones.get(r, set())
                        if order['subzona'] in r_subzones:
                            score -= 2500  # Aumentado de 1000 a 2500
                        else:
                            score += 1500  # Penalización por cambio de subzona (v4)

                        # Penalizar por carga actual
                        score += tech_total.get(r, 0) * 200
                        score += tech_pending.get(r, 0) * 100

                        # Distancia geografica
                        if order_has_loc and r in tech_locations:
                            centroid = get_centroid(tech_locations[r])
                            dist = haversine(order['lat'], order['lon'],
                                           centroid[0], centroid[1])
                            
                            # v4: Límite duro de distancia
                            if dist > MAX_ALLOWED_DISTANCE_KM:
                                continue # Traslado muy largo, descartado

                            score += dist * 100 # Mayor peso a la distancia (v4)

                            if dist < 1.0:
                                score -= 1000
                            elif dist < 3.0:
                                score -= 500
                        else:
                            # Sin coordenadas en el receptor o en la orden
                            # Solo permitir si son de la misma subzona literal para evitar errores v4
                            if order['subzona'] in r_subzones:
                                score += 500
                            else:
                                score += 8000 # Muy penalizado si no hay datos espaciales ni subzona match


                        if score < best_score:
                            best_score = score
                            best_order = order
                            best_receiver = r

                if best_order and best_receiver:
                    dist_str = "N/A"
                    if best_receiver in tech_locations and best_order['lat'] != 0:
                        centroid = get_centroid(tech_locations[best_receiver])
                        dist_km = haversine(best_order['lat'], best_order['lon'],
                                          centroid[0], centroid[1])
                        dist_str = f"{dist_km:.2f} km"

                    alerta = ""
                    o_start, o_end = parse_franja_hours(best_order['franja'])
                    if o_start is not None and o_start >= 14.5:
                        alerta = "Franja Tarde"
                    if o_end is not None:
                        active_r = sum(1 for o in tech_all_orders.get(best_receiver, [])
                                      if get_status_progress(o['status']) >= 1)
                        est = CURRENT_HOUR + active_r * ORDER_DURATION_HOURS
                        if est > o_start:
                            alerta += " | Ajustado en tiempo"

                    near_finish_tag = ""
                    if tech_has_near_finish.get(best_receiver, False):
                        near_finish_tag = " (por finalizar orden)"

                    suggestions.append({
                        'zona': z, 'subzona': best_order['subzona'],
                        'origen': donor,
                        'destino': best_receiver + near_finish_tag,
                        'destino_raw': best_receiver,
                        'order_id': best_order['order_id'],
                        'status_orden': best_order['status'],
                        'franja': best_order['franja'],
                        'distancia_estimada': dist_str,
                        'sector': best_order['sector'],
                        'address': best_order.get('address', ''),
                        'alerta': alerta,
                        'carga_origen_antes': tech_total.get(donor, 0),
                        'carga_destino_antes': tech_total.get(best_receiver, 0),
                        'pendientes_origen': tech_pending.get(donor, 0),
                        'pendientes_destino': tech_pending.get(best_receiver, 0),
                    })

                    # Actualizar contadores
                    if donor != "SIN_ASIGNAR":
                        tech_total[donor] -= 1
                        tech_pending[donor] = max(0, tech_pending.get(donor, 1) - 1)

                    tech_total[best_receiver] = tech_total.get(best_receiver, 0) + 1
                    tech_pending[best_receiver] = tech_pending.get(best_receiver, 0) + 1

                    if best_receiver not in tech_franja_counts:
                        tech_franja_counts[best_receiver] = {}
                    tech_franja_counts[best_receiver][best_order['franja']] = \
                        tech_franja_counts[best_receiver].get(best_order['franja'], 0) + 1

                    if best_order['lat'] != 0:
                        if best_receiver not in tech_locations:
                            tech_locations[best_receiver] = []
                        tech_locations[best_receiver].append(
                            (best_order['lat'], best_order['lon']))

                    if best_receiver not in tech_subzones:
                        tech_subzones[best_receiver] = set()
                    tech_subzones[best_receiver].add(best_order['subzona'])

                    # Agregar orden al registro del receptor
                    if best_receiver not in tech_all_orders:
                        tech_all_orders[best_receiver] = []
                    tech_all_orders[best_receiver].append(best_order)

                    donor_orders.remove(best_order)

    # ===========================================
    # GENERAR EXCEL
    # ===========================================
    wb_out = openpyxl.Workbook()

    # --- HOJA 1: Resumen por Zona ---
    ws_zona = wb_out.active
    ws_zona.title = "Resumen por Zona"
    zh = ['Zona', 'Tecnicos Activos', 'Total Pendientes', 'Total Finalizadas',
          'Promedio Pendientes', 'Min', 'Max', 'Desbalance']
    ws_zona.append(zh)
    style_header_row(ws_zona, len(zh))

    for s in zone_summaries:
        row_num = ws_zona.max_row + 1
        ws_zona.append([
            s['zona'], s['techs'], s['total_pendientes'], s['total_finalizadas'],
            s['avg'], s['min'], s['max'], s['max'] - s['min']
        ])
        d = s['max'] - s['min']
        if d >= 3: ws_zona.cell(row=row_num, column=8).fill = ALERT_FILL
        elif d >= 2: ws_zona.cell(row=row_num, column=8).fill = WARN_FILL

    auto_fit_columns(ws_zona)

    # --- HOJA 2: Resumen por Subzona ---
    ws_sub = wb_out.create_sheet("Resumen por Subzona")
    sh = ['Zona', 'Subzona', 'Tecnico', 'Pendientes', 'Finalizadas',
          'Carga Total', 'Detalle Estados']
    ws_sub.append(sh)
    style_header_row(ws_sub, len(sh))

    subzone_summaries.sort(key=lambda x: (x['zona'], x['subzona'], x['tecnico']))
    for s in subzone_summaries:
        row_num = ws_sub.max_row + 1
        ws_sub.append([
            s['zona'], s['subzona'], s['tecnico'], s['pendientes'],
            s['finalizadas'], s['carga_total'], s['detalle_estados']
        ])
        if s['carga_total'] > MAX_ABSOLUTE_LOAD:
            ws_sub.cell(row=row_num, column=6).fill = ALERT_FILL
        elif s['carga_total'] >= MAX_IDEAL_LOAD:
            ws_sub.cell(row=row_num, column=6).fill = WARN_FILL
        elif s['pendientes'] <= 2:
            ws_sub.cell(row=row_num, column=4).fill = SUCCESS_FILL

    auto_fit_columns(ws_sub)

    # --- HOJA 3: Sugerencias ---
    ws_sug = wb_out.create_sheet("Sugerencias de Movimientos")
    sugh = [
        'Zona', 'Subzona', 'Direccion', 'De Tecnico (Origen)',
        'Pendientes Origen', 'A Tecnico (Destino)', 'Pendientes Destino',
        'ID Orden', 'Franja', 'Distancia Aprox.', 'Alerta'
    ]
    ws_sug.append(sugh)
    style_header_row(ws_sug, len(sugh))

    for sug in suggestions:
        row_num = ws_sug.max_row + 1
        addr = sug.get('address', '') or sug['sector']
        ws_sug.append([
            sug['zona'], sug['subzona'], addr,
            sug['origen'], sug['pendientes_origen'],
            sug['destino'], sug['pendientes_destino'],
            sug['order_id'], sug['franja'],
            sug['distancia_estimada'], sug['alerta']
        ])
        if sug['alerta']:
            ws_sug.cell(row=row_num, column=11).fill = WARN_FILL

    auto_fit_columns(ws_sug)

    # --- HOJA 4: Alertas ---
    ws_alert = wb_out.create_sheet("Alertas")
    ah = ['Tipo', 'Zona', 'Tecnico', 'Detalle']
    ws_alert.append(ah)
    style_header_row(ws_alert, len(ah))

    # Ordenar alertas por prioridad
    priority_map = {'SOBRECARGA': 0, 'FRANJA EN RIESGO': 1, 'FRANJAS DUPLICADAS': 2,
                    'EXCESO TARDE': 3, 'FRANJA AJUSTADA': 4}
    alerts.sort(key=lambda a: priority_map.get(a['tipo'], 99))

    if alerts:
        for a in alerts:
            row_num = ws_alert.max_row + 1
            ws_alert.append([a['tipo'], a['zona'], a['tecnico'], a['detalle']])
            if a['tipo'] in ['SOBRECARGA', 'FRANJA EN RIESGO']:
                ws_alert.cell(row=row_num, column=1).fill = ALERT_FILL
            elif a['tipo'] in ['FRANJAS DUPLICADAS', 'EXCESO TARDE']:
                ws_alert.cell(row=row_num, column=1).fill = WARN_FILL
    else:
        ws_alert.append(['OK', 'N/A', 'N/A', 'No se detectaron alertas'])

    auto_fit_columns(ws_alert)

    # --- Guardar ---
    try:
        date_tag = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"sugerencias_nivelacion_{date_tag}.xlsx"
        output_path = os.path.join(os.getcwd(), filename)
        wb_out.save(output_path)

        msg_parts = [
            f"Nivelacion completada.",
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
    f = get_latest_preruta_file()
    if f:
        msg, path = generate_suggestions(f)
        print(msg)
        if path:
            print(f"Archivo guardado en: {path}")
    else:
        print("No se encontro archivo Excel valido.")
