import math
from datetime import datetime

from config import (
    EFFICIENT_TECH_PROTECTION_SCORE,
    FRAGMENTATION_PENALTY,
    INTERZONE_DISTANCE_PENALTY,
    MAX_ABSOLUTE_LOAD,
    MAX_ALLOWED_DISTANCE_KM,
    MAX_DUPLICATED_SLOTS,
    MAX_IDEAL_LOAD,
    MAX_INTERZONE_ASSIGNMENTS_PER_TECH,
    MAX_ORDERS_PER_SLOT,
    MAX_ORDER_DURATION_HOURS,
    MAX_SUBZONES_SOFT,
    MIN_IMBALANCE_TO_MOVE,
    MIN_ROUTE_SAVINGS_KM,
    MIN_ROUTE_SAVINGS_PCT,
    MOVABLE_STATUSES,
    ORDER_DURATION_HOURS,
    STATUS_PROGRESS,
    ZONE_ADJACENCY,
)
from normalization import norm_zone
from geo_utils import haversine, get_centroid
from status_rules import get_status_progress, is_status


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
    except Exception:
        return None, None


def estimate_remaining_hours(status, onsite_hour=None, current_hour=None):
    progress = get_status_progress(status)
    base_duration = ORDER_DURATION_HOURS
    if progress == 3:
        base_duration = 0.6
    if progress == 4:
        base_duration = 0.3
    if progress >= 5:
        base_duration = 0.1

    if onsite_hour is not None and current_hour is not None:
        elapsed = current_hour - onsite_hour
        if elapsed > 0:
            remaining = max(0.1, base_duration - elapsed)
            return remaining
    return base_duration


def count_duplicated_slots(franja_counts):
    return sum(1 for count in franja_counts.values() if count >= 2)


def format_hour(h_decimal):
    if h_decimal is None:
        return 'N/A'
    h = int(h_decimal)
    m = int((h_decimal - h) * 60)
    return f'{h:02d}:{m:02d}'


def can_tech_handle_franja(tech_franja_counts, tech_all_orders, order_franja, current_hour,
                           allow_same_unit_override=False, same_unit=False):
    franja_start, franja_end = parse_franja_hours(order_franja)

    current_in_slot = tech_franja_counts.get(order_franja, 0)
    if current_in_slot >= MAX_ORDERS_PER_SLOT:
        return False, 'Ya tiene 2 ordenes en esta franja'

    if current_in_slot >= 1:
        existing_dups = count_duplicated_slots(tech_franja_counts)
        if existing_dups >= MAX_DUPLICATED_SLOTS:
            if allow_same_unit_override and same_unit:
                return True, 'OK (EXCEPCION: MISMA UNIDAD)'
            return False, 'Ya tiene su franja duplicada permitida'

    tarde_count = 0
    if franja_start is not None and franja_start >= 14.5:
        for f, c in tech_franja_counts.items():
            fs, _ = parse_franja_hours(f)
            if fs is not None and fs >= 14.5:
                tarde_count += c
        if tarde_count >= 2:
            return False, 'Ya tiene 2 ordenes en franjas de tarde (>=14:30)'

    if franja_start is not None:
        active_order = next((o for o in tech_all_orders if 1 <= get_status_progress(o['status']) < 6), None)
        rem_hours = 0
        if active_order:
            rem_hours = estimate_remaining_hours(active_order['status'], active_order.get('onsite_hour'), current_hour)

        prog_before = 0
        for o in tech_all_orders:
            if is_status(str(o['status']).lower(), MOVABLE_STATUSES):
                fs_h, _ = parse_franja_hours(o['franja'])
                if fs_h is not None and fs_h < franja_start:
                    prog_before += 1

        estimated_ready_hour = current_hour + rem_hours + (prog_before * ORDER_DURATION_HOURS)

        if franja_end is not None:
            is_already_late = current_hour > (franja_end - 0.5)
            if not is_already_late and estimated_ready_hour > (franja_end + 0.5):
                return False, f'No alcanza: listo ~{estimated_ready_hour:.1f}h, franja termina {franja_end:.1f}h'

    return True, 'OK'


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
        if is_status(str(o['status']).lower(), MOVABLE_STATUSES):
            fs_h, _ = parse_franja_hours(o['franja'])
            if fs_h is not None and fs_h < franja_start:
                prog_before += 1

    ready_normal = current_hour + rem_n + (prog_before * ORDER_DURATION_HOURS)
    ready_max = current_hour + rem_m + (prog_before * MAX_ORDER_DURATION_HOURS)

    arrival_normal = max(ready_normal, franja_start)
    arrival_max = max(ready_max, franja_start)

    return arrival_normal, arrival_max, franja_start, franja_end


def compute_min_route_savings_threshold(donor_dist, receiver_dist):
    base = MIN_ROUTE_SAVINGS_KM
    if donor_dist is not None and donor_dist > 0:
        base = max(base, donor_dist * MIN_ROUTE_SAVINGS_PCT)
    if receiver_dist is not None and receiver_dist > 0:
        base = max(base, receiver_dist * 0.20)
    return base


def tech_route_fragmentation_score(tech_name, tech_all_orders, tech_subzones, tech_franja_counts):
    orders = tech_all_orders.get(tech_name, [])
    subzone_count = len(tech_subzones.get(tech_name, set()))
    duplicated = count_duplicated_slots(tech_franja_counts.get(tech_name, {}))
    active_count = sum(1 for o in orders if get_status_progress(o.get('status')) >= 1)
    movable_count = sum(1 for o in orders if is_status(str(o.get('status', '')).lower(), MOVABLE_STATUSES))
    return (subzone_count * 1.2) + (duplicated * 1.8) + (active_count * 0.9) + (movable_count * 0.35)


def is_tech_efficient(tech_name, tech_all_orders, tech_franja_counts, tech_subzones, tech_completion_credit_map, tech_pending, tech_effective_load):
    if tech_name == 'SIN_ASIGNAR':
        return False
    pending = tech_pending.get(tech_name, 0)
    completed = tech_completion_credit_map.get(tech_name, 0.0)
    duplicated = count_duplicated_slots(tech_franja_counts.get(tech_name, {}))
    subzone_count = len(tech_subzones.get(tech_name, set()))
    active_count = sum(1 for o in tech_all_orders.get(tech_name, []) if 1 <= get_status_progress(o.get('status')) < 6)
    eff_load = tech_effective_load.get(tech_name, 0.0)
    if pending <= 0:
        return False
    return (
        completed >= EFFICIENT_TECH_PROTECTION_SCORE
        and duplicated == 0
        and subzone_count <= 2
        and active_count <= 1
        and eff_load <= max(2.2, pending + 0.5)
    )


def zone_rank_from(origin_zone, target_zone):
    oz = norm_zone(origin_zone)
    tz = norm_zone(target_zone)
    if oz == tz:
        return 0
    neighbors = ZONE_ADJACENCY.get(oz, [])
    try:
        return neighbors.index(tz) + 1
    except ValueError:
        return 999


def should_offer_interzone_support(z_low, z_high, donor, zone_techs, tech_pending, tech_all_orders, tech_effective_load, receivers_interzone_count):
    if receivers_interzone_count.get(donor, 0) >= MAX_INTERZONE_ASSIGNMENTS_PER_TECH:
        return False, 'ya hizo apoyo interzona'
    if zone_rank_from(z_low, z_high) != 1:
        return False, 'zona no vecina directa'
    origin_techs = [t for t in zone_techs.get(z_low, []) if t != 'SIN_ASIGNAR']
    if len(origin_techs) < 3:
        return False, 'zona origen queda corta de capacidad'
    if sum(1 for t in origin_techs if tech_pending.get(t, 0) >= 1) < 2:
        return False, 'no quedan al menos 2 tecnicos utiles en origen'
    if any(1 <= get_status_progress(o.get('status')) < 6 for o in tech_all_orders.get(donor, [])):
        return False, 'tecnico donante tiene orden activa'
    if tech_pending.get(donor, 0) > 1 or tech_effective_load.get(donor, 0.0) > 1.8:
        return False, 'donante no esta suficientemente liviano'
    return True, 'OK'


def get_tech_reference_point(tech_name, tech_all_orders, tech_locations):
    for o in tech_all_orders.get(tech_name, []):
        if 1 <= get_status_progress(o.get('status')) < 6 and o.get('lat', 0) and o.get('lon', 0):
            return (o['lat'], o['lon'])

    upcoming = []
    for o in tech_all_orders.get(tech_name, []):
        if is_status(str(o.get('status', '')).lower(), MOVABLE_STATUSES) and o.get('lat', 0) and o.get('lon', 0):
            fs, _ = parse_franja_hours(o.get('franja'))
            upcoming.append((fs if fs is not None else 99, o))
    if upcoming:
        upcoming.sort(key=lambda x: x[0])
        od = upcoming[0][1]
        return (od['lat'], od['lon'])

    locs = tech_locations.get(tech_name, [])
    if locs:
        return get_centroid(locs)
    return (0, 0)


def estimate_distance_to_receiver(order, receiver, tech_all_orders, tech_locations):
    if not order.get('lat', 0) or not order.get('lon', 0):
        return None
    ref = get_tech_reference_point(receiver, tech_all_orders, tech_locations)
    if not ref or ref == (0, 0):
        return None
    return haversine(order['lat'], order['lon'], ref[0], ref[1])


def tech_load_score(tech_name, tech_pending, tech_total, tech_effective_load, tech_completion_credit_map, tech_franja_counts, tech_subzones, tech_all_orders=None):
    duplicated = count_duplicated_slots(tech_franja_counts.get(tech_name, {}))
    subzone_count = len(tech_subzones.get(tech_name, set()))
    fragmentation = tech_route_fragmentation_score(tech_name, tech_all_orders or {}, tech_subzones, tech_franja_counts)
    return (
        tech_effective_load.get(tech_name, 0.0) * 1350
        + tech_pending.get(tech_name, 0) * 750
        + tech_total.get(tech_name, 0) * 150
        + duplicated * 650
        + max(0, subzone_count - 1) * 180
        + fragmentation * FRAGMENTATION_PENALTY
        - tech_completion_credit_map.get(tech_name, 0.0) * 220
    )


def estimate_route_savings_km(order, donor, receiver, tech_all_orders, tech_locations):
    receiver_dist = estimate_distance_to_receiver(order, receiver, tech_all_orders, tech_locations)
    donor_dist = None
    if donor and donor != 'SIN_ASIGNAR':
        donor_dist = estimate_distance_to_receiver(order, donor, tech_all_orders, tech_locations)

    if receiver_dist is None and donor_dist is None:
        return None, donor_dist, receiver_dist
    if receiver_dist is None:
        return None, donor_dist, receiver_dist
    if donor_dist is None:
        return receiver_dist * 0.25, donor_dist, receiver_dist
    return donor_dist - receiver_dist, donor_dist, receiver_dist


def should_allow_move_by_balance_or_route(order, donor, receiver, donor_eff, recv_eff, tech_pending, tech_total, tech_all_orders, tech_locations, current_hour,
                                          tech_franja_counts=None, tech_subzones=None, tech_completion_credit_map=None, tech_effective_load=None):
    load_gap = donor_eff - recv_eff
    savings_km, donor_dist, receiver_dist = estimate_route_savings_km(order, donor, receiver, tech_all_orders, tech_locations)
    min_required_savings = compute_min_route_savings_threshold(donor_dist, receiver_dist)

    route_override = False
    route_reason = ''

    if donor and donor != 'SIN_ASIGNAR' and tech_franja_counts is not None and tech_subzones is not None and tech_completion_credit_map is not None and tech_effective_load is not None:
        if is_tech_efficient(donor, tech_all_orders, tech_franja_counts, tech_subzones, tech_completion_credit_map, tech_pending, tech_effective_load):
            if savings_km is None or savings_km < max(min_required_savings, MIN_ROUTE_SAVINGS_KM * 1.5):
                arr_d_n, arr_d_m, f_start, f_end = estimate_arrival_for_franja(tech_all_orders.get(donor, []), order.get('franja'), current_hour)
                arr_r_n, arr_r_m, _, _ = estimate_arrival_for_franja(tech_all_orders.get(receiver, []), order.get('franja'), current_hour)
                donor_late = 0.0
                recv_late = 0.0
                if f_end is not None:
                    if arr_d_m is not None and arr_d_m > f_end:
                        donor_late = arr_d_m - f_end
                    elif arr_d_n is not None and arr_d_n > f_end:
                        donor_late = arr_d_n - f_end
                    if arr_r_m is not None and arr_r_m > f_end:
                        recv_late = arr_r_m - f_end
                    elif arr_r_n is not None and arr_r_n > f_end:
                        recv_late = arr_r_n - f_end
                if donor_late <= recv_late + 0.20:
                    return False, 'proteger_tecnico_eficiente', savings_km, donor_dist, receiver_dist, load_gap

    if load_gap >= MIN_IMBALANCE_TO_MOVE:
        route_override = True
        route_reason = f'desbalance={load_gap:.2f}'
    else:
        if savings_km is not None and savings_km >= min_required_savings:
            route_override = True
            route_reason = f'ahorro_ruta={savings_km:.2f}km'
        else:
            arr_d_n, arr_d_m, f_start, f_end = estimate_arrival_for_franja(tech_all_orders.get(donor, []), order.get('franja'), current_hour) if donor and donor != 'SIN_ASIGNAR' else (None, None, None, None)
            arr_r_n, arr_r_m, _, _ = estimate_arrival_for_franja(tech_all_orders.get(receiver, []), order.get('franja'), current_hour)
            donor_late = 0.0
            recv_late = 0.0
            if f_end is not None:
                if arr_d_m is not None and arr_d_m > f_end:
                    donor_late = arr_d_m - f_end
                elif arr_d_n is not None and arr_d_n > f_end:
                    donor_late = arr_d_n - f_end
                if arr_r_m is not None and arr_r_m > f_end:
                    recv_late = arr_r_m - f_end
                elif arr_r_n is not None and arr_r_n > f_end:
                    recv_late = arr_r_n - f_end
            if donor_late > recv_late + 0.20:
                route_override = True
                route_reason = f'mejora_franja={donor_late - recv_late:.2f}h'

    return route_override, route_reason, savings_km, donor_dist, receiver_dist, load_gap
