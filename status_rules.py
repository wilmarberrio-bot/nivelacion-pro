from config import (
    EFFICIENT_TECH_PROTECTION_SCORE,
    FINALIZED_STATUSES,
    INTERZONE_DISTANCE_PENALTY,
    MAX_ABSOLUTE_LOAD,
    MAX_IDEAL_LOAD,
    MAX_INTERZONE_ASSIGNMENTS_PER_TECH,
    MAX_SUBZONES_SOFT,
    MIN_IMBALANCE_TO_MOVE,
    MIN_ROUTE_SAVINGS_KM,
    MIN_ROUTE_SAVINGS_PCT,
    MIN_ROUTE_SCORE_BENEFIT,
    MOVABLE_STATUSES,
    STATUS_PROGRESS,
    ZONE_ADJACENCY,
)

def status_effective_weight(status):
    sl = str(status).lower()
    if 'cancelado' in sl or 'cancelada' in sl:
        return 0.35
    if any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada']):
        return 0.05
    if any(k in sl for k in NEAR_FINISH_STATUSES):
        return 0.65
    if get_status_progress(sl) >= 1:
        return 1.25
    if is_status(sl, MOVABLE_STATUSES):
        return 1.05
    return 0.95

def status_completion_credit(status):
    sl = str(status).lower()
    if any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada']):
        return 1.0
    if 'cancelado' in sl or 'cancelada' in sl:
        return 0.25
    return 0.0


def is_completed_or_auditable(status):
    sl = str(status).lower()
    return any(k in sl for k in ['finalizado', 'finalizada', 'por auditar', 'cerrado', 'cerrada', 'completado', 'completada'])

def should_alert_low_progress(tech_name, completed_count, peer_counts):
    if tech_name == "SIN_ASIGNAR":
        return False
    if completed_count > 1:
        return False
    valid_peers = [c for c in peer_counts if c is not None]
    if len(valid_peers) < 2:
        return False
    peers_with_2_or_more = sum(1 for c in valid_peers if c >= 2)
    peers_with_3_or_more = sum(1 for c in valid_peers if c >= 3)
    peer_avg = sum(valid_peers) / len(valid_peers) if valid_peers else 0
    # Dispara cuando el técnico solo lleva 0-1 cerradas/auditables y el resto del grupo ya va claramente más adelante.
    return peers_with_2_or_more >= 2 and (peers_with_3_or_more >= 1 or peer_avg >= 2.0)

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
    if donor and donor != "SIN_ASIGNAR":
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
    route_reason = ""

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
        route_reason = f"desbalance={load_gap:.2f}"
    else:
        if savings_km is not None and savings_km >= min_required_savings:
            route_override = True
            route_reason = f"ahorro_ruta={savings_km:.2f}km"
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
                route_reason = f"mejora_franja={donor_late-recv_late:.2f}h"

    return route_override, route_reason, savings_km, donor_dist, receiver_dist, load_gap
