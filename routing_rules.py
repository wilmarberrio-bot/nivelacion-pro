from config import (
    MAX_DUPLICATED_SLOTS,
    MAX_ORDERS_PER_SLOT,
    MAX_ORDER_DURATION_HOURS,
    NEAR_FINISH_STATUSES,
    ORDER_DURATION_HOURS,
)
from status_rules import get_status_progress

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
