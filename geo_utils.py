import math
from normalization import norm_text, norm_zone, norm_subzone
from config import MAX_ALLOWED_DISTANCE_KM, NEARBY_BUILDING_RADIUS_KM, ZONE_ADJACENCY

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

def extract_coords_from_text(text):
    """
    Extrae lat/lon de:
      - google maps: .../@lat,lon,17z
      - query: ...?q=lat,lon
      - query: ...?ll=lat,lon
      - google maps data: ...!3dLAT!4dLON
      - texto: "lat, lon"
    Nota: si el link está acortado (bit.ly / goo.gl) o no contiene coords,
    no es posible extraer sin resolver redirecciones.
    """
    if not text:
        return None, None
    t = str(text)

    # patrón @lat,lon
    m = re.search(r'@(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    # patrón q=lat,lon
    m = re.search(r'[?&]q=(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    # patrón ll=lat,lon
    m = re.search(r'[?&]ll=(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    # patrón !3dLAT!4dLON (típico en /data=... de Google Maps)
    m = re.search(r'!3d(-?\d{1,3}\.\d+)!4d(-?\d{1,3}\.\d+)', t)
    if m:
        return float(m.group(1)), float(m.group(2))

    # patrón "lat, lon"
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




def order_has_coords(order: dict) -> bool:
    return bool(order.get('coords_ok')) or bool(order.get('lat', 0) and order.get('lon', 0))

def allow_subzone_move_when_no_coords(order: dict, receiver: str, tech_subzones: dict, tech_all_orders: dict) -> bool:
    """Regla dura solicitada:
    - Si una orden NO tiene coords (ni lat/lon ni pudo extraerse del Google Maps Link),
      NO permitir movimientos entre subzonas por distancia.
    - Solo permitir reasignación si:
        a) MISMA UNIDAD (addr_key match) con el receptor, o
        b) El receptor YA trabaja esa subzona (conocimiento local).
    """
    if order_has_coords(order):
        return True
    # sin coords: solo misma unidad o misma subzona
    ak = order.get('addr_key', '')
    if ak:
        for o in tech_all_orders.get(receiver, []):
            if o.get('addr_key') == ak:
                return True
    sz = order.get('subzona')
    if sz and (sz in tech_subzones.get(receiver, set())):
        return True
    return False

def allow_zone_only_assignment_when_no_coords(order: dict) -> bool:
    """Si no hay coords, permitimos sugerir SOLO por zona (sin prometer subzona)."""
    return not order_has_coords(order)

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

def coords_to_sector(lat, lon, subzona):
    if lat == 0 and lon == 0:
        return "Sin ubicacion"
    try:
        return f"{subzona} ({round(float(lat), 4)}, {round(float(lon), 4)})"
    except (ValueError, TypeError):
        return f"{subzona} (Err Coords)"
