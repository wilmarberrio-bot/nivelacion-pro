import os
import glob
import re

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
