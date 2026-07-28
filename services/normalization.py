"""services/normalization.py - Normalizacion de estados, franjas, zonas y coordenadas"""
import re, math
from config import (MOVABLE_STATUSES,BLOCKED_STATUSES,NEAR_FINISH_STATUSES,
    FINALIZED_STATUSES,STATUS_PROGRESS,FRANJAS,NEARBY_BUILDING_RADIUS_KM)

def norm_text(x,default=""): return str(x).strip() if x is not None else default
def norm_upper(x,default="SIN_VALOR"):
    s=norm_text(x,default)
    return default if not s or s.lower() in ("none","nan","") else s.upper()
def norm_zone(x): return norm_upper(x,"SIN_ZONA")
def norm_subzone(x): return norm_upper(x,"SIN_SUBZONA")
def norm_status(x):
    s=norm_text(x,"por programar").lower()
    for a,b in [("a","a"),("e","e"),("i","i"),("o","o"),("u","u"),
                ("\xe1","a"),("\xe9","e"),("\xed","i"),("\xf3","o"),("\xfa","u")]:
        s=s.replace(a,b)
    return s.strip()
def norm_franja(x):
    s=norm_text(x,"").replace("\u2013","-").replace("\u2014","-")
    if not s or s.lower() in ("none","nan",""): return "Sin Franja"
    clean=s.replace(" ","")
    for f in FRANJAS:
        if f.replace(" ","")==clean: return f
    return s
def parse_franja_hours(franja_str):
    if not franja_str or franja_str=="Sin Franja": return None,None
    try:
        clean=str(franja_str).replace("\u2013","-").replace("\u2014","-")
        parts=clean.split("-")
        if len(parts)<2: return None,None
        def pt(t):
            t=t.strip()
            for seg in t.split():
                if ":" in seg: h,m=seg.split(":")[:2]; return int(h)+int(m)/60.0
            if ":" in t: h,m=t.split(":")[:2]; return int(h)+int(m)/60.0
            return None
        s,e=pt(parts[0]),pt(parts[1])
        return (s,e) if s is not None and e is not None else (None,None)
    except: return None,None
def classify_status(estado):
    s=norm_status(estado)
    if any(m in s for m in MOVABLE_STATUSES): return "movible"
    if any(f in s for f in FINALIZED_STATUSES): return "finalizado"
    if any(n in s for n in NEAR_FINISH_STATUSES): return "avanzado"
    if any(b in s for b in BLOCKED_STATUSES): return "bloqueado"
    return "desconocido"
def is_movable(estado): return classify_status(estado)=="movible"
def is_blocked(estado): return classify_status(estado) in ("bloqueado","avanzado","finalizado")
def get_status_progress(estado):
    s=norm_status(estado)
    for k,v in STATUS_PROGRESS.items():
        if k in s: return v
    return 0
def status_effective_weight(estado):
    s=norm_status(estado)
    if "cancelad" in s: return 0.35
    if any(k in s for k in ["finaliz","por auditar","cerrad","completad"]): return 0.05
    if any(k in s for k in NEAR_FINISH_STATUSES): return 0.65
    if get_status_progress(s)>=1: return 1.25
    if any(k in s for k in MOVABLE_STATUSES): return 1.05
    return 0.95
def status_completion_credit(estado):
    s=norm_status(estado)
    if any(k in s for k in ["finaliz","por auditar","cerrad","completad"]): return 1.0
    if "cancelad" in s: return 0.25
    return 0.0
def haversine(lat1,lon1,lat2,lon2):
    R=6371.0; dlat=math.radians(lat2-lat1); dlon=math.radians(lon2-lon1)
    a=math.sin(dlat/2)**2+math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dlon/2)**2
    return R*2*math.atan2(math.sqrt(a),math.sqrt(1-a))
def get_centroid(locs):
    if not locs: return(0.0,0.0)
    lats=[l[0] for l in locs if l[0]]; lons=[l[1] for l in locs if l[1]]
    return(sum(lats)/len(lats),sum(lons)/len(lons)) if lats else(0.0,0.0)
def normalize_address_key(addr):
    if not addr: return ""
    a=addr.strip().upper()
    for ch in [",",".","#","-","_"]: a=a.replace(ch," ")
    return " ".join(a.split())[:80]
def build_address_key(address,subzona):
    b=normalize_address_key(address) if address else ""
    return b if b else (subzona or "SIN_SUBZONA").upper()[:60]
def order_has_coords(order): return bool(order.get("lat",0) and order.get("lon",0))
def is_same_unit(o1,o2):
    ak1=o1.get("addr_key",""); ak2=o2.get("addr_key","")
    if ak1 and ak1==ak2: return True
    if order_has_coords(o1) and order_has_coords(o2):
        try: return haversine(o1["lat"],o1["lon"],o2["lat"],o2["lon"])<=NEARBY_BUILDING_RADIUS_KM
        except: pass
    return False
def normalize_order(order):
    o=dict(order)
    o["tecnico"]=norm_text(o.get("tecnico"),"SIN_ASIGNAR") or "SIN_ASIGNAR"
    o["estado"]=norm_text(o.get("estado"),"por programar")
    o["franja"]=norm_franja(o.get("franja"))
    # Zona con fallback a Cities__name cuando está vacía
    zona_raw = o.get("zona", "") or ""
    ciudad_raw = o.get("ciudad", "") or ""
    if not zona_raw or str(zona_raw).strip().upper() in ("SIN_ZONA", "NONE", "NAN", ""):
        o["zona"] = norm_zone(ciudad_raw) if ciudad_raw else "SIN_ZONA"
    else:
        o["zona"] = norm_zone(zona_raw)
    o["subzona"]=norm_subzone(o.get("subzona"))
    o["tipo"]=norm_text(o.get("tipo"),"instalacion").lower()
    o["lat"]=float(o.get("lat") or 0); o["lon"]=float(o.get("lon") or 0)
    o["estado_clase"]=classify_status(o["estado"])
    o["progress"]=get_status_progress(o["estado"])
    o["effective_weight"]=status_effective_weight(o["estado"])
    o["completion_credit"]=status_completion_credit(o["estado"])
    o["addr_key"]=build_address_key(o.get("direccion",""),o["subzona"])
    o["movible"]=o["estado_clase"]=="movible"
    return o

# ── Turno detection ──────────────────────────────────────────────────────────
import csv, os

_TURNOS_CACHE: dict = {}
_TURNOS_LOADED: bool = False

def _strip_accents(s: str) -> str:
    """Quita tildes para comparacion robusta: ANDRES == ANDRES."""
    return (s.replace("A\u0301","A").replace("E\u0301","E")
             .replace("\xc1","A").replace("\xc9","E").replace("\xcd","I")
             .replace("\xd3","O").replace("\xda","U").replace("\xd1","N")
             .replace("\xe1","a").replace("\xe9","e").replace("\xed","i")
             .replace("\xf3","o").replace("\xfa","u").replace("\xf1","n"))

def _load_turnos_csv() -> dict:
    """Carga data/turnos.csv una sola vez. Retorna {tecnico_upper: 'T1'|'T2'}."""
    global _TURNOS_CACHE, _TURNOS_LOADED
    if _TURNOS_LOADED:
        return _TURNOS_CACHE
    try:
        from config import TURNOS_CSV_PATH
        csv_path = TURNOS_CSV_PATH
    except ImportError:
        csv_path = os.path.join(os.path.dirname(__file__), "..", "data", "turnos.csv")
    result = {}
    try:
        with open(csv_path, newline="", encoding="utf-8") as fh:
            reader = csv.DictReader(row for row in fh if not row.strip().startswith("#"))
            for row in reader:
                tech = row.get("tecnico", "").strip().upper()
                turno = row.get("turno", "").strip().upper()
                if tech and turno in ("T1", "T2"):
                    result[_strip_accents(tech)] = turno
    except Exception:
        pass
    _TURNOS_CACHE = result
    _TURNOS_LOADED = True
    return result

def reload_turnos_csv() -> dict:
    """Fuerza recarga del CSV (útil en desarrollo / tests)."""
    global _TURNOS_LOADED
    _TURNOS_LOADED = False
    return _load_turnos_csv()

def detect_turno(tech: str, tech_orders_list: list) -> str:
    """
    Determina el turno de un técnico.
    Prioridad:
      1. Lista maestra data/turnos.csv
      2. Auto-detección desde franjas de sus órdenes:
         - T1 si tiene ≥1 orden en 08:00-09:30
         - T2 si tiene ≥1 orden en 16:00-17:30
         - T2 si TODAS sus órdenes son ≥ 10:00
         - T1 por defecto
    Nota: si un técnico tiene órdenes de 08:00 Y de 16:00 (caso muy raro),
    se prioriza T1 — no debería ocurrir operativamente.
    """
    # 1. Lookup en CSV
    roster = _load_turnos_csv()
    tech_key = _strip_accents((tech or "").strip().upper())
    if tech_key in roster:
        return roster[tech_key]
    # Búsqueda parcial (por si el nombre en Metabase es más largo)
    for key, turno in roster.items():
        if key and (key in tech_key or tech_key in key):
            return turno

    # 2. Auto-detección desde órdenes
    if not tech_orders_list:
        return "T1"
    franjas = [
        o.get("franja", "")
        for o in tech_orders_list
        if o.get("franja") and o.get("franja") != "Sin Franja"
    ]
    if not franjas:
        return "T1"
    # Regla explícita: orden en 08:00 → T1
    if any("08:00" in f for f in franjas):
        return "T1"
    # Regla explícita: orden en 16:00 → T2
    if any("16:00" in f for f in franjas):
        return "T2"
    # Si todas las órdenes empiezan a las 10:00 o más tarde → T2
    starts = []
    for f in franjas:
        h, _ = parse_franja_hours(f)
        if h is not None:
            starts.append(h)
    if starts and min(starts) >= 10.0:
        return "T2"
    return "T1"
