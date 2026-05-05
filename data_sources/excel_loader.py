"""data_sources/excel_loader.py - Carga ordenes desde Excel como fallback"""
import io, re, logging
logger = logging.getLogger(__name__)
_ALIASES = {
    "id":["appointment_id","id","id_orden","orden","ticket"],
    "tecnico":["Technician","tecnico","Tecnico","tech_name"],
    "estado":["Status","estado","Estado","appointment_status"],
    "franja":["Slot","franja","Franja","Horario","time_slot"],
    "tipo":["Category","categoria","tipo","Tipo","Type"],
    "zona":["Zone Name","zona","Zone","municipio"],
    "subzona":["Subzone","subzona","barrio","sector"],
    "direccion":["Addresses_address","direccion","address"],
    "gmaps":["Google Maps Link","gmaps","maps","link"],
    "lat":["Latitude","lat","latitud"],
    "lon":["Longitude","lon","longitud"],
    "updated_at":["Updated At","updated_at","fecha"],
}
def _find(headers_lower, aliases):
    for a in aliases:
        try: return headers_lower.index(a.lower())
        except ValueError: pass
    return -1
def _coords(text):
    if not text: return 0.0, 0.0
    for p in [r'@(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)',
              r'[?&]q=(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)',
              r'!3d(-?\d{1,3}\.\d+)!4d(-?\d{1,3}\.\d+)',
              r'(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)']:
        m = re.search(p, str(text))
        if m:
            try: return float(m.group(1)), float(m.group(2))
            except: pass
    return 0.0, 0.0
def load_from_bytes(file_bytes: bytes, filename: str = "") -> list:
    try: import openpyxl
    except ImportError: raise RuntimeError("openpyxl no instalado")
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)
    if not header_row: return []
    hl = [str(h).strip().lower() if h else "" for h in header_row]
    col = {f: _find(hl, aliases) for f, aliases in _ALIASES.items()}
    orders = []
    for row in rows_iter:
        if not any(row): continue
        def g(f, d=""):
            i = col.get(f, -1)
            if i < 0 or i >= len(row): return d
            v = row[i]; return str(v).strip() if v is not None else d
        def gf(f):
            try:
                i = col.get(f, -1)
                return float(row[i]) if i >= 0 and i < len(row) and row[i] not in (None,"") else 0.0
            except: return 0.0
        lat, lon = gf("lat"), gf("lon")
        if not lat and not lon: lat, lon = _coords(g("gmaps"))
        orders.append({"id":g("id",f"row_{len(orders)}"),"tecnico":g("tecnico","SIN_ASIGNAR"),
            "estado":g("estado","por programar"),"franja":g("franja","Sin Franja"),
            "tipo":g("tipo","instalacion").lower(),"zona":g("zona","SIN_ZONA").upper() or "SIN_ZONA",
            "subzona":g("subzona","SIN_SUBZONA").upper() or "SIN_SUBZONA",
            "direccion":g("direccion"),"gmaps":g("gmaps"),"lat":lat,"lon":lon,
            "updated_at":g("updated_at"),"_source":"excel"})
    wb.close()
    logger.info(f"Excel: {len(orders)} ordenes desde '{filename}'")
    return orders
