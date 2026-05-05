"""
data_sources/excel_loader.py
Carga ordenes desde Excel. Columnas del export real Metabase #26359.
"""
import io, re, logging
logger = logging.getLogger(__name__)

_ALIASES = {
    "id":        ["appointment_id","id","id_orden","orden","ticket"],
    "tecnico":   ["tecnico","Technician","Tech","Nombre Tecnico"],
    "estado":    ["status_txt","Status","estado","Estado","appointment_status"],
    "franja":    ["franja_label","Slot","franja","Franja","Horario","time_slot"],
    "tipo":      ["appointment_type_txt","Category","tipo","Tipo","Type"],
    "zona":      ["Zone Name","zona","Zone","Cities__name","Cities_name","municipio"],
    "subzona":   ["Subzone","subzona","barrio","sector"],
    "direccion": ["Addresses__address","Addresses_address","direccion","address"],
    "gmaps":     ["Google Maps Link","gmaps","maps","link"],
    "lat":       ["Latitude","lat","latitud"],
    "lon":       ["Longitude","lon","longitud"],
    "updated_at":["onsite_at_cot","Updated At","updated_at","fecha"],
    "site":      ["Sites","site","edificio"],
}

def _find(hl, aliases):
    for a in aliases:
        try: return hl.index(a.lower())
        except ValueError: pass
    return -1

def _coords(text):
    if not text: return 0.0, 0.0
    for p in [r'@(-?\d{1,3}\.\d+),\s*(-?\d{1,3}\.\d+)',
              r'!3d(-?\d{1,3}\.\d+)!4d(-?\d{1,3}\.\d+)',
              r'(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)']:
        m = re.search(p, str(text))
        if m:
            try: return float(m.group(1)), float(m.group(2))
            except: pass
    return 0.0, 0.0

def load_from_bytes(file_bytes, filename=""):
    try: import openpyxl
    except ImportError: raise RuntimeError("openpyxl no instalado")
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)
    if not header_row: return []
    hl = [str(h).strip().lower() if h is not None else "" for h in header_row]
    col = {field: _find(hl, aliases) for field, aliases in _ALIASES.items()}
    found = {f: header_row[i] for f,i in col.items() if i>=0}
    logger.info(f"Excel columnas mapeadas: {found}")
    orders = []
    for row in rows_iter:
        if not any(v is not None for v in row): continue
        def g(field, default=""):
            i = col.get(field, -1)
            if i < 0 or i >= len(row): return default
            v = row[i]
            if v is None: return default
            s = str(v).strip()
            return "" if s.lower() == "none" else s
        def gf(field):
            try:
                i = col.get(field, -1)
                if i < 0 or i >= len(row): return 0.0
                v = row[i]
                return float(v) if v not in (None, "", "None") else 0.0
            except: return 0.0
        raw_id = g("id")
        if raw_id.endswith(".0"): raw_id = raw_id[:-2]
        order_id = raw_id or f"row_{len(orders)}"
        lat, lon = gf("lat"), gf("lon")
        if not lat and not lon: lat, lon = _coords(g("gmaps"))
        zona = g("zona", "").strip()
        if not zona or zona.lower() == "none": zona = "SIN_ZONA"
        else: zona = zona.upper()
        estado = g("estado", "por programar").lower()
        orders.append({
            "id":        order_id,
            "tecnico":   g("tecnico",   "SIN_ASIGNAR"),
            "estado":    estado,
            "franja":    g("franja",    "Sin Franja"),
            "tipo":      g("tipo",      "instalacion").lower(),
            "zona":      zona,
            "subzona":   g("subzona",   "SIN_SUBZONA").upper() or "SIN_SUBZONA",
            "site":      g("site",      ""),
            "direccion": g("direccion", ""),
            "gmaps":     g("gmaps",     ""),
            "lat":       lat,
            "lon":       lon,
            "updated_at":g("updated_at",""),
            "_source":   "excel",
        })
    wb.close()
    logger.info(f"Excel '{filename}': {len(orders)} ordenes")
    return orders
