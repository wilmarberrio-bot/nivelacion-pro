"""
data_sources/metabase_client.py - Cliente Metabase pregunta #26359
URL: https://metabase.somosinternet.com | Card ID: 26359
Parametros: Fecha, zona, cat | Deteccion automatica de columnas
"""
import os, time, logging, urllib.request, urllib.error
import json as _json
from config import (METABASE_URL,METABASE_USER,METABASE_PASSWORD,
    METABASE_CARD_ID,METABASE_API_KEY,DATA_CACHE_TTL)
logger = logging.getLogger(__name__)
METABASE_PARAM_ZONA=os.environ.get("METABASE_PARAM_ZONA","1")
METABASE_PARAM_CAT=os.environ.get("METABASE_PARAM_CAT","0")
COL_ORDER_ID=os.environ.get("COL_ORDER_ID","appointment_id")
COL_SITE=os.environ.get("COL_SITE","Sites")
COL_ZONE=os.environ.get("COL_ZONE","Zone Name")
COL_SUBZONE=os.environ.get("COL_SUBZONE","Subzone")
COL_ADDRESS=os.environ.get("COL_ADDRESS","Addresses_address")
COL_CITY=os.environ.get("COL_CITY","Cities_name")
COL_GMAPS=os.environ.get("COL_GMAPS","Google Maps Link")
COL_LAT=os.environ.get("COL_LAT","Latitude")
COL_LON=os.environ.get("COL_LON","Longitude")
COL_TECH_ALIASES=["Technician","tecnico","Tecnico","Tech","Nombre Tecnico","nombre_tecnico"]
COL_STATUS_ALIASES=["Status","estado","Estado","appointment_status","Estado Orden"]
COL_FRANJA_ALIASES=["Slot","franja","Franja","Horario","Franja Horaria","time_slot"]
COL_TIPO_ALIASES=["Category","categoria","Categoria","tipo","Tipo","Type","work_type"]
COL_UPDATED_ALIASES=["Updated At","updated_at","Fecha Actualizacion","last_update"]
_sc={"token":None,"expires_at":0}
_dc={"data":None,"fetched_at":0}
def _get_token():
    now=time.time()
    if _sc["token"] and now<_sc["expires_at"]: return _sc["token"]
    if METABASE_API_KEY:
        _sc.update({"token":METABASE_API_KEY,"expires_at":now+86400}); return METABASE_API_KEY
    if not all([METABASE_URL,METABASE_USER,METABASE_PASSWORD]):
        raise RuntimeError("Faltan METABASE_URL/USER/PASSWORD")
    req=urllib.request.Request(f"{METABASE_URL.rstrip('/')}/api/session",
        data=_json.dumps({"username":METABASE_USER,"password":METABASE_PASSWORD}).encode(),
        headers={"Content-Type":"application/json"},method="POST")
    try:
        with urllib.request.urlopen(req,timeout=15) as r:
            body=_json.loads(r.read().decode()); token=body.get("id")
            if not token: raise RuntimeError(f"Sin token: {body}")
            _sc.update({"token":token,"expires_at":now+3600}); return token
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Auth {e.code}: {e.read().decode()}")
def _headers():
    t=_get_token()
    return ({"x-api-key":t,"Content-Type":"application/json"} if METABASE_API_KEY
            else {"X-Metabase-Session":t,"Content-Type":"application/json"})
def _query_card(card_id,fecha,zona,cat):
    url=f"{METABASE_URL.rstrip('/')}/api/card/{card_id}/query/json"
    params=[{"type":"date/single","target":["variable",["template-tag","Fecha"]],"value":fecha},
            {"type":"number/=","target":["variable",["template-tag","zona"]],"value":zona},
            {"type":"number/=","target":["variable",["template-tag","cat"]],"value":cat}]
    req=urllib.request.Request(url,data=_json.dumps({"parameters":params}).encode(),
        headers=_headers(),method="POST")
    try:
        with urllib.request.urlopen(req,timeout=45) as r: raw=_json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Metabase {e.code}: {e.read().decode()[:200]}")
    if isinstance(raw,list): return raw
    data=raw.get("data",{})
    cols=[c.get("display_name") or c.get("name",f"col{i}") for i,c in enumerate(data.get("cols",[]))]
    logger.info(f"Columnas Metabase: {cols}")
    return [dict(zip(cols,row)) for row in data.get("rows",[])]
def _find_col(s,aliases):
    kl={k.lower():k for k in s.keys()}
    for a in aliases:
        if a in s: return a
        if a.lower() in kl: return kl[a.lower()]
    return None
def _build_col_map(rows):
    if not rows: return {}
    s=rows[0]
    return {"id":COL_ORDER_ID if COL_ORDER_ID in s else _find_col(s,["appointment_id","id","orden"]),
            "tecnico":_find_col(s,COL_TECH_ALIASES),"estado":_find_col(s,COL_STATUS_ALIASES),
            "franja":_find_col(s,COL_FRANJA_ALIASES),"tipo":_find_col(s,COL_TIPO_ALIASES),
            "zona":COL_ZONE if COL_ZONE in s else _find_col(s,["Zone Name","zona"]),
            "subzona":COL_SUBZONE if COL_SUBZONE in s else _find_col(s,["Subzone","subzona"]),
            "ciudad":COL_CITY if COL_CITY in s else _find_col(s,["Cities_name","ciudad"]),
            "site":COL_SITE if COL_SITE in s else _find_col(s,["Sites","site"]),
            "direccion":COL_ADDRESS if COL_ADDRESS in s else _find_col(s,["Addresses_address","direccion"]),
            "gmaps":COL_GMAPS if COL_GMAPS in s else _find_col(s,["Google Maps Link","gmaps"]),
            "lat":COL_LAT if COL_LAT in s else _find_col(s,["Latitude","lat"]),
            "lon":COL_LON if COL_LON in s else _find_col(s,["Longitude","lon"]),
            "updated_at":_find_col(s,COL_UPDATED_ALIASES)}
def _norm(row,cm):
    def g(k,d=""): v=row.get(k) if k else None; return str(v).strip() if v not in (None,"") else d
    def gf(k):
        try: v=row.get(k) if k else None; return float(v) if v not in (None,"","None") else 0.0
        except: return 0.0
    zona=g(cm.get("zona"),"SIN_ZONA").upper() or "SIN_ZONA"
    if zona=="SIN_ZONA": zona=g(cm.get("ciudad"),"SIN_ZONA").upper() or "SIN_ZONA"
    return {"id":g(cm.get("id"),f"sid_{id(row)}"),"tecnico":g(cm.get("tecnico"),"SIN_ASIGNAR"),
            "estado":g(cm.get("estado"),"por programar"),"franja":g(cm.get("franja"),"Sin Franja"),
            "tipo":g(cm.get("tipo"),"instalacion").lower(),"zona":zona,
            "subzona":g(cm.get("subzona"),"SIN_SUBZONA").upper() or "SIN_SUBZONA",
            "ciudad":g(cm.get("ciudad"),""),"site":g(cm.get("site"),""),
            "direccion":g(cm.get("direccion"),""),"gmaps":g(cm.get("gmaps"),""),
            "lat":gf(cm.get("lat")),"lon":gf(cm.get("lon")),
            "updated_at":g(cm.get("updated_at"),""),"_source":"metabase"}
def fetch_orders(force=False,fecha=None,zona=None,cat=None):
    from config import now_bogota
    now=time.time()
    if not force and _dc["data"] and (now-_dc["fetched_at"])<DATA_CACHE_TTL: return _dc["data"]
    if not METABASE_URL: return []
    fR=fecha or now_bogota().strftime("%Y-%m-%d")
    zR=str(zona or METABASE_PARAM_ZONA); cR=str(cat or METABASE_PARAM_CAT)
    try:
        rows=_query_card(METABASE_CARD_ID,fR,zR,cR)
        if not rows: return []
        cm=_build_col_map(rows)
        logger.info(f"Mapa columnas: { {k:v for k,v in cm.items() if v} }")
        norm=[_norm(r,cm) for r in rows]; _dc.update({"data":norm,"fetched_at":now})
        logger.info(f"{len(norm)} ordenes fecha={fR} zona={zR}"); return norm
    except Exception as e:
        logger.error(f"Error Metabase: {e}"); return _dc["data"] or []
def invalidate_cache(): _dc.update({"data":None,"fetched_at":0})
def cache_info():
    now=time.time(); age=now-_dc["fetched_at"] if _dc["fetched_at"] else None
    return {"has_data":bool(_dc["data"]),"rows":len(_dc["data"]) if _dc["data"] else 0,
            "age_seconds":round(age,1) if age else None,"ttl_seconds":DATA_CACHE_TTL,
            "fresh":age is not None and age<DATA_CACHE_TTL}
