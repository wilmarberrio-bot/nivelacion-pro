"""routes/api.py - Endpoints JSON de la plataforma de nivelacion. Sin Excel como flujo principal."""
import io, csv, logging
from flask import Blueprint, jsonify, request, Response
from data_sources.metabase_client import fetch_orders, invalidate_cache, cache_info
from data_sources.excel_loader import load_from_bytes
from services.leveling_engine import run_leveling
logger=logging.getLogger(__name__)
api_bp=Blueprint("api",__name__,url_prefix="/api")
_ss={"applied":[],"dismissed":[],"last_result":None}
def _get(force=False,fecha=None,zona=None,cat=None):
    if force or _ss["last_result"] is None:
        _ss["last_result"]=run_leveling(fetch_orders(force=force,fecha=fecha,zona=zona,cat=cat))
    return _ss["last_result"]
@api_bp.get("/nivelacion")
def get_nivelacion():
    try:
        force=request.args.get("refresh","false").lower()=="true"
        fecha=request.args.get("fecha"); zona=request.args.get("zona"); cat=request.args.get("cat")
        if fecha or zona or cat: force=True
        return jsonify({"status":"ok",**_get(force=force,fecha=fecha,zona=zona,cat=cat)})
    except Exception as e:
        logger.exception("Error /api/nivelacion"); return jsonify({"status":"error","message":str(e)}),500
@api_bp.get("/resumen")
def get_resumen():
    try:
        r=_get()
        return jsonify({"status":"ok","generado_en":r.get("generado_en"),"resumen":r.get("resumen",{}),"cache":cache_info()})
    except Exception as e: return jsonify({"status":"error","message":str(e)}),500
@api_bp.get("/alertas")
def get_alertas():
    try:
        r=_get(); alertas=r.get("alertas",[])
        sev=request.args.get("severidad")
        if sev: alertas=[a for a in alertas if a.get("severidad")==sev]
        return jsonify({"status":"ok","total":len(alertas),"alertas":alertas})
    except Exception as e: return jsonify({"status":"error","message":str(e)}),500
@api_bp.get("/sugerencias")
def get_sugerencias():
    try:
        r=_get(); sugs=r.get("sugerencias",[])
        ap={s["orden"] for s in _ss["applied"]}; di={s["orden"] for s in _ss["dismissed"]}
        sugs=[{**s,"session_status":"aplicada" if s["orden"] in ap else "descartada" if s["orden"] in di else "pendiente"} for s in sugs]
        m=request.args.get("estado","pendiente")
        if m!="todas": sugs=[s for s in sugs if s["session_status"]==m]
        rf=request.args.get("riesgo")
        if rf: sugs=[s for s in sugs if s.get("riesgo")==rf]
        return jsonify({"status":"ok","total":len(sugs),"sugerencias":sugs})
    except Exception as e: return jsonify({"status":"error","message":str(e)}),500
@api_bp.post("/sugerencias/accion")
def post_accion():
    body=request.get_json(silent=True) or {}
    orden=body.get("orden"); accion=body.get("accion")
    if not orden or accion not in ("aplicar","descartar","revertir"):
        return jsonify({"status":"error","message":"orden y accion requeridos"}),400
    r=_get(); sug=next((s for s in r.get("sugerencias",[]) if s["orden"]==orden),None)
    if not sug: return jsonify({"status":"error","message":f"No hay sugerencia para {orden}"}),404
    if accion=="aplicar":
        _ss["dismissed"]=[s for s in _ss["dismissed"] if s["orden"]!=orden]
        if not any(s["orden"]==orden for s in _ss["applied"]): _ss["applied"].append(sug)
    elif accion=="descartar":
        _ss["applied"]=[s for s in _ss["applied"] if s["orden"]!=orden]
        if not any(s["orden"]==orden for s in _ss["dismissed"]): _ss["dismissed"].append(sug)
    else:
        _ss["applied"]=[s for s in _ss["applied"] if s["orden"]!=orden]
        _ss["dismissed"]=[s for s in _ss["dismissed"] if s["orden"]!=orden]
    return jsonify({"status":"ok","orden":orden,"accion":accion,"aplicadas":len(_ss["applied"]),"descartadas":len(_ss["dismissed"])})
@api_bp.post("/refresh")
def post_refresh():
    try:
        body=request.get_json(silent=True) or {}
        fecha=body.get("fecha") or request.args.get("fecha")
        zona=body.get("zona") or request.args.get("zona")
        cat=body.get("cat") or request.args.get("cat")
        invalidate_cache(); _ss["last_result"]=None
        r=_get(force=True,fecha=fecha,zona=zona,cat=cat)
        return jsonify({"status":"ok","mensaje":"Datos actualizados","generado_en":r.get("generado_en"),
                        "total_ordenes":r["resumen"]["total_ordenes"],"cache":cache_info()})
    except Exception as e: logger.exception("Error /api/refresh"); return jsonify({"status":"error","message":str(e)}),500
@api_bp.post("/upload")
def post_upload():
    if "file" not in request.files: return jsonify({"status":"error","message":"No se envio archivo"}),400
    f=request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"status":"error","message":"Solo .xlsx"}),400
    try:
        orders=load_from_bytes(f.read(),f.filename); r=run_leveling(orders); _ss["last_result"]=r
        return jsonify({"status":"ok","fuente":"excel","archivo":f.filename,"total_ordenes":r["resumen"]["total_ordenes"],"generado_en":r.get("generado_en")})
    except Exception as e: logger.exception("Error upload"); return jsonify({"status":"error","message":str(e)}),500
@api_bp.get("/cache")
def get_cache(): return jsonify({"status":"ok","cache":cache_info()})
@api_bp.get("/export/csv")
def export_csv():
    try:
        r=_get(); orders=r.get("ordenes_movibles",[])+r.get("ordenes_bloqueadas",[])
        if not orders: return jsonify({"status":"error","message":"Sin datos"}),404
        buf=io.StringIO()
        writer=csv.DictWriter(buf,fieldnames=["id","tecnico","estado","estado_clase","franja","tipo","zona","subzona","direccion","movible","updated_at"],extrasaction="ignore")
        writer.writeheader(); writer.writerows(orders)
        return Response(buf.getvalue().encode("utf-8-sig"),mimetype="text/csv",
                        headers={"Content-Disposition":"attachment; filename=nivelacion.csv"})
    except Exception as e: return jsonify({"status":"error","message":str(e)}),500
