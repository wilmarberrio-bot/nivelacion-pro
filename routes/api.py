"""
routes/api.py
Todos los endpoints JSON de la plataforma de nivelacion.
Sin generacion de Excel como flujo principal.
"""
import logging
from flask import Blueprint, jsonify, request
from data_sources.metabase_client import fetch_orders, invalidate_cache, cache_info
from data_sources.excel_loader import load_from_bytes
from services.leveling_engine import run_leveling

logger = logging.getLogger(__name__)
api_bp = Blueprint("api", __name__, url_prefix="/api")

_session_state = {
    "applied":   [],
    "dismissed": [],
    "last_result": None,
}


def _get_result(force=False, fecha=None, zona=None, cat=None):
    if force or _session_state["last_result"] is None:
        orders = fetch_orders(force=force, fecha=fecha, zona=zona, cat=cat)
        result = run_leveling(orders)
        _session_state["last_result"] = result
    return _session_state["last_result"]


@api_bp.get("/nivelacion")
def get_nivelacion():
    try:
        force = request.args.get("refresh","false").lower()=="true"
        fecha = request.args.get("fecha")
        zona  = request.args.get("zona")
        cat   = request.args.get("cat")
        if fecha or zona or cat:
            force = True
        result = _get_result(force=force, fecha=fecha, zona=zona, cat=cat)
        try:
            from routes.blacklist import filter_suggestions, get_blacklist
            bl = get_blacklist()
            if bl and result.get("sugerencias"):
                result = dict(result)
                result["sugerencias"] = filter_suggestions(result["sugerencias"])
                result["blacklist_activa"] = bl
        except Exception:
            pass
        return jsonify({"status":"ok",**result})
    except Exception as e:
        logger.exception("Error en /api/nivelacion")
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.get("/resumen")
def get_resumen():
    try:
        result = _get_result()
        return jsonify({"status":"ok","generado_en":result.get("generado_en"),
                        "resumen":result.get("resumen",{}),"cache":cache_info()})
    except Exception as e:
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.get("/alertas")
def get_alertas():
    try:
        result = _get_result()
        alertas = result.get("alertas",[])
        sev = request.args.get("severidad")
        if sev: alertas = [a for a in alertas if a.get("severidad")==sev]
        return jsonify({"status":"ok","total":len(alertas),"alertas":alertas})
    except Exception as e:
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.get("/sugerencias")
def get_sugerencias():
    try:
        result = _get_result()
        sugs = result.get("sugerencias",[])
        applied  = {s["orden"] for s in _session_state["applied"]}
        dismissed = {s["orden"] for s in _session_state["dismissed"]}
        activas = [{**s,"session_status":"aplicada" if s["orden"] in applied else "descartada" if s["orden"] in dismissed else "pendiente"} for s in sugs]
        mostrar = request.args.get("estado","pendiente")
        if mostrar!="todas": activas=[s for s in activas if s["session_status"]==mostrar]
        return jsonify({"status":"ok","total":len(activas),"sugerencias":activas})
    except Exception as e:
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.post("/sugerencias/accion")
def post_sugerencia_accion():
    body = request.get_json(silent=True) or {}
    orden = str(body.get("orden") or "").strip()
    accion = body.get("accion")
    if not orden or accion not in ("aplicar","descartar","revertir"):
        return jsonify({"status":"error","message":"orden y accion requeridos"}),400
    result = _get_result()
    sug = next((s for s in result.get("sugerencias",[]) if str(s.get("orden"))==orden),None)
    if not sug: return jsonify({"status":"error","message":f"No encontrada: {orden}"}),404
    if accion=="aplicar":
        _session_state["dismissed"]=[s for s in _session_state["dismissed"] if str(s.get("orden"))!=orden]
        if not any(str(s.get("orden"))==orden for s in _session_state["applied"]): _session_state["applied"].append(sug)
    elif accion=="descartar":
        _session_state["applied"]=[s for s in _session_state["applied"] if str(s.get("orden"))!=orden]
        if not any(str(s.get("orden"))==orden for s in _session_state["dismissed"]): _session_state["dismissed"].append(sug)
    else:
        _session_state["applied"]=[s for s in _session_state["applied"] if str(s.get("orden"))!=orden]
        _session_state["dismissed"]=[s for s in _session_state["dismissed"] if str(s.get("orden"))!=orden]
    return jsonify({"status":"ok","orden":orden,"accion":accion,
                    "aplicadas":len(_session_state["applied"]),"descartadas":len(_session_state["dismissed"])})


@api_bp.post("/refresh")
def post_refresh():
    try:
        body = request.get_json(silent=True) or {}
        fecha = body.get("fecha") or request.args.get("fecha")
        zona  = body.get("zona")  or request.args.get("zona")
        cat   = body.get("cat")   or request.args.get("cat")
        invalidate_cache()
        _session_state["last_result"] = None
        result = _get_result(force=True, fecha=fecha, zona=zona, cat=cat)
        return jsonify({"status":"ok","mensaje":"Datos actualizados",
                        "generado_en":result.get("generado_en"),
                        "total_ordenes":result["resumen"]["total_ordenes"],"cache":cache_info()})
    except Exception as e:
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.post("/upload")
def post_upload():
    if "file" not in request.files:
        return jsonify({"status":"error","message":"No se envio archivo"}),400
    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"status":"error","message":"Solo se aceptan .xlsx"}),400
    try:
        orders = load_from_bytes(f.read(), f.filename)
        result = run_leveling(orders)
        _session_state["last_result"] = result
        snap_info = None
        try:
            from services.snapshot_service import registrar_corte
            corte = registrar_corte(orders)
            snap_info = {"label":corte["label"],"hora":corte["hora"],"total":corte["total"]}
            logger.info(f"Corte automatico '{corte['label']}': {corte['total']} ordenes")
        except Exception as se:
            logger.warning(f"Corte fallo (no critico): {se}")
        return jsonify({"status":"ok","fuente":"excel","archivo":f.filename,
                        "total_ordenes":result["resumen"]["total_ordenes"],
                        "generado_en":result.get("generado_en"),"snapshot":snap_info})
    except Exception as e:
        logger.exception("Error procesando Excel")
        return jsonify({"status":"error","message":str(e)}),500


@api_bp.get("/cache")
def get_cache():
    return jsonify({"status":"ok","cache":cache_info()})
