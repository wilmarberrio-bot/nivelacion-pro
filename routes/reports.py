"""
routes/reports.py
Endpoints para el sistema de reportes evolutivos diarios.
"""
import logging
from flask import Blueprint, jsonify, request, Response
from services import snapshot_service as ss

logger = logging.getLogger(__name__)
reports_bp = Blueprint("reports", __name__, url_prefix="/api/reports")


def _get_orders():
    """Obtiene todas las ordenes del resultado en sesion de forma robusta."""
    try:
        from routes.api import _session_state
        result = (_session_state.get("last_result") or {})
        return list(result.get("ordenes_movibles") or []) + list(result.get("ordenes_bloqueadas") or [])
    except Exception as e:
        logger.warning(f"_get_orders error: {e}")
        return []


@reports_bp.get("/snapshots")
def get_snapshots():
    try:
        fecha = request.args.get("fecha")
        cortes = ss.get_cortes(fecha)
        return jsonify({"status":"ok","fecha":fecha or ss._now_naive().strftime("%Y-%m-%d"),
                        "cortes":cortes,"total":len(cortes),"fechas":ss.get_fechas()})
    except Exception as e:
        logger.exception("Error en /api/reports/snapshots")
        return jsonify({"status":"error","message":str(e)}),500


@reports_bp.post("/snapshot")
def post_snapshot():
    try:
        orders = _get_orders()
        if not orders:
            return jsonify({"status":"error","message":"No hay datos. Sube el Excel primero."}),400
        body = request.get_json(silent=True) or {}
        corte = ss.registrar_corte(orders, body.get("label"))
        return jsonify({"status":"ok","mensaje":f"Corte '{corte['label']}' registrado - {corte['total']} ordenes",
                        "corte":{k:v for k,v in corte.items() if not k.startswith("_")}})
    except Exception as e:
        logger.exception("Error en /api/reports/snapshot")
        return jsonify({"status":"error","message":f"Error al tomar corte: {str(e)}"}),500


@reports_bp.get("/export")
def get_export():
    try:
        from config import now_bogota
        fecha = request.args.get("fecha") or now_bogota().strftime("%Y-%m-%d")
        cortes = ss.get_cortes(fecha)
        if not cortes:
            return jsonify({"status":"error","message":f"Sin cortes para {fecha}."}),404
        data = ss.generar_excel(fecha)
        filename = f"reporte_nivelacion_{fecha}.xlsx"
        return Response(data,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition":f"attachment; filename={filename}"})
    except Exception as e:
        logger.exception("Error generando Excel")
        return jsonify({"status":"error","message":str(e)}),500


@reports_bp.get("/fechas")
def get_fechas():
    return jsonify({"status":"ok","fechas":ss.get_fechas()})
