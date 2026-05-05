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
    """Obtiene todas las ordenes del resultado en sesion."""
    from routes.api import _session_state
    result = _session_state.get("last_result") or {}
    return (result.get("ordenes_movibles") or []) + (result.get("ordenes_bloqueadas") or [])


@reports_bp.get("/snapshots")
def get_snapshots():
    """Devuelve todos los cortes del dia con sus datos."""
    fecha  = request.args.get("fecha")
    cortes = ss.get_cortes(fecha)
    return jsonify({
        "status":   "ok",
        "fecha":    fecha or ss._now_naive().strftime("%Y-%m-%d"),
        "cortes":   cortes,
        "total":    len(cortes),
        "fechas":   ss.get_fechas(),
    })


@reports_bp.post("/snapshot")
def post_snapshot():
    """Toma un corte manual del estado actual."""
    orders = _get_orders()
    if not orders:
        return jsonify({"status": "error", "message": "No hay datos. Sube el Excel primero."}), 400

    body  = request.get_json(silent=True) or {}
    label = body.get("label")
    corte = ss.registrar_corte(orders, label)
    return jsonify({
        "status":  "ok",
        "mensaje": f"Corte '{corte['label']}' registrado - {corte['total']} ordenes",
        "corte":   corte,
    })


@reports_bp.get("/export")
def get_export():
    """Genera y descarga el Excel del reporte evolutivo del dia."""
    from config import now_bogota
    fecha = request.args.get("fecha") or now_bogota().strftime("%Y-%m-%d")

    cortes = ss.get_cortes(fecha)
    if not cortes:
        return jsonify({
            "status":  "error",
            "message": f"Sin cortes para {fecha}. Sube el Excel para registrar el primer corte.",
        }), 404

    try:
        data     = ss.generar_excel(fecha)
        filename = f"reporte_nivelacion_{fecha}.xlsx"
        return Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except Exception as e:
        logger.exception("Error generando Excel de reportes")
        return jsonify({"status": "error", "message": str(e)}), 500


@reports_bp.get("/fechas")
def get_fechas():
    return jsonify({"status": "ok", "fechas": ss.get_fechas()})
