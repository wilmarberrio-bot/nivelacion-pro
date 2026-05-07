"""
routes/reports.py
Endpoints del apartado de reportes diarios.
"""
import logging
from flask import Blueprint, jsonify, request, Response
from services import snapshot_service as ss

logger = logging.getLogger(__name__)
reports_bp = Blueprint("reports", __name__, url_prefix="/api/reports")


def _get_orders():
    """Obtiene las ordenes del resultado vigente para tomar una foto manual."""
    try:
        from routes.api import _session_state
        result = (_session_state.get("last_result") or {})
        return list(result.get("ordenes_movibles") or []) + list(result.get("ordenes_bloqueadas") or [])
    except Exception as e:
        logger.warning("_get_orders error: %s", e)
        return []


@reports_bp.get("/snapshots")
def get_snapshots():
    try:
        hoy = ss._now_naive().strftime("%Y-%m-%d")
        fecha = request.args.get("fecha") or hoy
        cortes = ss.get_cortes(fecha)
        return jsonify({
            "status": "ok",
            "fecha": hoy,
            "fecha_solicitada": fecha,
            "cortes": cortes,
            "total": len(cortes),
            "fechas": ss.get_fechas(),
            "solo_hoy": True,
            "message": "El informe disponible corresponde solamente al dia actual.",
        })
    except Exception as e:
        logger.exception("Error en /api/reports/snapshots")
        return jsonify({"status": "error", "message": str(e)}), 500


@reports_bp.post("/snapshot")
def post_snapshot():
    try:
        orders = _get_orders()
        if not orders:
            return jsonify({"status": "error", "message": "No hay datos. Sube el Excel primero."}), 400
        body = request.get_json(silent=True) or {}
        corte = ss.registrar_corte(orders, body.get("label"))
        visible = {k: v for k, v in corte.items() if not k.startswith("_")}
        return jsonify({
            "status": "ok",
            "mensaje": f"Foto registrada a las {corte['hora']} - {corte['total']} ordenes",
            "corte": visible,
        })
    except Exception as e:
        logger.exception("Error en /api/reports/snapshot")
        return jsonify({"status": "error", "message": f"Error al tomar foto: {str(e)}"}), 500


@reports_bp.get("/export")
def get_export():
    try:
        hoy = ss._now_naive().strftime("%Y-%m-%d")
        fecha = request.args.get("fecha") or hoy
        if fecha != hoy:
            return jsonify({
                "status": "error",
                "message": "Ese informe ya vencio. Solo se puede descargar el reporte del dia actual.",
            }), 410
        cortes = ss.get_cortes(hoy)
        if not cortes:
            return jsonify({"status": "error", "message": f"Sin fotos/cortes registrados para hoy ({hoy})."}), 404
        data = ss.generar_excel(hoy)
        filename = f"reporte_seguimiento_fotografico_{hoy}.xlsx"
        return Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Cache-Control": "no-store, max-age=0",
            },
        )
    except Exception as e:
        logger.exception("Error generando Excel")
        return jsonify({"status": "error", "message": str(e)}), 500


@reports_bp.get("/fechas")
def get_fechas():
    return jsonify({"status": "ok", "fechas": ss.get_fechas(), "solo_hoy": True})
