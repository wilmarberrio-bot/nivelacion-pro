"""
routes/api.py
Todos los endpoints JSON de la plataforma de nivelación.
Sin generación de Excel como flujo principal.
"""
import logging
from flask import Blueprint, jsonify, request
from data_sources.metabase_client import fetch_orders, invalidate_cache, cache_info
from data_sources.excel_loader import load_from_bytes
from services.leveling_engine import run_leveling

logger = logging.getLogger(__name__)
api_bp = Blueprint("api", __name__, url_prefix="/api")

# Estado de sesión en memoria (historial de sugerencias aplicadas)
_session_state = {
    "applied":   [],   # sugerencias marcadas como aplicadas
    "dismissed": [],   # sugerencias descartadas
    "last_result": None,
}


def _get_result(force: bool = False):
    """Obtiene o recalcula el resultado de nivelación."""
    if force or _session_state["last_result"] is None:
        orders = fetch_orders(force=force)
        result = run_leveling(orders)
        _session_state["last_result"] = result
    return _session_state["last_result"]


# ─── /api/nivelacion ─────────────────────────

@api_bp.get("/nivelacion")
def get_nivelacion():
    """
    Devuelve el resultado completo de nivelación:
    resumen, carga, órdenes, alertas y sugerencias.
    """
    try:
        force = request.args.get("refresh", "false").lower() == "true"
        result = _get_result(force=force)
        return jsonify({"status": "ok", **result})
    except Exception as e:
        logger.exception("Error en /api/nivelacion")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/resumen ────────────────────────────

@api_bp.get("/resumen")
def get_resumen():
    """Devuelve solo el resumen ejecutivo."""
    try:
        result = _get_result()
        return jsonify({
            "status":      "ok",
            "generado_en": result.get("generado_en"),
            "resumen":     result.get("resumen", {}),
            "cache":       cache_info(),
        })
    except Exception as e:
        logger.exception("Error en /api/resumen")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/alertas ────────────────────────────

@api_bp.get("/alertas")
def get_alertas():
    """Devuelve la lista de alertas activas con severidad."""
    try:
        result = _get_result()
        alertas = result.get("alertas", [])
        severidad = request.args.get("severidad")  # critica | alta | media
        if severidad:
            alertas = [a for a in alertas if a.get("severidad") == severidad]
        return jsonify({
            "status":  "ok",
            "total":   len(alertas),
            "alertas": alertas,
        })
    except Exception as e:
        logger.exception("Error en /api/alertas")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/sugerencias ───────────────────────

@api_bp.get("/sugerencias")
def get_sugerencias():
    """Devuelve sugerencias de nivelación, filtradas por estado de sesión."""
    try:
        result = _get_result()
        sugerencias = result.get("sugerencias", [])

        # Excluir aplicadas/descartadas en esta sesión
        applied_ids   = {s["orden"] for s in _session_state["applied"]}
        dismissed_ids = {s["orden"] for s in _session_state["dismissed"]}

        activas = [
            {**s, "session_status": "aplicada" if s["orden"] in applied_ids
                                    else "descartada" if s["orden"] in dismissed_ids
                                    else "pendiente"}
            for s in sugerencias
        ]

        mostrar = request.args.get("estado", "pendiente")
        if mostrar != "todas":
            activas = [s for s in activas if s["session_status"] == mostrar]

        return jsonify({
            "status":     "ok",
            "total":      len(activas),
            "sugerencias": activas,
        })
    except Exception as e:
        logger.exception("Error en /api/sugerencias")
        return jsonify({"status": "error", "message": str(e)}), 500


@api_bp.post("/sugerencias/accion")
def post_sugerencia_accion():
    """
    Registra una acción sobre una sugerencia (aplicar/descartar/revertir).
    Solo persiste en sesión (memoria). No modifica Metabase.
    Body: {"orden": "id_orden", "accion": "aplicar"|"descartar"|"revertir"}
    """
    body = request.get_json(silent=True) or {}
    orden  = body.get("orden")
    accion = body.get("accion")

    if not orden or accion not in ("aplicar", "descartar", "revertir"):
        return jsonify({"status": "error", "message": "orden y accion requeridos (aplicar|descartar|revertir)"}), 400

    result = _get_result()
    sugerencia = next((s for s in result.get("sugerencias", []) if s["orden"] == orden), None)
    if not sugerencia:
        return jsonify({"status": "error", "message": f"No se encontró sugerencia para orden {orden}"}), 404

    if accion == "aplicar":
        _session_state["dismissed"] = [s for s in _session_state["dismissed"] if s["orden"] != orden]
        if not any(s["orden"] == orden for s in _session_state["applied"]):
            _session_state["applied"].append(sugerencia)
    elif accion == "descartar":
        _session_state["applied"] = [s for s in _session_state["applied"] if s["orden"] != orden]
        if not any(s["orden"] == orden for s in _session_state["dismissed"]):
            _session_state["dismissed"].append(sugerencia)
    elif accion == "revertir":
        _session_state["applied"]   = [s for s in _session_state["applied"]   if s["orden"] != orden]
        _session_state["dismissed"] = [s for s in _session_state["dismissed"] if s["orden"] != orden]

    return jsonify({
        "status":   "ok",
        "orden":    orden,
        "accion":   accion,
        "aplicadas":   len(_session_state["applied"]),
        "descartadas": len(_session_state["dismissed"]),
    })


# ─── /api/refresh ────────────────────────────

@api_bp.post("/refresh")
def post_refresh():
    """Fuerza recarga de datos desde Metabase e invalida el caché."""
    try:
        invalidate_cache()
        _session_state["last_result"] = None
        result = _get_result(force=True)
        return jsonify({
            "status":      "ok",
            "mensaje":     "Datos actualizados desde Metabase",
            "generado_en": result.get("generado_en"),
            "total_ordenes": result["resumen"]["total_ordenes"],
            "cache":       cache_info(),
        })
    except Exception as e:
        logger.exception("Error en /api/refresh")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/upload (Excel fallback) ───────────

@api_bp.post("/upload")
def post_upload():
    """
    Sube un archivo Excel como fuente de datos alternativa.
    Solo activo si Metabase no está configurado o falla.
    No guarda el archivo en disco permanentemente.
    """
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No se envió archivo"}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"status": "error", "message": "Solo se aceptan archivos .xlsx"}), 400

    try:
        file_bytes = f.read()  # Leer en memoria, sin guardar en disco
        orders = load_from_bytes(file_bytes, f.filename)
        result = run_leveling(orders)
        _session_state["last_result"] = result
        return jsonify({
            "status":        "ok",
            "fuente":        "excel",
            "archivo":       f.filename,
            "total_ordenes": result["resumen"]["total_ordenes"],
            "generado_en":   result.get("generado_en"),
        })
    except Exception as e:
        logger.exception("Error procesando Excel")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/cache ──────────────────────────────

@api_bp.get("/cache")
def get_cache():
    """Estado actual del caché de datos."""
    return jsonify({"status": "ok", "cache": cache_info()})


# ─── /api/export (secundario, no default) ───

@api_bp.get("/export/csv")
def export_csv():
    """
    Exporta órdenes movibles a CSV.
    Opción secundaria y explícita. No es el flujo principal.
    """
    import io
    import csv
    try:
        result = _get_result()
        orders = result.get("ordenes_movibles", []) + result.get("ordenes_bloqueadas", [])
        if not orders:
            return jsonify({"status": "error", "message": "Sin datos para exportar"}), 404

        buf = io.StringIO()
        fieldnames = ["id", "tecnico", "estado", "estado_clase", "franja", "tipo",
                      "zona", "subzona", "direccion", "movible", "updated_at"]
        writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(orders)
        csv_content = buf.getvalue().encode("utf-8-sig")

        from flask import Response
        return Response(
            csv_content,
            mimetype="text/csv",
            headers={"Content-Disposition": "attachment; filename=nivelacion.csv"},
        )
    except Exception as e:
        logger.exception("Error exportando CSV")
        return jsonify({"status": "error", "message": str(e)}), 500


# ─── /api/appointments/export ────────────────
# Usado por Google Sheets (nivelacion_informe_diario.gs)
# para sincronizar el estado de los appointments del día.

@api_bp.get("/appointments/export")
def get_appointments_export():
    """
    Exporta todos los appointments del día en formato plano
    compatible con el Apps Script de Google Sheets.
    Incluye movibles + bloqueados con campos de tracking.
    """
    try:
        result = _get_result()

        movibles   = result.get("ordenes_movibles",  [])
        bloqueadas = result.get("ordenes_bloqueadas", [])
        todas = movibles + bloqueadas

        appointments = []
        for o in todas:
            appointments.append({
                "appointment_id":      o.get("id", ""),
                "fecha_cita":          o.get("fecha_cita") or o.get("updated_at", "")[:10] if o.get("updated_at") else "",
                "estado":              o.get("estado", ""),
                "estado_anterior":     o.get("estado_anterior", ""),
                "tipo_cita":           o.get("tipo", ""),
                "tecnico":             o.get("tecnico", ""),
                "supervisor":          o.get("supervisor", ""),
                "zona":                o.get("zona", ""),
                "motivo_cancelacion":  o.get("motivo", ""),
                "fecha_original":      "",
                "fecha_nueva":         "",
                "hora_evento":         o.get("hora_evento", ""),
                "observacion":         o.get("subzona", ""),
            })

        return jsonify({
            "status":       "ok",
            "total":        len(appointments),
            "generado_en":  result.get("generado_en", ""),
            "appointments": appointments,
        })
    except Exception as e:
        logger.exception("Error en /api/appointments/export")
        return jsonify({"status": "error", "message": str(e)}), 500
