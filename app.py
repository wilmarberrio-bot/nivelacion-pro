"""
app.py - Nivelacion Pro Web
Flask app principal. Compatible con Render (gunicorn).
Sin generacion de Excel como flujo principal.
"""
import os, sys, logging
from flask import Flask, render_template, jsonify

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger(__name__)

sys.path.insert(0, os.path.dirname(__file__))

app = Flask(__name__)
app.config["JSON_ENSURE_ASCII"] = False

from routes.api import api_bp
app.register_blueprint(api_bp)

@app.route("/")
def index():
    return render_template("dashboard.html")

@app.route("/health")
def health():
    return "ok", 200

@app.route("/analyze", methods=["POST"])
def analyze_legacy():
    from flask import request as req
    from data_sources.metabase_client import fetch_orders
    from services.leveling_engine import run_leveling
    try:
        if "file" in req.files:
            from data_sources.excel_loader import load_from_bytes
            f = req.files["file"]
            orders = load_from_bytes(f.read(), f.filename)
        else:
            fecha = req.form.get("fecha")
            zona  = req.form.get("zona")
            cat   = req.form.get("cat")
            orders = fetch_orders(fecha=fecha, zona=zona, cat=cat)
        result = run_leveling(orders)
        return jsonify({"status":"ok","message":f"Nivela completada. {result['resumen']['total_ordenes']} ordenes procesadas.","data_url":"/api/nivelacion","dashboard":"/",**result})
    except Exception as e:
        logger.exception("Error en /analyze legacy")
        return jsonify({"status":"error","message":str(e)}), 500

@app.errorhandler(404)
def not_found(e):
    return jsonify({"status":"error","message":"Ruta no encontrada"}), 404

@app.errorhandler(500)
def internal_error(e):
    return jsonify({"status":"error","message":"Error interno"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV","production") == "development"
    logger.info(f"Iniciando Nivelacion Pro Web en puerto {port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
