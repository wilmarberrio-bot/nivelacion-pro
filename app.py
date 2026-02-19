from flask import Flask, render_template, jsonify, send_file, request
import os
import sys
import uuid
import logging

# Configurar logs para verlos en Render
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Configuración global básica
Params = {
    'last_report': None
}

@app.route('/')
def index():
    logger.info("Cargando página principal...")
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error al renderizar index.html: {str(e)}")
        return f"Error interno: {str(e)}", 500

@app.route('/analyze', methods=['POST'])
def analyze():
    logger.info("Iniciando análisis...")
    try:
        # Importación tardía para evitar que la app no arranque si hay errores en el motor
        try:
            import sugerir_nivelacion
        except Exception as e:
            logger.error(f"Error fatal al importar motor de nivelación: {str(e)}")
            return jsonify({'status': 'error', 'message': f"Error al cargar el motor lógico: {str(e)}"})

        input_file = None
        
        # Procesar archivo subido
        if 'file' in request.files:
            file = request.files['file']
            if file.filename != '':
                input_file = f"temp_{uuid.uuid4().hex}.xlsx"
                file.save(input_file)
                logger.info(f"Archivo guardado temporalmente como: {input_file}")
        
        if not input_file:
            # Modo legado: buscar último archivo en carpeta
            input_file = sugerir_nivelacion.get_latest_preruta_file()
            
        if not input_file:
            return jsonify({'status': 'error', 'message': 'No se cargó ningún archivo.'})
        
        # Ejecutar lógica
        result = sugerir_nivelacion.generate_suggestions(input_file)
        
        # Manejar resultados
        if isinstance(result, tuple) and len(result) == 2:
            msg, output_path = result
        else:
            msg = result
            output_path = None
        
        # Limpieza
        if input_file.startswith("temp_") and os.path.exists(input_file):
            try:
                os.remove(input_file)
            except: pass

        if output_path:
            Params['last_report'] = output_path
            return jsonify({
                'status': 'success', 
                'message': msg,
                'file_available': True,
                'filename': os.path.basename(output_path)
            })
        else:
            return jsonify({'status': 'error', 'message': msg})
            
    except Exception as e:
        logger.error(f"Error crítico en analyze: {str(e)}")
        return jsonify({'status': 'error', 'message': f"Error inesperado: {str(e)}"})

@app.route('/download')
def download():
    if Params['last_report'] and os.path.exists(Params['last_report']):
        return send_file(Params['last_report'], as_attachment=True)
    return "Archivo no disponible", 404

# Endpoint de salud para Render
@app.route('/healthz')
def healthz():
    return "OK", 200

if __name__ == '__main__':
    # Configuración para ejecución local
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
