from flask import Flask, render_template, jsonify, send_file, request
import os
import sys
import uuid
import importlib
import traceback

# Asegurar path del proyecto
sys.path.append(os.getcwd())

app = Flask(__name__)
print(f"Aplicacion Flask instanciada. PORT: {os.environ.get('PORT', '5000 (default)')}")
print(f"Directorio actual: {os.getcwd()}")

# Health check
@app.get('/health')
def health():
    return 'ok', 200

# Almacen temporal para el ultimo reporte
Params = {
    'last_report': None
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    input_file = None
    try:
        # Lazy import: solo carga el motor cuando se usa /analyze
        print('Cargando modulo sugerir_nivelacion bajo demanda...')
        sugerir_nivelacion = importlib.import_module('sugerir_nivelacion')
        print('Modulo sugerir_nivelacion cargado OK.')
        try:
            print('Archivo del modulo sugerir_nivelacion:', sugerir_nivelacion.__file__)
        except Exception:
            pass

        # Procesar archivo subido (Cloud Mode)
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename:
                input_file = f"temp_{uuid.uuid4().hex}.xlsx"
                file.save(input_file)
                print(f"Archivo temporal guardado: {input_file}")

        # Si no hay archivo subido, buscar el mas reciente (Local Mode)
        if not input_file:
            print('No se subio archivo. Buscando archivo mas reciente...')
            input_file = sugerir_nivelacion.get_latest_preruta_file()

        if not input_file:
            return jsonify({
                'status': 'error',
                'message': 'Por favor selecciona un archivo Excel primero.'
            })

        print(f"Ejecutando motor con archivo: {input_file}")
        result = sugerir_nivelacion.generate_suggestions(input_file)
        print(f"Resultado del motor: {type(result)}")

        # Manejar la respuesta (Tupla: Mensaje, Path)
        if isinstance(result, tuple) and len(result) == 2:
            msg, output_path = result
        else:
            msg = result
            output_path = None

        # Eliminar archivo temporal de entrada
        if input_file and input_file.startswith('temp_') and os.path.exists(input_file):
            try:
                os.remove(input_file)
                print(f"Archivo temporal eliminado: {input_file}")
            except Exception as cleanup_error:
                print(f"No se pudo eliminar archivo temporal {input_file}: {cleanup_error}")

        if output_path:
            Params['last_report'] = output_path
            print(f"Reporte generado correctamente: {output_path}")
            return jsonify({
                'status': 'success',
                'message': msg,
                'file_available': True,
                'download_url': f"/download/{os.path.basename(output_path)}"
            })

        print(f"Motor devolvio mensaje sin archivo: {msg}")
        return jsonify({
            'status': 'error',
            'message': msg
        })

    except Exception as e:
        print('=== ERROR EN /analyze ===')
        print(traceback.format_exc())

        # Intentar limpiar temporal si existe
        if input_file and isinstance(input_file, str) and input_file.startswith('temp_') and os.path.exists(input_file):
            try:
                os.remove(input_file)
                print(f"Archivo temporal eliminado tras error: {input_file}")
            except Exception as cleanup_error:
                print(f"No se pudo eliminar archivo temporal tras error {input_file}: {cleanup_error}")

        return jsonify({
            'status': 'error',
            'message': f"Error interno: {str(e)}"
        })

@app.route('/download/<path:filename>')
def download(filename):
    # Por seguridad, solo permitir descargas en el directorio raiz
    if '..' in filename or '/' in filename or '\\' in filename:
        filename = os.path.basename(filename)

    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)

    return 'Archivo no disponible. Intenta procesar de nuevo.', 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
