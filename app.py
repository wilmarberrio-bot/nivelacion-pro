from flask import Flask, render_template, jsonify, send_file, request
import os
import sys
import uuid
import importlib
import traceback

sys.path.append(os.getcwd())

app = Flask(__name__)
print(f"Aplicacion Flask instanciada. PORT: {os.environ.get('PORT', '5000 (default)')}")

@app.get("/health")
def health():
    return "ok", 200

Params = {
    'last_report': None
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        print("Cargando modulo sugerir_nivelacion bajo demanda...")
        sugerir_nivelacion = importlib.import_module("sugerir_nivelacion")
        print("Modulo sugerir_nivelacion cargado OK.")

        input_file = None

        if 'file' in request.files:
            file = request.files['file']
            if file.filename != '':
                input_file = f"temp_{uuid.uuid4().hex}.xlsx"
                file.save(input_file)

        if not input_file:
            input_file = sugerir_nivelacion.get_latest_preruta_file()

        if not input_file:
            return jsonify({'status': 'error', 'message': 'Por favor selecciona un archivo Excel primero.'})

        result = sugerir_nivelacion.generate_suggestions(input_file)

        if isinstance(result, tuple) and len(result) == 2:
            msg, output_path = result
        else:
            msg = result
            output_path = None

        if input_file.startswith("temp_") and os.path.exists(input_file):
            try:
                os.remove(input_file)
            except Exception:
                pass

        if output_path:
            return jsonify({
                'status': 'success',
                'message': msg,
                'file_available': True,
                'download_url': f"/download/{os.path.basename(output_path)}"
            })

        return jsonify({'status': 'error', 'message': msg})

    except Exception as e:
        print("=== ERROR EN /analyze ===")
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'message': f"Error interno: {str(e)}"})

@app.route('/download/<path:filename>')
def download(filename):
    if ".." in filename or "/" in filename or "\\" in filename:
        filename = os.path.basename(filename)

    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    return "Archivo no disponible. Intenta procesar de nuevo.", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)))
