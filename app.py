from flask import Flask, render_template, jsonify, send_file, request
import os
import sys
import uuid

# Import core logic
print("Iniciando importacion de sugerir_nivelacion...")
sys.path.append(os.getcwd())
try:
    import sugerir_nivelacion
    print("Importacion completada exitosamente.")
except Exception as e:
    print(f"ERROR CRITICO al importar sugerir_nivelacion: {str(e)}")
    raise

app = Flask(__name__)
print(f"Aplicacion Flask instanciada. PORT: {os.environ.get('PORT', '5000 (default)')}")

# Almacén temporal para el último reporte (solo para esta sesión)
Params = {
    'last_report': None
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        input_file = None
        
        # Procesar archivo subido (Cloud Mode)
        if 'file' in request.files:
            file = request.files['file']
            if file.filename != '':
                input_file = f"temp_{uuid.uuid4().hex}.xlsx"
                file.save(input_file)
        
        # Si no hay archivo subido, buscar el más reciente (Local Mode)
        if not input_file:
            input_file = sugerir_nivelacion.get_latest_preruta_file()
            
        if not input_file:
            return jsonify({'status': 'error', 'message': 'Por favor selecciona un archivo Excel primero.'})
        
        # Ejecutar el motor de nivelación
        result = sugerir_nivelacion.generate_suggestions(input_file)
        
        # Manejar la respuesta (Tupla: Mensaje, Path)
        if isinstance(result, tuple) and len(result) == 2:
            msg, output_path = result
        else:
            msg = result
            output_path = None
        
        # Eliminar archivo temporal de entrada
        if input_file.startswith("temp_") and os.path.exists(input_file):
            try:
                os.remove(input_file)
            except: pass

        if output_path:
            return jsonify({
                'status': 'success', 
                'message': msg,
                'file_available': True,
                'download_url': f"/download/{os.path.basename(output_path)}"
            })
        else:
            return jsonify({'status': 'error', 'message': msg})
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error interno: {str(e)}"})

@app.route('/download/<path:filename>')
def download(filename):
    # Por seguridad, solo permitir descargas en el directorio raiz
    if ".." in filename or "/" in filename or "\\" in filename:
        # Pero os.path.basename limpia eso si viene de la URL
        filename = os.path.basename(filename)
        
    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    return "Archivo no disponible. Intenta procesar de nuevo.", 404

if __name__ == '__main__':
    # Para pruebas locales
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)))
