from flask import Flask, render_template, jsonify, send_file, request
import os
import sys
import threading
import webbrowser
import time
import shutil

# Import core logic
sys.path.append(os.getcwd())
import sugerir_nivelacion

app = Flask(__name__)

# Configuration
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
        
        # Check if file is uploaded
        if 'file' in request.files:
            file = request.files['file']
            if file.filename != '':
                # Save to unique temp file
                import uuid
                input_file = f"temp_{uuid.uuid4().hex}.xlsx"
                file.save(input_file)
        
        # If no file uploaded, try to find one in directory (legacy mode)
        if not input_file:
            input_file = sugerir_nivelacion.get_latest_preruta_file()
            
        if not input_file:
            return jsonify({'status': 'error', 'message': 'No se cargó ningún archivo.'})
        
        # Run logic
        result = sugerir_nivelacion.generate_suggestions(input_file)
        
        # Determine if it was successful based on return type
        if isinstance(result, tuple) and len(result) == 2:
            msg, output_path = result
        else:
            msg = result
            output_path = None
        
        # Cleanup temp file if it was an upload
        if input_file.startswith("temp_") and os.path.exists(input_file):
            try:
                os.remove(input_file)
            except: pass

        if output_path:
            Params['last_report'] = output_path
            # We could parse suggestions here if we want to display a table
            return jsonify({
                'status': 'success', 
                'message': msg,
                'file_available': True,
                'filename': os.path.basename(output_path)
            })
        else:
            return jsonify({'status': 'error', 'message': msg})
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/download')
def download():
    if Params['last_report'] and os.path.exists(Params['last_report']):
        return send_file(Params['last_report'], as_attachment=True)
    return "Archivo no disponible", 404

def open_browser():
    time.sleep(1.5)
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == '__main__':
    # Auto-open browser
    threading.Thread(target=open_browser).start()
    app.run(port=5000, debug=False)
