import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import subprocess
import threading
import sys
# Import core logic
import sugerir_nivelacion

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(fill='both', expand=True, padx=10, pady=10)
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        self.lbl_title = tk.Label(self, text="Herramienta de Nivelación", font=("Helvetica", 16, "bold"))
        self.lbl_title.pack(pady=10)

        # File Selection Frame
        self.frm_file = tk.Frame(self)
        self.frm_file.pack(fill='x', pady=5)
        
        self.lbl_file = tk.Label(self.frm_file, text="Archivo:")
        self.lbl_file.pack(side='left')
        
        self.ent_file = tk.Entry(self.frm_file, width=50)
        self.ent_file.pack(side='left', padx=5, fill='x', expand=True)
        
        self.btn_browse = tk.Button(self.frm_file, text="...", command=self.browse_file)
        self.btn_browse.pack(side='left')

        # Auto Detect Button
        self.btn_detect = tk.Button(self, text="Detectar Último Archivo 'Pre-Ruta'", command=self.auto_detect)
        self.btn_detect.pack(pady=5)

        # Action Button
        self.btn_run = tk.Button(self, text="GENERAR SUGERENCIAS", bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"), command=self.run_process)
        self.btn_run.pack(pady=15, fill='x')

        # Output Log
        self.txt_log = scrolledtext.ScrolledText(self, height=10)
        self.txt_log.pack(fill='both', expand=True)

        # Open Result Button (Initially disabled)
        self.btn_open = tk.Button(self, text="Abrir Resultados", state='disabled', command=self.open_result)
        self.btn_open.pack(pady=10)
        
        self.last_output_file = None

        # Try auto-detect on startup
        self.auto_detect()

    def log(self, message):
        self.txt_log.insert(tk.END, message + "\n")
        self.txt_log.see(tk.END)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.ent_file.delete(0, tk.END)
            self.ent_file.insert(0, filename)

    def auto_detect(self):
        f = sugerir_nivelacion.get_latest_preruta_file()
        if f:
            abs_path = os.path.abspath(f)
            self.ent_file.delete(0, tk.END)
            self.ent_file.insert(0, abs_path)
            self.log(f"Auto-detectado: {os.path.basename(f)}")
        else:
            self.log("No se detectó automáticamente ningún archivo 'pre_ruta'.")

    def run_process(self):
        input_file = self.ent_file.get()
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Por favor seleccione un archivo válido.")
            return

        self.btn_run.config(state='disabled')
        self.log("Iniciando análisis...")
        
        # Run in thread to not freeze GUI
        thread = threading.Thread(target=self.process_thread, args=(input_file,))
        thread.start()

    def process_thread(self, input_file):
        try:
            msg, output_path = sugerir_nivelacion.generate_suggestions(input_file)
            
            # Update GUI from thread
            self.after(0, lambda: self.process_done(msg, output_path))
        except Exception as e:
            self.after(0, lambda: self.log(f"Error crítico: {e}"))
            self.after(0, lambda: self.btn_run.config(state='normal'))

    def process_done(self, msg, output_path):
        self.log(msg)
        self.btn_run.config(state='normal')
        
        if output_path:
            self.last_output_file = output_path
            self.btn_open.config(state='normal', bg="#2196F3", fg="white")
            messagebox.showinfo("Completado", "Proceso finalizado correctamente.")

    def open_result(self):
        if self.last_output_file:
            if sys.platform == 'win32':
                os.startfile(self.last_output_file)
            else:
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, self.last_output_file])

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Nivelación de Técnicos")
    root.geometry("600x500")
    app = Application(master=root)
    app.mainloop()
